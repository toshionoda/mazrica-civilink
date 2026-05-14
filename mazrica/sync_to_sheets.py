"""
Mazrica → Google Sheets 同期スクリプト

Mazricaから案件一覧（商品内訳付き）を取得し、
Google スプレッドシートに同期する
"""
import sys
import logging
from datetime import datetime
from typing import Optional

from .config import Config, load_dotenv
from .mazrica_client import MazricaClient, Deal, MazricaAPIError
from .google_sheets_client import GoogleSheetsClient, GoogleSheetsError


# ロギング設定
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


# スプレッドシートのヘッダー定義（案件一覧_v2 / 2026-04-22〜）
# 注: 列18(R)エリア / 列19(S)契約方法 はデプロイ済 Apps Script が案件名から自動抽出・
# 強制上書きするため、Python側は空欄のまま送る。履行開始日/終了日はその後ろに配置。
HEADERS = [
    "案件ID",
    "取引先名",
    "案件名",
    "担当者名",
    "フェーズ",
    "商品",
    "契約金額",
    "契約予定日",
    "案件発生日",
    "基本機能_オーナー",
    "基本機能_エンジニア",
    "基本機能_サポート",
    "照査AI",
    "サービス",
    "プリセールス担当者",
    "無料トライアル開始日",
    "無料トライアル終了日",
    "エリア",       # Apps Scriptが列18を上書き（案件名2番目の要素）
    "契約方法",     # Apps Scriptが列19を上書き（無料トライアル/サブスクリプション/有償化クローズ）
    "履行開始日",
    "履行終了日",
    "請求日",
    "ユーザー数",  # 案件名から抽出
    "期間",        # 案件名から抽出
    "分野",        # dealCustoms[75529] single_select
    "特記事項",    # dealCustoms[92214] text
    "更新日時",
]

# dealCustoms の管理番号マップ（取得時にこのIDで引く）
CUSTOM_ITEM_IDS = {
    "service": 80085,              # サービス
    "presales_owner": 92215,       # プリセールス担当者
    "trial_start": 99770,          # 無料トライアル開始日（CiviLink）
    "trial_end": 99771,            # 無料トライアル終了日（CiviLink）
    "delivery_start": 41081,       # 履行開始日
    "delivery_end": 34237,         # 履行終了日
    "billing_date": 34243,         # 請求日
    "field": 75529,                # 分野（single_select）
    "remarks": 92214,              # 特記事項（text）
}


import re

def extract_users_and_period(deal_name: str) -> tuple[str, str]:
    """
    案件名からユーザー数と期間を抽出
    
    例: "CiviLink_会社名_部署_機能_無料トライアル_10ユーザー_3カ月"
    → ("10", "3カ月")
    
    Args:
        deal_name: 案件名
    
    Returns:
        (ユーザー数, 期間) のタプル
    """
    users = ""
    period = ""
    
    # ユーザー数を抽出（例: "10ユーザー", "Xユーザー"）
    users_match = re.search(r'(\d+|X)ユーザー', deal_name)
    if users_match:
        users = users_match.group(1)
    
    # 期間を抽出（例: "3カ月", "Xカ月", "12ヶ月"）
    period_match = re.search(r'(\d+|X)(カ月|ヶ月|か月)', deal_name)
    if period_match:
        period = period_match.group(1) + period_match.group(2)
    
    return users, period


def _date_part(dt: str) -> str:
    """ISO形式の日時文字列から日付部分を取り出す（'2026-02-19T15:58:46+09:00' → '2026-02-19'）"""
    if not dt:
        return ""
    return dt.split("T", 1)[0]


def _bucket_product_detail(name: str) -> Optional[str]:
    """商品内訳の名前から集計バケットを判定。いずれにも該当しなければNone"""
    if not name:
        return None
    if "AI" in name or "配筋図" in name:
        return "ai"
    if "サポート" in name:
        return "support"
    if "オーナー" in name:
        return "owner"
    if "エンジニア" in name:
        return "engineer"
    return None


def _aggregate_product_buckets(deal: Deal) -> dict:
    """商品内訳を4バケット(owner/engineer/support/ai)に集計。数量を合算"""
    buckets = {"owner": 0.0, "engineer": 0.0, "support": 0.0, "ai": 0.0}
    matched_any = {"owner": False, "engineer": False, "support": False, "ai": False}
    for pd in deal.product_details:
        bucket = _bucket_product_detail(pd.product_name)
        if bucket is None:
            continue
        qty = pd.quantity if pd.quantity is not None else 0.0
        buckets[bucket] += qty
        matched_any[bucket] = True
    # 未マッチのバケットは空文字、マッチありは数値（整数なら整数で、小数なら小数で）
    result = {}
    for k, v in buckets.items():
        if not matched_any[k]:
            result[k] = ""
        elif v == int(v):
            result[k] = int(v)
        else:
            result[k] = v
    return result


def deal_to_rows(deal: Deal) -> list[list]:
    """
    案件データをスプレッドシートの行に変換（1案件1行）
    商品内訳は4バケット（基本機能_オーナー/エンジニア/サポート/照査AI）に集計
    """
    # 案件名からユーザー数・期間を抽出（既存ロジック）
    users, period = extract_users_and_period(deal.name)

    # 商品内訳を4バケットに集計
    buckets = _aggregate_product_buckets(deal)

    customs_by_id = deal.customs_by_id or {}

    row = [
        deal.id,
        deal.customer_name or "",
        deal.name,
        deal.user_name or "",
        deal.phase_name or "",
        deal.product_name or "",
        deal.amount if deal.amount is not None else "",
        deal.expected_contract_date or "",
        _date_part(deal.created_at),
        buckets["owner"],
        buckets["engineer"],
        buckets["support"],
        buckets["ai"],
        customs_by_id.get(CUSTOM_ITEM_IDS["service"], ""),
        customs_by_id.get(CUSTOM_ITEM_IDS["presales_owner"], ""),
        customs_by_id.get(CUSTOM_ITEM_IDS["trial_start"], ""),
        customs_by_id.get(CUSTOM_ITEM_IDS["trial_end"], ""),
        "",  # 列18: エリア（Apps Scriptが案件名から自動抽出して上書き）
        "",  # 列19: 契約方法（同上）
        customs_by_id.get(CUSTOM_ITEM_IDS["delivery_start"], ""),
        customs_by_id.get(CUSTOM_ITEM_IDS["delivery_end"], ""),
        customs_by_id.get(CUSTOM_ITEM_IDS["billing_date"], ""),
        users,
        period,
        customs_by_id.get(CUSTOM_ITEM_IDS["field"], ""),    # 列25 Y: 分野
        customs_by_id.get(CUSTOM_ITEM_IDS["remarks"], ""),  # 列26 Z: 特記事項
        deal.updated_at,
    ]
    return [row]


def filter_deal(deal: Deal, deal_name_filter: str, phase_name_filters: list[str]) -> bool:
    """
    案件がフィルタ条件に一致するかチェック

    Args:
        deal: 案件データ
        deal_name_filter: 案件名フィルタ（部分一致、大文字小文字無視、空の場合はフィルタなし）
        phase_name_filters: フェーズ名フィルタのリスト（いずれかに完全一致、空リストの場合はフィルタなし）

    Returns:
        条件に一致する場合はTrue
    """
    # フェーズ名フィルタ（リスト内のいずれかに一致すればOK）
    if phase_name_filters:
        if deal.phase_name not in phase_name_filters:
            return False

    # 案件名フィルタ（deal.name に部分一致、大文字小文字無視）
    if deal_name_filter:
        if not deal.name:
            return False
        if deal_name_filter.lower() not in deal.name.lower():
            return False

    return True


def _sync_one_spreadsheet(
    sheets: GoogleSheetsClient,
    sheet_name: str,
    deals: list,
    spreadsheet_id: Optional[str] = None,
) -> dict:
    """単一スプレッドシートに対して差分同期を実行"""
    label = f"spreadsheet_id={spreadsheet_id}" if spreadsheet_id else "active spreadsheet"
    logger.info(f"[{label}] Getting existing IDs from sheet '{sheet_name}'...")
    existing_ids = sheets.get_existing_ids(
        sheet_name, id_column=1, spreadsheet_id=spreadsheet_id
    )
    existing_id_set = set(str(id) for id in existing_ids)

    mazrica_id_set = set(str(deal.id) for deal in deals)
    new_deals = [d for d in deals if str(d.id) not in existing_id_set]
    update_deals = [d for d in deals if str(d.id) in existing_id_set]
    delete_ids = [id for id in existing_ids if str(id) not in mazrica_id_set]

    new_rows = []
    for deal in new_deals:
        new_rows.extend(deal_to_rows(deal))

    update_rows = []
    for deal in update_deals:
        update_rows.extend(deal_to_rows(deal))

    logger.info(
        f"[{label}] Adding {len(new_rows)} rows, "
        f"Updating {len(update_rows)} rows, "
        f"Deleting {len(delete_ids)} rows"
    )
    result = sheets.sync_data(
        sheet_name=sheet_name,
        headers=HEADERS,
        new_rows=new_rows,
        update_rows=update_rows,
        delete_ids=delete_ids,
        id_column=1,
        spreadsheet_id=spreadsheet_id,
    )
    logger.info(f"[{label}] Sync result: {result.get('message')}")

    return {
        "target": label,
        "existing_ids": len(existing_ids),
        "new_rows": len(new_rows),
        "updated_rows": len(update_rows),
        "deleted_rows": len(delete_ids),
        "message": result.get("message"),
    }


def sync_deals_to_sheets(
    deal_type_id: Optional[int] = None,
    sheet_name: Optional[str] = None,
    deal_name_filter: Optional[str] = None,
    phase_name_filters: Optional[list[str]] = None,
    additional_spreadsheet_ids: Optional[list[str]] = None,
) -> dict:
    """
    Mazricaの案件一覧をGoogle スプレッドシートに同期

    Args:
        deal_type_id: 同期する案件タイプID（Noneの場合は全案件）
        sheet_name: 出力先シート名
        deal_name_filter: 案件名フィルタ（部分一致、大文字小文字無視）
        phase_name_filters: フェーズ名フィルタのリスト（いずれかに完全一致）
        additional_spreadsheet_ids: 追加書き込み先のスプレッドシートIDリスト

    Returns:
        同期結果の統計情報（per_targetに各シートの結果）
    """
    sheet_name = sheet_name or Config.SHEET_NAME
    deal_type_id = deal_type_id or Config.DEAL_TYPE_ID
    deal_name_filter = deal_name_filter if deal_name_filter is not None else Config.FILTER_DEAL_NAME
    phase_name_filters = phase_name_filters if phase_name_filters is not None else Config.get_phase_name_list()
    if additional_spreadsheet_ids is None:
        additional_spreadsheet_ids = Config.get_additional_spreadsheet_ids()

    stats = {
        "total_deals": 0,
        "filtered_deals": 0,
        "per_target": [],
        "synced_at": datetime.now().isoformat(),
        "success": False,
        "error": None,
    }

    try:
        logger.info("Starting differential sync process...")
        logger.info(f"Filters: deal_name='{deal_name_filter}', phase_names={phase_name_filters}")
        logger.info(f"Targets: active + {additional_spreadsheet_ids}")

        # Mazricaクライアント初期化
        logger.info("Initializing Mazrica client...")
        mazrica = MazricaClient()

        # Google Sheetsクライアント初期化
        logger.info("Initializing Google Sheets client...")
        sheets = GoogleSheetsClient()

        # 案件データ取得（全ターゲットで共有）
        logger.info(f"Fetching deals from Mazrica (deal_type_id={deal_type_id})...")
        all_deals = mazrica.fetch_deals_with_products(deal_type_id=deal_type_id)
        stats["total_deals"] = len(all_deals)
        logger.info(f"Fetched {len(all_deals)} deals")

        # フィルタリング
        if deal_name_filter or phase_name_filters:
            deals = [d for d in all_deals if filter_deal(d, deal_name_filter, phase_name_filters)]
            logger.info(f"After filtering: {len(deals)} deals")
        else:
            deals = all_deals
        stats["filtered_deals"] = len(deals)

        # active spreadsheet（従来動作）
        stats["per_target"].append(
            _sync_one_spreadsheet(sheets, sheet_name, deals, spreadsheet_id=None)
        )

        # 追加スプレッドシート
        for sid in additional_spreadsheet_ids:
            stats["per_target"].append(
                _sync_one_spreadsheet(sheets, sheet_name, deals, spreadsheet_id=sid)
            )

        stats["success"] = True
        logger.info("Differential sync completed successfully!")
        
    except MazricaAPIError as e:
        stats["error"] = f"Mazrica API Error: {e.message}"
        logger.error(f"Mazrica API error: {e}")
    except GoogleSheetsError as e:
        stats["error"] = f"Google Sheets Error: {str(e)}"
        logger.error(f"Google Sheets error: {e}")
    except Exception as e:
        stats["error"] = f"Unexpected error: {str(e)}"
        logger.error(f"Unexpected error: {e}", exc_info=True)
    
    return stats


def main():
    """メインエントリーポイント"""
    # ローカル開発時は.envファイルから環境変数を読み込む
    load_dotenv()
    
    # 設定の検証
    errors = Config.validate()
    if errors:
        for error in errors:
            logger.error(error)
        sys.exit(1)
    
    # 同期実行
    stats = sync_deals_to_sheets()
    
    # 結果出力
    logger.info("=== Sync Statistics ===")
    logger.info(f"Total deals fetched: {stats['total_deals']}")
    logger.info(f"Deals after filter: {stats['filtered_deals']}")
    for t in stats.get("per_target", []):
        logger.info(
            f"  [{t['target']}] existing={t['existing_ids']} "
            f"new={t['new_rows']} updated={t.get('updated_rows', 0)} "
            f"deleted={t['deleted_rows']}"
        )
    logger.info(f"Synced at: {stats['synced_at']}")
    logger.info(f"Success: {stats['success']}")
    
    if stats["error"]:
        logger.error(f"Error: {stats['error']}")
        sys.exit(1)
    
    sys.exit(0)


if __name__ == "__main__":
    main()

