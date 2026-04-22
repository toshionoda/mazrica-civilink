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
    "履行開始日",
    "履行終了日",
    "請求日",
    "ユーザー数",  # 案件名から抽出
    "期間",        # 案件名から抽出
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
        customs_by_id.get(CUSTOM_ITEM_IDS["delivery_start"], ""),
        customs_by_id.get(CUSTOM_ITEM_IDS["delivery_end"], ""),
        customs_by_id.get(CUSTOM_ITEM_IDS["billing_date"], ""),
        users,
        period,
        deal.updated_at,
    ]
    return [row]


def filter_deal(deal: Deal, product_name_filter: str, phase_name_filters: list[str]) -> bool:
    """
    案件がフィルタ条件に一致するかチェック
    
    Args:
        deal: 案件データ
        product_name_filter: 商品名フィルタ（部分一致、空の場合はフィルタなし）
        phase_name_filters: フェーズ名フィルタのリスト（いずれかに完全一致、空リストの場合はフィルタなし）
    
    Returns:
        条件に一致する場合はTrue
    """
    # フェーズ名フィルタ（リスト内のいずれかに一致すればOK）
    if phase_name_filters:
        if deal.phase_name not in phase_name_filters:
            return False
    
    # 商品名フィルタ（deal.product_name を使用）
    if product_name_filter:
        product_name_lower = product_name_filter.lower()
        
        # まず案件の商品名（product.name）をチェック
        if deal.product_name and product_name_lower in deal.product_name.lower():
            return True
        
        # 商品内訳詳細もチェック
        for pd in deal.product_details:
            if pd.product_name and product_name_lower in pd.product_name.lower():
                return True
        
        return False
    
    return True


def sync_deals_to_sheets(
    deal_type_id: Optional[int] = None,
    sheet_name: Optional[str] = None,
    product_name_filter: Optional[str] = None,
    phase_name_filters: Optional[list[str]] = None
) -> dict:
    """
    Mazricaの案件一覧をGoogle スプレッドシートに同期
    
    Args:
        deal_type_id: 同期する案件タイプID（Noneの場合は全案件）
        sheet_name: 出力先シート名
        product_name_filter: 商品名フィルタ（部分一致）
        phase_name_filters: フェーズ名フィルタのリスト（いずれかに完全一致）
    
    Returns:
        同期結果の統計情報
    """
    sheet_name = sheet_name or Config.SHEET_NAME
    deal_type_id = deal_type_id or Config.DEAL_TYPE_ID
    product_name_filter = product_name_filter if product_name_filter is not None else Config.FILTER_PRODUCT_NAME
    phase_name_filters = phase_name_filters if phase_name_filters is not None else Config.get_phase_name_list()
    
    stats = {
        "total_deals": 0,
        "filtered_deals": 0,
        "existing_ids": 0,
        "new_rows": 0,
        "deleted_rows": 0,
        "skipped_rows": 0,
        "synced_at": datetime.now().isoformat(),
        "success": False,
        "error": None
    }
    
    try:
        logger.info("Starting differential sync process...")
        logger.info(f"Filters: product_name='{product_name_filter}', phase_names={phase_name_filters}")
        
        # Mazricaクライアント初期化
        logger.info("Initializing Mazrica client...")
        mazrica = MazricaClient()
        
        # Google Sheetsクライアント初期化
        logger.info("Initializing Google Sheets client...")
        sheets = GoogleSheetsClient()
        
        # 案件データ取得
        logger.info(f"Fetching deals from Mazrica (deal_type_id={deal_type_id})...")
        all_deals = mazrica.fetch_deals_with_products(deal_type_id=deal_type_id)
        stats["total_deals"] = len(all_deals)
        logger.info(f"Fetched {len(all_deals)} deals")
        
        # フィルタリング
        if product_name_filter or phase_name_filters:
            deals = [d for d in all_deals if filter_deal(d, product_name_filter, phase_name_filters)]
            stats["filtered_deals"] = len(deals)
            logger.info(f"After filtering: {len(deals)} deals")
        else:
            deals = all_deals
            stats["filtered_deals"] = len(deals)
        
        # 既存の案件IDを取得
        logger.info(f"Getting existing IDs from sheet '{sheet_name}'...")
        existing_ids = sheets.get_existing_ids(sheet_name, id_column=1)
        existing_id_set = set(str(id) for id in existing_ids)
        stats["existing_ids"] = len(existing_ids)
        logger.info(f"Found {len(existing_ids)} existing IDs in sheet")
        
        # Mazricaの案件IDセットを作成
        mazrica_id_set = set(str(deal.id) for deal in deals)
        
        # 新規案件を判定（Mazricaにあり、シートにない）
        new_deals = [d for d in deals if str(d.id) not in existing_id_set]
        logger.info(f"New deals to add: {len(new_deals)}")
        
        # 削除対象を判定（シートにあり、Mazricaにない）
        delete_ids = [id for id in existing_ids if str(id) not in mazrica_id_set]
        logger.info(f"Deals to delete: {len(delete_ids)}")
        
        # スキップ対象を計算
        stats["skipped_rows"] = len(existing_id_set) - len(delete_ids)
        
        # 新規行のデータ作成
        new_rows = []
        for deal in new_deals:
            rows = deal_to_rows(deal)
            new_rows.extend(rows)
        
        stats["new_rows"] = len(new_rows)
        stats["deleted_rows"] = len(delete_ids)
        
        # 差分同期を実行
        logger.info(f"Syncing to sheet '{sheet_name}'...")
        logger.info(f"  Adding {len(new_rows)} rows, Deleting {len(delete_ids)} rows")
        
        result = sheets.sync_data(
            sheet_name=sheet_name,
            headers=HEADERS,
            new_rows=new_rows,
            delete_ids=delete_ids,
            id_column=1
        )
        
        logger.info(f"Sync result: {result.get('message')}")
        
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
    logger.info(f"Existing IDs in sheet: {stats['existing_ids']}")
    logger.info(f"New rows added: {stats['new_rows']}")
    logger.info(f"Rows deleted: {stats['deleted_rows']}")
    logger.info(f"Rows skipped: {stats['skipped_rows']}")
    logger.info(f"Synced at: {stats['synced_at']}")
    logger.info(f"Success: {stats['success']}")
    
    if stats["error"]:
        logger.error(f"Error: {stats['error']}")
        sys.exit(1)
    
    sys.exit(0)


if __name__ == "__main__":
    main()

