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


# スプレッドシートのヘッダー定義
HEADERS = [
    "案件ID",
    "案件名",
    "取引先",
    "取引先ID",
    "案件タイプ",
    "フェーズ",
    "担当者",
    "商品名",
    "数量",
    "単価",
    "商品金額",
    "案件金額",
    "受注予定日",
    "作成日時",
    "更新日時",
]


def deal_to_rows(deal: Deal) -> list[list]:
    """
    案件データをスプレッドシートの行に変換
    商品内訳がある場合は商品ごとに行を作成
    """
    rows = []
    
    # 共通データ
    base_data = [
        deal.id,
        deal.name,
        deal.customer_name or "",
        deal.customer_id or "",
        deal.deal_type_name or "",
        deal.phase_name or "",
        deal.user_name or "",
    ]
    
    # 商品名は deal.product_name から取得
    product_name = deal.product_name or ""
    
    if deal.product_details:
        # 商品内訳詳細がある場合は商品ごとに行を作成
        for pd in deal.product_details:
            row = base_data + [
                pd.product_name or product_name,  # 商品内訳の商品名、なければ案件の商品名
                pd.quantity if pd.quantity is not None else "",
                pd.unit_price if pd.unit_price is not None else "",
                pd.amount if pd.amount is not None else "",
                deal.amount if deal.amount is not None else "",
                deal.expected_contract_date or "",
                deal.created_at,
                deal.updated_at,
            ]
            rows.append(row)
    else:
        # 商品内訳がない場合は1行のみ（商品名は deal.product_name を使用）
        row = base_data + [
            product_name,  # 商品名
            "",  # 数量
            "",  # 単価
            "",  # 商品金額
            deal.amount if deal.amount is not None else "",
            deal.expected_contract_date or "",
            deal.created_at,
            deal.updated_at,
        ]
        rows.append(row)
    
    return rows


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
        "total_rows": 0,
        "synced_at": datetime.now().isoformat(),
        "success": False,
        "error": None
    }
    
    try:
        logger.info("Starting sync process...")
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
        
        # スプレッドシート用データ作成
        logger.info("Converting deals to spreadsheet format...")
        data = [HEADERS]
        for deal in deals:
            rows = deal_to_rows(deal)
            data.extend(rows)
        
        stats["total_rows"] = len(data) - 1  # ヘッダー行を除く
        logger.info(f"Total rows to write: {stats['total_rows']}")
        
        # スプレッドシートに書き込み
        logger.info(f"Writing data to sheet '{sheet_name}'...")
        sheets.write_data(sheet_name, data)
        
        # 書式設定
        logger.info("Applying formatting...")
        sheets.format_header_row(sheet_name)
        sheets.auto_resize_columns(sheet_name, len(HEADERS))
        
        stats["success"] = True
        logger.info("Sync completed successfully!")
        
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
    logger.info(f"Total rows: {stats['total_rows']}")
    logger.info(f"Synced at: {stats['synced_at']}")
    logger.info(f"Success: {stats['success']}")
    
    if stats["error"]:
        logger.error(f"Error: {stats['error']}")
        sys.exit(1)
    
    sys.exit(0)


if __name__ == "__main__":
    main()

