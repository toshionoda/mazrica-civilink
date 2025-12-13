"""
Mazrica → Google Sheets 同期機能の設定
"""
import os
from typing import Optional


class Config:
    """環境変数から設定を読み込むクラス"""
    
    # Mazrica API設定
    MAZRICA_API_KEY: str = os.environ.get("MAZRICA_API_KEY", "")
    MAZRICA_BASE_URL: str = "https://senses-open-api.mazrica.com/v1"
    
    # Google Sheets設定（Apps Script連携）
    APPS_SCRIPT_URL: str = os.environ.get("APPS_SCRIPT_URL", "")
    APPS_SCRIPT_SECRET: str = os.environ.get("APPS_SCRIPT_SECRET", "")
    SHEET_NAME: str = os.environ.get("SHEET_NAME", "案件一覧")
    
    # 同期設定
    DEAL_TYPE_ID: Optional[int] = (
        int(os.environ.get("DEAL_TYPE_ID")) 
        if os.environ.get("DEAL_TYPE_ID") 
        else None
    )
    
    # フィルタ設定
    FILTER_PRODUCT_NAME: str = os.environ.get("FILTER_PRODUCT_NAME", "civilink")  # 商品名フィルタ
    FILTER_PHASE_NAMES: str = os.environ.get("FILTER_PHASE_NAMES", "受注,内示")  # フェーズフィルタ（カンマ区切りで複数指定可）
    
    @classmethod
    def get_phase_name_list(cls) -> list[str]:
        """フェーズ名フィルタをリストとして取得"""
        if not cls.FILTER_PHASE_NAMES:
            return []
        return [p.strip() for p in cls.FILTER_PHASE_NAMES.split(",") if p.strip()]
    
    # API制限
    API_RATE_LIMIT: float = 0.34  # 3リクエスト/秒 = 0.33秒間隔
    API_PAGE_SIZE: int = 100  # 1ページあたりの取得件数
    
    @classmethod
    def validate(cls) -> list[str]:
        """設定の検証を行い、エラーメッセージのリストを返す"""
        errors = []
        
        if not cls.MAZRICA_API_KEY:
            errors.append("MAZRICA_API_KEY が設定されていません")
        
        if not cls.APPS_SCRIPT_URL:
            errors.append("APPS_SCRIPT_URL が設定されていません")
        
        return errors


# ローカル開発用: .envファイルから読み込み
def load_dotenv():
    """ローカル開発時に.envファイルから環境変数を読み込む"""
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(env_path):
        with open(env_path, "r") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, value = line.split("=", 1)
                    os.environ.setdefault(key.strip(), value.strip())
