"""
Mazrica → Google Sheets 同期機能の設定
"""
import os
import json
from typing import Optional


class Config:
    """環境変数から設定を読み込むクラス"""
    
    # Mazrica API設定
    MAZRICA_API_KEY: str = os.environ.get("MAZRICA_API_KEY", "")
    MAZRICA_BASE_URL: str = "https://senses-open-api.mazrica.com/v1"
    
    # Google Sheets設定
    GOOGLE_CREDENTIALS_JSON: str = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    SPREADSHEET_ID: str = os.environ.get("SPREADSHEET_ID", "")
    SHEET_NAME: str = os.environ.get("SHEET_NAME", "案件一覧")
    
    # 同期設定
    DEAL_TYPE_ID: Optional[int] = (
        int(os.environ.get("DEAL_TYPE_ID")) 
        if os.environ.get("DEAL_TYPE_ID") 
        else None
    )
    
    # API制限
    API_RATE_LIMIT: float = 0.34  # 3リクエスト/秒 = 0.33秒間隔
    API_PAGE_SIZE: int = 100  # 1ページあたりの取得件数
    
    @classmethod
    def get_google_credentials(cls) -> Optional[dict]:
        """Google認証情報をJSONとしてパース"""
        if not cls.GOOGLE_CREDENTIALS_JSON:
            return None
        try:
            return json.loads(cls.GOOGLE_CREDENTIALS_JSON)
        except json.JSONDecodeError:
            return None
    
    @classmethod
    def validate(cls) -> list[str]:
        """設定の検証を行い、エラーメッセージのリストを返す"""
        errors = []
        
        if not cls.MAZRICA_API_KEY:
            errors.append("MAZRICA_API_KEY が設定されていません")
        
        if not cls.GOOGLE_CREDENTIALS_JSON:
            errors.append("GOOGLE_CREDENTIALS_JSON が設定されていません")
        elif cls.get_google_credentials() is None:
            errors.append("GOOGLE_CREDENTIALS_JSON のJSON形式が不正です")
        
        if not cls.SPREADSHEET_ID:
            errors.append("SPREADSHEET_ID が設定されていません")
        
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

