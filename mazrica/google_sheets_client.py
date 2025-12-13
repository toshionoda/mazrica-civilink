"""
Google Sheets クライアント（Apps Script連携版）
Apps Script WebアプリへのPOSTリクエストでスプレッドシートにデータを書き込む
"""
import json
import requests
from typing import Optional, Any

from .config import Config


class GoogleSheetsError(Exception):
    """Google Sheets エラー"""
    pass


class GoogleSheetsClient:
    """Google Sheets クライアント（Apps Script連携）"""
    
    def __init__(
        self,
        apps_script_url: Optional[str] = None,
        secret_key: Optional[str] = None
    ):
        """
        Args:
            apps_script_url: Apps ScriptのWebアプリURL
            secret_key: オプションのセキュリティキー
        """
        self.apps_script_url = apps_script_url or Config.APPS_SCRIPT_URL
        self.secret_key = secret_key or getattr(Config, 'APPS_SCRIPT_SECRET', '')
        
        if not self.apps_script_url:
            raise GoogleSheetsError("APPS_SCRIPT_URL is required")
    
    def _post(self, data: dict) -> dict:
        """
        Apps ScriptにPOSTリクエストを送信
        
        Args:
            data: 送信するデータ
        
        Returns:
            レスポンスJSON
        """
        import logging
        logger = logging.getLogger(__name__)
        
        if self.secret_key:
            data['secret_key'] = self.secret_key
        
        try:
            # Apps Scriptはリダイレクトを返すため、allow_redirects=Trueで自動追従
            # セッションを使用してCookieを保持
            session = requests.Session()
            
            logger.info(f"Sending POST request to Apps Script...")
            
            response = session.post(
                self.apps_script_url,
                json=data,
                headers={
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                timeout=300,  # 5分タイムアウト（大量データ対応）
                allow_redirects=True  # リダイレクトを自動追従
            )
            
            logger.info(f"Response status: {response.status_code}")
            logger.info(f"Response URL: {response.url}")
            
            # レスポンスの内容をログ出力（デバッグ用）
            if response.status_code != 200:
                logger.error(f"Response text: {response.text[:500]}")
            
            response.raise_for_status()
            
            result = response.json()
            
            if not result.get('success'):
                raise GoogleSheetsError(f"Apps Script error: {result.get('message')}")
            
            return result
            
        except requests.exceptions.Timeout:
            raise GoogleSheetsError("Request timeout - data may be too large")
        except requests.exceptions.RequestException as e:
            raise GoogleSheetsError(f"Request failed: {e}")
        except json.JSONDecodeError:
            raise GoogleSheetsError("Invalid response from Apps Script")
    
    def ping(self) -> bool:
        """
        接続テスト
        
        Returns:
            接続成功ならTrue
        """
        try:
            result = self._post({'action': 'ping'})
            return result.get('success', False)
        except GoogleSheetsError:
            return False
    
    def write_data(
        self,
        sheet_name: str,
        data: list[list[Any]],
        start_cell: str = "A1",
        clear_before_write: bool = True
    ):
        """
        データをシートに書き込む
        
        Args:
            sheet_name: シート名
            data: 2次元配列のデータ（ヘッダー行を含む）
            start_cell: 書き込み開始セル（未使用、互換性のため）
            clear_before_write: 書き込み前にシートをクリアするか
        """
        if not data:
            raise GoogleSheetsError("No data to write")
        
        # ヘッダー行とデータ行を分離
        headers = data[0] if data else []
        rows = data[1:] if len(data) > 1 else []
        
        # データを文字列に変換（JSON互換性のため）
        def convert_value(v):
            if v is None:
                return ""
            return v
        
        converted_rows = [
            [convert_value(cell) for cell in row]
            for row in rows
        ]
        converted_headers = [convert_value(cell) for cell in headers]
        
        # Apps Scriptにデータを送信
        payload = {
            'action': 'write',
            'sheet_name': sheet_name,
            'headers': converted_headers,
            'rows': converted_rows,
            'clear_before': clear_before_write
        }
        
        result = self._post(payload)
        return result
    
    def clear_sheet(self, sheet_name: str):
        """シートの内容をクリア"""
        payload = {
            'action': 'clear',
            'sheet_name': sheet_name
        }
        return self._post(payload)
    
    def format_header_row(self, sheet_name: str):
        """
        ヘッダー行の書式設定
        注: Apps Script側で自動的に行われるため、この関数は何もしない
        """
        pass  # Apps Script側で処理
    
    def auto_resize_columns(self, sheet_name: str, column_count: int):
        """
        列幅を自動調整
        注: Apps Script側で自動的に行われるため、この関数は何もしない
        """
        pass  # Apps Script側で処理


# テスト用
if __name__ == "__main__":
    from .config import load_dotenv
    load_dotenv()
    
    client = GoogleSheetsClient()
    
    # 接続テスト
    print("Testing connection...")
    if client.ping():
        print("Connection successful!")
    else:
        print("Connection failed!")
        exit(1)
    
    # テストデータ
    test_data = [
        ["案件ID", "案件名", "取引先", "商品名", "数量", "単価", "金額", "更新日時"],
        [1, "テスト案件A", "ABC社", "商品X", 10, 1000, 10000, "2025-12-11"],
        [2, "テスト案件B", "DEF社", "商品Y", 5, 2000, 10000, "2025-12-11"],
    ]
    
    sheet_name = "テスト"
    result = client.write_data(sheet_name, test_data)
    print(f"Result: {result}")
