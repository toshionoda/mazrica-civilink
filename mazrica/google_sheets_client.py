"""
Google Sheets API クライアント
スプレッドシートへのデータ書き込みを行う
"""
import json
from typing import Optional, Any
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from .config import Config


class GoogleSheetsError(Exception):
    """Google Sheets API エラー"""
    pass


class GoogleSheetsClient:
    """Google Sheets APIクライアント"""
    
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    
    def __init__(
        self,
        credentials_json: Optional[str] = None,
        spreadsheet_id: Optional[str] = None
    ):
        """
        Args:
            credentials_json: サービスアカウントのJSON文字列
            spreadsheet_id: スプレッドシートID
        """
        self.credentials_json = credentials_json or Config.GOOGLE_CREDENTIALS_JSON
        self.spreadsheet_id = spreadsheet_id or Config.SPREADSHEET_ID
        
        self._service = None
    
    def _get_credentials(self) -> Credentials:
        """認証情報を取得"""
        if not self.credentials_json:
            raise GoogleSheetsError("Google credentials JSON is required")
        
        try:
            credentials_dict = json.loads(self.credentials_json)
            return Credentials.from_service_account_info(
                credentials_dict,
                scopes=self.SCOPES
            )
        except json.JSONDecodeError as e:
            raise GoogleSheetsError(f"Invalid JSON format: {e}")
        except Exception as e:
            raise GoogleSheetsError(f"Failed to create credentials: {e}")
    
    @property
    def service(self):
        """Sheets APIサービスを取得（遅延初期化）"""
        if self._service is None:
            credentials = self._get_credentials()
            self._service = build("sheets", "v4", credentials=credentials)
        return self._service
    
    def get_sheet_id(self, sheet_name: str) -> Optional[int]:
        """シート名からシートIDを取得"""
        try:
            spreadsheet = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
            
            for sheet in spreadsheet.get("sheets", []):
                props = sheet.get("properties", {})
                if props.get("title") == sheet_name:
                    return props.get("sheetId")
            
            return None
        except HttpError as e:
            raise GoogleSheetsError(f"Failed to get sheet info: {e}")
    
    def create_sheet(self, sheet_name: str) -> int:
        """新しいシートを作成"""
        try:
            request = {
                "requests": [{
                    "addSheet": {
                        "properties": {
                            "title": sheet_name
                        }
                    }
                }]
            }
            
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=request
            ).execute()
            
            return result["replies"][0]["addSheet"]["properties"]["sheetId"]
        except HttpError as e:
            raise GoogleSheetsError(f"Failed to create sheet: {e}")
    
    def ensure_sheet_exists(self, sheet_name: str) -> int:
        """シートが存在することを確認（なければ作成）"""
        sheet_id = self.get_sheet_id(sheet_name)
        if sheet_id is None:
            sheet_id = self.create_sheet(sheet_name)
        return sheet_id
    
    def clear_sheet(self, sheet_name: str):
        """シートの内容をクリア"""
        try:
            self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=f"'{sheet_name}'!A:Z"
            ).execute()
        except HttpError as e:
            raise GoogleSheetsError(f"Failed to clear sheet: {e}")
    
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
            start_cell: 書き込み開始セル
            clear_before_write: 書き込み前にシートをクリアするか
        """
        try:
            # シートの存在確認・作成
            self.ensure_sheet_exists(sheet_name)
            
            # クリア
            if clear_before_write:
                self.clear_sheet(sheet_name)
            
            # データ書き込み
            range_name = f"'{sheet_name}'!{start_cell}"
            
            self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption="USER_ENTERED",
                body={"values": data}
            ).execute()
            
        except HttpError as e:
            raise GoogleSheetsError(f"Failed to write data: {e}")
    
    def format_header_row(self, sheet_name: str):
        """ヘッダー行の書式設定（太字、背景色）"""
        try:
            sheet_id = self.get_sheet_id(sheet_name)
            if sheet_id is None:
                return
            
            requests = [
                # ヘッダー行を太字に
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 0.9,
                                    "green": 0.9,
                                    "blue": 0.9
                                },
                                "textFormat": {
                                    "bold": True
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat)"
                    }
                },
                # 1行目を固定
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {
                                "frozenRowCount": 1
                            }
                        },
                        "fields": "gridProperties.frozenRowCount"
                    }
                }
            ]
            
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": requests}
            ).execute()
            
        except HttpError as e:
            # 書式設定の失敗は警告のみ
            print(f"Warning: Failed to format header: {e}")
    
    def auto_resize_columns(self, sheet_name: str, column_count: int):
        """列幅を自動調整"""
        try:
            sheet_id = self.get_sheet_id(sheet_name)
            if sheet_id is None:
                return
            
            requests = [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": column_count
                    }
                }
            }]
            
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": requests}
            ).execute()
            
        except HttpError as e:
            # 列幅調整の失敗は警告のみ
            print(f"Warning: Failed to auto-resize columns: {e}")


# テスト用
if __name__ == "__main__":
    from .config import load_dotenv
    load_dotenv()
    
    client = GoogleSheetsClient()
    
    # テストデータ
    test_data = [
        ["案件ID", "案件名", "取引先", "商品名", "数量", "単価", "金額", "更新日時"],
        [1, "テスト案件A", "ABC社", "商品X", 10, 1000, 10000, "2025-12-11"],
        [2, "テスト案件B", "DEF社", "商品Y", 5, 2000, 10000, "2025-12-11"],
    ]
    
    sheet_name = "テスト"
    client.write_data(sheet_name, test_data)
    client.format_header_row(sheet_name)
    client.auto_resize_columns(sheet_name, len(test_data[0]))
    
    print("Test data written successfully!")

