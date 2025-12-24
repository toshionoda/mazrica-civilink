"""
Civilink 組織・ユーザー情報を取得してGoogle Sheetsに出力するスクリプト

使用方法:
  python civilink_org_users.py

環境変数:
  CIVILINK_EMAIL: Civilinkログインメール
  CIVILINK_PASSWORD: Civilinkログインパスワード
  GOOGLE_CREDENTIALS_JSON: Google認証情報JSON
  SPREADSHEET_ID: 出力先スプレッドシートID
"""

import os
import sys
from datetime import datetime
from typing import Optional

from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeoutError

# Google Sheets クライアントを再利用
sys.path.insert(0, os.path.dirname(__file__))
from mazrica.google_sheets_client import GoogleSheetsClient, GoogleSheetsError


class CivilinkScraper:
    """Civilink管理画面のスクレイパー"""

    BASE_URL = "https://civilink.malme.app"
    LOGIN_URL = f"{BASE_URL}/login"
    ACCOUNTS_URL = f"{BASE_URL}/admin/accounts"

    def __init__(self, email: str, password: str, headless: bool = True):
        self.email = email
        self.password = password
        self.headless = headless
        self.browser = None
        self.page: Optional[Page] = None

    def __enter__(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=self.headless)
        self.page = self.browser.new_page()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

    def login(self) -> bool:
        """Civilinkにログイン"""
        print(f"ログイン中: {self.LOGIN_URL}")
        self.page.goto(self.LOGIN_URL)

        # ページ読み込み完了を待機（networkidleはSPAでハングするためloadを使用）
        self.page.wait_for_load_state("load", timeout=30000)
        self.page.wait_for_timeout(3000)

        # デバッグ: スクリーンショット
        self.page.screenshot(path="debug_login_page.png")
        print(f"ログインページURL: {self.page.url}")

        # ログインフォームに入力
        email_input = self.page.locator('input[type="email"], input[name="email"]').first
        password_input = self.page.locator('input[type="password"], input[name="password"]').first

        email_input.wait_for(timeout=10000)
        email_input.fill(self.email)
        password_input.fill(self.password)

        print("認証情報を入力完了")

        # ログインボタンをクリック
        login_button = self.page.locator('button:has-text("メールアドレスでログイン")').first
        login_button.click()

        print("ログインボタンをクリック")

        # ログイン処理完了を待機（ログインフォームが消えるまで）
        try:
            # ログインボタンが消えるか、URLが/loginから変わるまで待機
            self.page.wait_for_timeout(3000)

            # ログイン後のページを確認
            self.page.wait_for_load_state("load", timeout=30000)
            self.page.wait_for_timeout(3000)

            current_url = self.page.url
            print(f"ログイン後URL: {current_url}")

            # デバッグ: スクリーンショット
            self.page.screenshot(path="debug_after_login.png")

            # /login ページに留まっていたら失敗
            if "/login" in current_url:
                print("ログイン失敗: ログインページに留まっています")
                return False

            # callbackUrl付きのリダイレクトはログインページへの戻り
            if "callbackUrl" in current_url:
                print("ログイン失敗: 認証されていません")
                return False

            print("ログイン成功")
            return True

        except PlaywrightTimeoutError:
            print("ログイン失敗: タイムアウト")
            self.page.screenshot(path="debug_login_timeout.png")
            return False

    def navigate_to_accounts(self):
        """アカウント一覧ページに遷移"""
        print(f"アカウント一覧に遷移: {self.ACCOUNTS_URL}")
        self.page.goto(self.ACCOUNTS_URL)
        self.page.wait_for_load_state("load", timeout=30000)

        # デバッグ: 現在のURL確認
        print(f"現在のURL: {self.page.url}")
        print(f"ページタイトル: {self.page.title()}")

        # SPAの動的コンテンツを待機（最大30秒）
        print("テーブルの読み込みを待機中...")
        try:
            # テーブルまたはtr要素が表示されるまで待機
            self.page.wait_for_selector("table, tr", timeout=30000)
            print("テーブル要素を検出")
        except PlaywrightTimeoutError:
            print("テーブル要素が見つかりません（30秒待機後）")

        # 追加の待機（動的コンテンツのため）
        self.page.wait_for_timeout(5000)

        # デバッグ: スクリーンショット保存
        self.page.screenshot(path="debug_accounts_page.png")
        print("スクリーンショット保存: debug_accounts_page.png")

        # デバッグ: ページのHTML構造を出力
        html_content = self.page.content()
        print(f"ページHTML（先頭3000文字）:\n{html_content[:3000]}")

        # デバッグ: 各種セレクタで要素数を確認
        print("\n=== セレクタ検証 ===")
        print(f"table: {self.page.locator('table').count()}")
        print(f"tr: {self.page.locator('tr').count()}")
        print(f"div with role=row: {self.page.locator('[role=row]').count()}")
        print(f"div with role=grid: {self.page.locator('[role=grid]').count()}")
        print(f"text=未登録: {self.page.locator('text=未登録').count()}")
        print(f"text=株式会社: {self.page.locator('text=株式会社').count()}")
        print(f"text=招待中: {self.page.locator('text=招待中').count()}")
        dots_selector = 'button:has-text("...")'
        print(f"button with ...: {self.page.locator(dots_selector).count()}")

    def get_organizations_and_users(self) -> list[dict]:
        """組織とユーザー情報を取得"""
        results = []

        # デバッグ: ページ内のtr要素数を確認
        all_tr = self.page.locator("tr").all()
        print(f"デバッグ: 全tr要素数: {len(all_tr)}")

        # テーブルの行を取得（CSSのエスケープを回避）
        rows = self.page.locator("table tr").all()
        print(f"組織数: {len(rows)}")

        for i, row in enumerate(rows):
            try:
                # 招待中ラベルがあるかチェック（オレンジ色のバッジ）
                invited_badge = row.locator("span.bg-orange-100")
                if invited_badge.count() > 0:
                    print(f"  [{i+1}] スキップ: 招待中")
                    continue

                # 組織情報を取得
                cells = row.locator("td").all()
                if len(cells) < 6:
                    continue

                org_name = cells[0].inner_text().strip()
                dept_name = cells[1].inner_text().strip()
                contractor_name = cells[2].inner_text().strip()
                org_email = cells[3].inner_text().strip()
                phone = cells[4].inner_text().strip()
                created_at = cells[5].inner_text().strip()

                # 組織名が空または"-"の場合はスキップ
                if not org_name or org_name == "-" or org_name == "未登録":
                    print(f"  [{i+1}] スキップ: 組織名なし")
                    continue

                print(f"  [{i+1}] 処理中: {org_name}")

                org_info = {
                    "org_name": org_name,
                    "dept_name": dept_name if dept_name != "-" else "",
                    "contractor_name": contractor_name if contractor_name != "-" else "",
                    "org_email": org_email if org_email != "-" else "",
                    "phone": phone if phone != "-" else "",
                    "created_at": created_at if created_at != "-" else "",
                }

                # 三点リーダーをクリック
                menu_button = row.locator('button[data-slot="dropdown-menu-trigger"]').first
                if menu_button.count() == 0:
                    # 別のセレクタを試す
                    menu_button = row.locator('[aria-label="メニュー"], [data-testid="menu-button"]').first

                if menu_button.count() > 0:
                    menu_button.click()
                    self.page.wait_for_timeout(300)

                    # 「組織ユーザー」メニューをクリック
                    user_menu = self.page.locator('text=組織ユーザー').first
                    if user_menu.count() > 0:
                        user_menu.click()
                        self.page.wait_for_timeout(500)

                        # ポップアップからユーザー情報を取得
                        users = self._get_users_from_popup()

                        for user in users:
                            results.append({
                                **org_info,
                                "user_email": user["email"],
                                "user_name": user["name"],
                                "role": user["role"],
                            })

                        # ポップアップを閉じる
                        self._close_popup()
                    else:
                        # メニューを閉じる
                        self.page.keyboard.press("Escape")
                else:
                    # ユーザー情報なしで登録
                    results.append({
                        **org_info,
                        "user_email": "",
                        "user_name": "",
                        "role": "",
                    })

            except Exception as e:
                print(f"  エラー: {e}")
                continue

        return results

    def _get_users_from_popup(self) -> list[dict]:
        """ポップアップからユーザー一覧を取得"""
        users = []

        try:
            # ポップアップが表示されるまで待機
            popup = self.page.locator('[role="dialog"], .modal, [data-testid="modal"]').first
            popup.wait_for(timeout=5000)

            # ユーザー行を取得（ポップアップ内のテーブルまたはリスト）
            user_rows = popup.locator("table tbody tr, .user-row, [data-testid='user-row']").all()

            if len(user_rows) == 0:
                # テーブルがない場合、別の構造を試す
                # スクリーンショットから見ると、メールアドレスとユーザー名が並んでいる
                email_elements = popup.locator("text=@").all()

                for elem in email_elements:
                    parent = elem.locator("..").first
                    text = parent.inner_text()
                    lines = [l.strip() for l in text.split("\n") if l.strip()]

                    if len(lines) >= 2:
                        email = lines[0] if "@" in lines[0] else ""
                        name = lines[1] if len(lines) > 1 else ""
                        role = lines[2] if len(lines) > 2 else ""

                        if email:
                            users.append({
                                "email": email,
                                "name": name,
                                "role": role,
                            })
            else:
                for user_row in user_rows:
                    cells = user_row.locator("td").all()
                    if len(cells) >= 2:
                        email = cells[0].inner_text().strip()
                        name = cells[1].inner_text().strip()
                        role = cells[2].inner_text().strip() if len(cells) > 2 else ""

                        users.append({
                            "email": email,
                            "name": name,
                            "role": role,
                        })

        except PlaywrightTimeoutError:
            print("    ポップアップが見つかりません")
        except Exception as e:
            print(f"    ユーザー取得エラー: {e}")

        print(f"    ユーザー数: {len(users)}")
        return users

    def _close_popup(self):
        """ポップアップを閉じる"""
        try:
            # ×ボタンを探す
            close_btn = self.page.locator('button:has-text("×"), [aria-label="閉じる"], .close-button').first
            if close_btn.count() > 0:
                close_btn.click()
            else:
                # Escapeキーで閉じる
                self.page.keyboard.press("Escape")

            self.page.wait_for_timeout(300)
        except Exception:
            self.page.keyboard.press("Escape")


def main():
    """メイン処理"""
    # 環境変数から認証情報を取得
    email = os.environ.get("CIVILINK_EMAIL")
    password = os.environ.get("CIVILINK_PASSWORD")
    apps_script_url = os.environ.get("APPS_SCRIPT_URL")

    # バリデーション
    errors = []
    if not email:
        errors.append("CIVILINK_EMAIL が設定されていません")
    if not password:
        errors.append("CIVILINK_PASSWORD が設定されていません")
    if not apps_script_url:
        errors.append("APPS_SCRIPT_URL が設定されていません")

    if errors:
        for e in errors:
            print(f"エラー: {e}")
        sys.exit(1)

    # ヘッドレスモード（GitHub Actionsでは True、ローカルでは False にして確認可能）
    headless = os.environ.get("HEADLESS", "true").lower() == "true"

    print("=" * 50)
    print("Civilink 組織・ユーザー情報取得")
    print(f"開始: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)

    # スクレイピング実行
    with CivilinkScraper(email, password, headless=headless) as scraper:
        if not scraper.login():
            print("ログインに失敗しました")
            sys.exit(1)

        scraper.navigate_to_accounts()
        data = scraper.get_organizations_and_users()

    print(f"\n取得件数: {len(data)}")

    if not data:
        print("データがありません")
        sys.exit(0)

    # Google Sheetsに書き込み
    print("\nGoogle Sheetsに書き込み中...")

    try:
        client = GoogleSheetsClient(
            apps_script_url=apps_script_url
        )

        # ヘッダー行
        headers = [
            "組織名", "部署名", "契約者名", "組織メールアドレス",
            "電話番号", "アカウント作成日", "ユーザーメールアドレス",
            "ユーザー名", "権限"
        ]

        # データ行
        rows = [headers]
        for item in data:
            rows.append([
                item["org_name"],
                item["dept_name"],
                item["contractor_name"],
                item["org_email"],
                item["phone"],
                item["created_at"],
                item["user_email"],
                item["user_name"],
                item["role"],
            ])

        sheet_name = "Civilink_ユーザー"
        client.write_data(sheet_name, rows)
        client.format_header_row(sheet_name)
        client.auto_resize_columns(sheet_name, len(headers))

        print(f"書き込み完了: {len(rows) - 1} 件")

    except GoogleSheetsError as e:
        print(f"Google Sheets書き込みエラー: {e}")
        sys.exit(1)

    print("\n" + "=" * 50)
    print(f"完了: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)


if __name__ == "__main__":
    main()
