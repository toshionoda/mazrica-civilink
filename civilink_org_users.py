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
        """組織とユーザー情報を取得（ページリロード方式）"""
        results = []

        # 最初に組織数を取得
        rows = self.page.locator("table tr").all()
        total_rows = len(rows)
        print(f"組織数: {total_rows}")

        # 各組織をインデックスで処理（リロード後も継続できるように）
        i = 0
        while i < total_rows:
            try:
                # ページリロード後はテーブルを再取得
                row = self.page.locator("table tr").nth(i)

                # 招待中ラベルがあるかチェック（オレンジ色のバッジ）
                invited_badge = row.locator("span.bg-orange-100")
                if invited_badge.count() > 0:
                    print(f"  [{i+1}] スキップ: 招待中")
                    i += 1
                    continue

                # 組織情報を取得
                cells = row.locator("td").all()
                if len(cells) < 6:
                    i += 1
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
                    i += 1
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

                # トグル状態の初期値
                bulletin_board = "---"
                rebar_ai = "---"

                # 三点リーダーをクリック
                menu_button = row.locator('button[data-slot="dropdown-menu-trigger"]').first
                if menu_button.count() == 0:
                    menu_button = row.locator('[aria-label="メニュー"], [data-testid="menu-button"]').first

                if menu_button.count() > 0:
                    menu_button.click()
                    self.page.wait_for_timeout(500)

                    # 「編集」メニューをクリックしてトグル状態を取得
                    edit_menu = self.page.locator('text=編集').first
                    if edit_menu.count() > 0:
                        try:
                            edit_menu.click()
                            self.page.wait_for_timeout(1000)

                            toggles = self._get_edit_popup_toggles()
                            bulletin_board = toggles["bulletin_board"]
                            rebar_ai = toggles["rebar_ai"]

                            self._close_edit_popup()
                            self.page.wait_for_timeout(500)
                        except Exception as e:
                            print(f"    編集ポップアップ処理エラー: {e}")
                            # エラー時はポップアップを閉じる試行
                            try:
                                self._close_edit_popup()
                            except:
                                pass

                        # メニューを再度開く
                        row = self.page.locator("table tr").nth(i)
                        menu_button = row.locator('button[data-slot="dropdown-menu-trigger"]').first
                        if menu_button.count() == 0:
                            menu_button = row.locator('[aria-label="メニュー"], [data-testid="menu-button"]').first
                        menu_button.click()
                        self.page.wait_for_timeout(500)
                    else:
                        print("    「編集」メニュー項目が見つかりません、スキップ")
                        # メニューを閉じて再度開く
                        self.page.keyboard.press("Escape")
                        self.page.wait_for_timeout(300)
                        menu_button = row.locator('button[data-slot="dropdown-menu-trigger"]').first
                        if menu_button.count() == 0:
                            menu_button = row.locator('[aria-label="メニュー"], [data-testid="menu-button"]').first
                        menu_button.click()
                        self.page.wait_for_timeout(500)

                    # 「組織ユーザー」メニューをクリック
                    user_menu = self.page.locator('text=組織ユーザー').first
                    if user_menu.count() > 0:
                        user_menu.click()
                        self.page.wait_for_timeout(1000)

                        # ポップアップが表示されるまで待機
                        popup = self.page.locator('#organization_users_id')
                        try:
                            popup.wait_for(state="visible", timeout=5000)

                            # ポップアップからユーザー情報を取得
                            users = self._get_users_from_popup()
                            print(f"    ユーザー数: {len(users)}")

                            for user in users:
                                results.append({
                                    **org_info,
                                    "user_email": user["email"],
                                    "user_name": user["name"],
                                    "role": user["role"],
                                    "bulletin_board": bulletin_board,
                                    "rebar_ai": rebar_ai,
                                })

                            # ページをリロードしてSPA状態をリセット
                            print(f"    ページリロード")
                            self.page.reload()
                            self.page.wait_for_load_state("load", timeout=30000)
                            self.page.wait_for_selector("table tr", timeout=10000)
                            self.page.wait_for_timeout(1000)

                        except Exception as e:
                            print(f"    ポップアップエラー: {e}")
                            # エラー時もリロードして継続
                            self.page.reload()
                            self.page.wait_for_load_state("load", timeout=30000)
                            self.page.wait_for_selector("table tr", timeout=10000)
                    else:
                        self.page.keyboard.press("Escape")
                else:
                    # ユーザー情報なしで登録
                    results.append({
                        **org_info,
                        "user_email": "",
                        "user_name": "",
                        "role": "",
                        "bulletin_board": bulletin_board,
                        "rebar_ai": rebar_ai,
                    })

                i += 1

            except Exception as e:
                print(f"  エラー: {e}")
                # エラー時はリロードして継続
                try:
                    self.page.reload()
                    self.page.wait_for_load_state("load", timeout=30000)
                    self.page.wait_for_selector("table tr", timeout=10000)
                except:
                    pass
                i += 1
                continue

        return results

    def _get_users_from_popup(self) -> list[dict]:
        """ポップアップからユーザー一覧を取得"""
        users = []

        try:
            # ポップアップを取得（#organization_users_id を使用）
            popup = self.page.locator('#organization_users_id')
            if popup.count() == 0:
                print("    ポップアップが見つかりません")
                return users

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

    def _read_toggle_state(self, dialog, label_text: str) -> str:
        """ダイアログ内のトグル状態を読み取る

        Args:
            dialog: ダイアログのLocator
            label_text: トグルのラベルテキスト（例: "掲示板スレッド"）

        Returns:
            "表示" or "---"
        """
        try:
            # ラベルテキストを含む要素を検索
            label = dialog.locator(f"text={label_text}").first
            if label.count() == 0:
                print(f"    ラベル '{label_text}' が見つかりません")
                return "---"

            # ラベルの親要素からボタンを取得
            parent = label.locator("..").first
            buttons = parent.locator("button").all()

            if not buttons:
                # さらに上の親を試す
                parent = parent.locator("..").first
                buttons = parent.locator("button").all()

            if not buttons:
                print(f"    '{label_text}' のボタンが見つかりません")
                return "---"

            # アクティブなボタンを判定
            for button in buttons:
                # 戦略1: data-state 属性（Radix UI パターン）
                data_state = button.get_attribute("data-state")
                if data_state in ("on", "active"):
                    text = button.inner_text().strip()
                    return "表示" if text == "表示" else "---"

                # 戦略2: aria-pressed 属性
                aria_pressed = button.get_attribute("aria-pressed")
                if aria_pressed == "true":
                    text = button.inner_text().strip()
                    return "表示" if text == "表示" else "---"

            # 戦略3: CSSクラスで判定（bg-blue等）
            for button in buttons:
                class_attr = button.get_attribute("class") or ""
                if "bg-blue" in class_attr or "bg-primary" in class_attr or "active" in class_attr:
                    text = button.inner_text().strip()
                    return "表示" if text == "表示" else "---"

            # 判定できない場合はデバッグ情報を出力
            print(f"    '{label_text}' のアクティブボタンを判定できません:")
            for idx, button in enumerate(buttons):
                attrs = {
                    "text": button.inner_text().strip(),
                    "data-state": button.get_attribute("data-state"),
                    "aria-pressed": button.get_attribute("aria-pressed"),
                    "class": button.get_attribute("class"),
                }
                print(f"      ボタン{idx}: {attrs}")
            return "---"

        except Exception as e:
            print(f"    トグル状態読み取りエラー ({label_text}): {e}")
            return "---"

    def _get_edit_popup_toggles(self) -> dict:
        """編集ポップアップからトグル状態を取得

        Returns:
            {"bulletin_board": "---"/"表示", "rebar_ai": "---"/"表示"}
        """
        defaults = {"bulletin_board": "---", "rebar_ai": "---"}
        try:
            # ダイアログが表示されるまで待機
            dialog = self.page.locator('[role="dialog"]').first
            dialog.wait_for(state="visible", timeout=5000)
            self.page.wait_for_timeout(500)

            # デバッグ用スクリーンショット
            self.page.screenshot(path="debug_edit_popup.png")
            print("    スクリーンショット保存: debug_edit_popup.png")

            # トグル状態を読み取り
            bulletin_board = self._read_toggle_state(dialog, "掲示板スレッド")
            rebar_ai = self._read_toggle_state(dialog, "鉄筋照査AI")

            print(f"    掲示板スレッド: {bulletin_board}, 鉄筋照査AI: {rebar_ai}")
            return {"bulletin_board": bulletin_board, "rebar_ai": rebar_ai}

        except PlaywrightTimeoutError:
            print("    編集ポップアップが表示されませんでした")
            return defaults
        except Exception as e:
            print(f"    編集ポップアップ読み取りエラー: {e}")
            return defaults

    def _close_edit_popup(self):
        """編集ポップアップを閉じる"""
        try:
            dialog = self.page.locator('[role="dialog"]')
            if dialog.count() == 0:
                print("    編集ポップアップは既に閉じています")
                return

            # 方法1: × ボタンをクリック
            close_button = dialog.locator('button:has-text("×"), button[aria-label="Close"], button[aria-label="閉じる"]').first
            if close_button.count() > 0:
                close_button.click()
                self.page.wait_for_timeout(500)
                if dialog.count() == 0:
                    print("    ×ボタンで編集ポップアップを閉じました")
                    return

            # 方法2: Escape キー
            self.page.keyboard.press("Escape")
            self.page.wait_for_timeout(500)
            if dialog.count() == 0:
                print("    Escapeキーで編集ポップアップを閉じました")
                return

            # 方法3: JavaScript で強制削除
            self.page.evaluate("document.querySelector('[role=\"dialog\"]')?.closest('[data-state]')?.remove() || document.querySelector('[role=\"dialog\"]')?.remove()")
            self.page.wait_for_timeout(300)
            print("    JavaScriptで編集ポップアップを削除しました")

        except Exception as e:
            print(f"    編集ポップアップを閉じる際のエラー: {e}")
            try:
                self.page.evaluate("document.querySelector('[role=\"dialog\"]')?.closest('[data-state]')?.remove() || document.querySelector('[role=\"dialog\"]')?.remove()")
                print("    フォールバック: JavaScriptで削除")
            except:
                pass
            self.page.wait_for_timeout(500)

    def _close_popup(self):
        """ポップアップを閉じる"""
        try:
            # organization_users_id のポップアップを閉じる
            popup = self.page.locator('#organization_users_id')

            if popup.count() > 0:
                # 方法1: JavaScriptで直接削除（最も確実）
                self.page.evaluate("document.getElementById('organization_users_id')?.remove()")
                print("    JavaScriptでポップアップを削除しました")
                self.page.wait_for_timeout(300)

                # 削除されたか確認
                if popup.count() == 0:
                    print("    ポップアップ削除成功")
                else:
                    # 方法2: オーバーレイ背景をクリック
                    print("    まだポップアップが存在、背景クリックを試行")
                    self.page.mouse.click(10, 10)
                    self.page.wait_for_timeout(500)
            else:
                print("    ポップアップは既に閉じています")

            self.page.wait_for_timeout(300)
        except Exception as e:
            print(f"    ポップアップを閉じる際のエラー: {e}")
            # フォールバック: JavaScriptで強制削除
            try:
                self.page.evaluate("document.getElementById('organization_users_id')?.remove()")
                print("    フォールバック: JavaScriptで削除")
            except:
                pass
            self.page.wait_for_timeout(500)


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
            "ユーザー名", "権限", "掲示板スレッド", "鉄筋照査AI"
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
                item["bulletin_board"],
                item["rebar_ai"],
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
