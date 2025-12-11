# Mazrica → Google Sheets 同期ツール

Mazrica APIから案件一覧（商品内訳付き）を取得し、Google スプレッドシートに自動同期するツールです。

## 機能

- Mazricaの案件一覧を取得
- 商品内訳情報を含めて取得・展開
- Google スプレッドシートへの自動書き込み
- GitHub Actionsによる定期実行（毎日9:00 JST）

## 出力データ形式

| カラム | 説明 |
|--------|------|
| 案件ID | Mazricaの案件ID |
| 案件名 | 案件の名称 |
| 取引先 | 取引先名 |
| 取引先ID | 取引先ID |
| 案件タイプ | 案件タイプ名 |
| フェーズ | 現在のフェーズ |
| 担当者 | 担当者名 |
| 商品名 | 商品内訳の商品名 |
| 数量 | 商品内訳の数量 |
| 単価 | 商品内訳の単価 |
| 商品金額 | 商品内訳の金額 |
| 案件金額 | 案件全体の金額 |
| 受注予定日 | 受注予定日 |
| 作成日時 | 案件作成日時 |
| 更新日時 | 案件更新日時 |

## セットアップ手順

### 1. Google Cloud設定

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新しいプロジェクトを作成（または既存プロジェクトを選択）
3. 「APIとサービス」→「ライブラリ」から **Google Sheets API** を有効化
4. 「APIとサービス」→「認証情報」から「サービスアカウント」を作成
5. 作成したサービスアカウントの「キー」タブから **JSONキー** をダウンロード

### 2. スプレッドシートの準備

1. 同期先のGoogle スプレッドシートを作成
2. スプレッドシートの共有設定で、サービスアカウントのメールアドレス（`xxx@xxx.iam.gserviceaccount.com`）を **編集者** として追加
3. スプレッドシートのURLからIDを取得（`https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit`）

### 3. GitHub Secrets設定

リポジトリの「Settings」→「Secrets and variables」→「Actions」で以下のシークレットを設定:

| Secret名 | 内容 | 例 |
|----------|------|-----|
| `MAZRICA_API_KEY` | Mazrica APIキー | `FBHKOJjizP9l2gNP9G6866Y22n0d64QC3VVY4HmA` |
| `GOOGLE_CREDENTIALS_JSON` | サービスアカウントJSONキー（**中身をそのままペースト**） | `{"type":"service_account",...}` |
| `SPREADSHEET_ID` | スプレッドシートID | `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms` |

### 4. Mazrica APIキーの取得

1. [Mazrica管理画面](https://product-senses.mazrica.com/) にログイン
2. 「設定」→「API設定」からAPIキーを発行

## ローカル実行

### 環境変数の設定

`.env` ファイルを `mazrica/` ディレクトリに作成:

```bash
# mazrica/.env
MAZRICA_API_KEY=your_api_key_here
GOOGLE_CREDENTIALS_JSON={"type":"service_account",...}
SPREADSHEET_ID=your_spreadsheet_id
SHEET_NAME=案件一覧
DEAL_TYPE_ID=  # 特定の案件タイプのみ同期する場合は指定
```

### 依存パッケージのインストール

```bash
pip install -r mazrica/requirements.txt
```

### 実行

```bash
python -m mazrica.sync_to_sheets
```

## GitHub Actions手動実行

1. リポジトリの「Actions」タブを開く
2. 「Mazrica Sheets Sync」ワークフローを選択
3. 「Run workflow」をクリック
4. 必要に応じてパラメータを入力して実行

### パラメータ

| パラメータ | 説明 | デフォルト |
|------------|------|-----------|
| `deal_type_id` | 同期する案件タイプID（空で全案件） | 空 |
| `sheet_name` | 出力先シート名 | `案件一覧` |

## API制限

Mazrica APIには以下の制限があります:

- **50,000リクエスト/日**
- **3リクエスト/秒**
- **レスポンス最大10MB**

本ツールはレート制限を考慮して実装されていますが、大量の案件がある場合は注意してください。

## トラブルシューティング

### エラー: MAZRICA_API_KEY が設定されていません

- 環境変数またはGitHub Secretsに `MAZRICA_API_KEY` が設定されているか確認

### エラー: Google Sheets Error

- サービスアカウントがスプレッドシートに共有されているか確認
- `GOOGLE_CREDENTIALS_JSON` のJSON形式が正しいか確認
- スプレッドシートIDが正しいか確認

### エラー: Mazrica API Error (429)

- APIレート制限に達しています。しばらく待ってから再実行してください

## ファイル構成

```
mazrica/
├── __init__.py              # パッケージ初期化
├── config.py                # 設定管理
├── mazrica_client.py        # Mazrica APIクライアント
├── google_sheets_client.py  # Google Sheets APIクライアント
├── sync_to_sheets.py        # メイン同期スクリプト
├── requirements.txt         # 依存パッケージ
└── README.md                # このファイル

.github/workflows/
└── mazrica_sync.yml         # GitHub Actions定期実行
```

## ライセンス

Private - Internal use only

