# Google Search Automation

Google Custom Search APIを使ってスプレッドシートのキーワードを自動検索し、結果をスプレッドシートに保存するツールです。

## GitHub Actionsでの自動実行設定手順

### 1. GitHubリポジトリの準備

1. GitHubでリポジトリを作成
2. このコードをプッシュ

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

### 2. Google APIs の設定

#### Google Custom Search API
1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクト作成
2. Custom Search API を有効化
3. APIキーを作成（制限: Custom Search API のみ）

#### Google Custom Search Engine
1. [Custom Search Engine](https://cse.google.com/) でカスタム検索エンジンを作成
2. 検索エンジンIDをコピー

#### Google Sheets API
1. Google Cloud Console で Sheets API を有効化
2. サービスアカウントを作成
3. JSON認証ファイルをダウンロード
4. スプレッドシートにサービスアカウントのメールアドレスを編集者として共有

### 3. GitHub Secrets の設定

リポジトリの `Settings` > `Secrets and variables` > `Actions` で以下を設定：

| Secret名 | 値 |
|----------|-----|
| `GOOGLE_API_KEY` | Google Custom Search APIキー |
| `CUSTOM_SEARCH_ENGINE_ID` | カスタム検索エンジンID |
| `SPREADSHEET_ID` | スプレッドシートID（URLの一部） |
| `SHEET_NAME` | シート名（例: "Sheet1"） |
| `MAX_REQUESTS` | 最大リクエスト数（例: 100） |
| `GOOGLE_SHEETS_CREDENTIALS` | サービスアカウントJSONファイルの内容全体 |

### 4. スプレッドシートの形式

| A列 | B列 | C列 | D列 | E列 | F列 | G列 | H列 |
|-----|-----|-----|-----|-----|-----|-----|-----|
| ID | キーワード1 | キーワード2 | 説明 | メモ | 完了 | リンク | JSON結果 |

- **1行目はヘッダー行として自動でスキップされます**
- B列 + C列が検索キーワードとして結合されます
- F列に"○"が入ると処理済みとしてスキップされます
- G列に検索結果のリンク（カンマ区切り）が保存されます
- H列に詳細なJSON結果が保存されます

### 5. 実行スケジュール

- **自動実行**: 毎日日本時間午前5時
- **手動実行**: GitHubのActionsタブから「Run workflow」で実行可能

### 6. 実行結果の確認

1. GitHub Actionsの実行ログで進捗確認
2. スプレッドシートで結果確認
3. Actions の Artifacts で詳細データをダウンロード可能

## ローカル実行

```bash
# 依存関係をインストール
pip install -r requirements.txt

# 設定ファイルを作成
cp config.example.json config.json
# config.json を編集

# Google Sheets認証ファイルを配置
# credentials.json

# 実行
python google_search.py
```

## 注意事項

- Google Custom Search APIは1日100回まで無料
- APIクォータを超えないよう `max_requests` を適切に設定
- スプレッドシートのサービスアカウント共有を忘れずに
- 認証情報は絶対にコードに含めない

## トラブルシューティング

### よくあるエラー

1. **403 Forbidden**: APIキーの権限不足 → APIキーの制限設定を確認
2. **400 Bad Request**: 検索エンジンID間違い → CSE IDを再確認
3. **Sheets API エラー**: 認証問題 → サービスアカウントの共有設定を確認