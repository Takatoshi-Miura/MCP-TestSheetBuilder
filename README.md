# MCP Test Sheet Builder

GoogleスプレッドシートとGoogleドライブを使用して、テスト項目を自動生成するMCPサーバーです。

## 機能

- GoogleDriveにアクセス可能
- GoogleDrive内のスプレッドシートを閲覧・編集可能
- 指定したテンプレートのスプレッドシートをコピーしてファイルを作成可能
- プロンプトで指定された要件をテストするための因子水準を作成し、スプレッドシートに記載可能
- 因子水準に基づいてテスト項目をスプレッドシートに作成可能

## MCPツール

このサーバーは以下のMCPツールを提供します：

1. `mcp_test_sheet_builder_get_spreadsheet`: スプレッドシートの情報を取得するツール
   - パラメータ:
     - `id`: スプレッドシートのID（必須）
     - `range`: 取得する範囲（例: Sheet1!A1:Z100）（省略可）

2. `mcp_test_sheet_builder_generate_test_items`: 因子・水準シートからテスト項目を自動生成するツール
   - パラメータ:
     - `spreadsheetId`: テスト項目を生成するスプレッドシートのID（必須）

## セットアップ方法

1. Google Cloud Consoleでプロジェクトを作成
2. Google Drive APIとGoogle Sheets APIを有効化
3. OAuth同意画面を設定
4. OAuth 2.0クライアントIDを作成し、認証情報をダウンロード
5. ダウンロードしたJSONファイルを`credentials/client_secret.json`として保存
6. 以下のコマンドでトークンを生成:
   ```
   node simple-token-generator.js
   ```
   - このコマンドを実行すると、認証URLが表示されます
   - URLをコピーしてブラウザで開き、Google認証を行います
   - 認証が完了すると自動的に`mcp-test-sheet-builder/credentials/token.json`が生成されます
   - 「トークンが保存されました」と表示されれば認証は成功です
7. Cursor SettingsのMCP Serversで「Add new global MCP server」を押下し、mcp.jsonに以下を追記
   ```json
   "mcp_test_sheet_builder": {
      "command": "node",
      "args": [
        "[実際のパスを設定する]/MCP-TestSheetBuilder/mcp-test-sheet-builder/build/index.js"
      ]
    }
   ```
7. mcp_test_sheet_builderを有効にする

## トラブルシューティング

- トークンが期限切れになった場合は、`simple-token-generator.js`を使用して新しいトークンを生成してください
- Google APIの権限が不足している場合は、Google Cloud Consoleでプロジェクトの権限を確認してください
- トークン生成プロセスで問題が発生した場合の対処法:
  - `client_secret.json`の内容が正しいか確認する
  - OAuth同意画面で適切なスコープが設定されているか確認する
  - Google Cloud Consoleでリダイレクトに`http://localhost:8080`が登録されているか確認する
  - Node.jsのバージョンが17以上であることを確認する（ESMサポートのため）

## 技術詳細

- このバージョンではMCPプロトコルのみをサポートしています
- パラメータのバリデーションにZodを使用しています
- GoogleスプレッドシートとGoogleドライブの操作にGoogle API Node.js クライアントライブラリを使用しています
- 認証にはOAuth 2.0を使用し、ブラウザ経由でのユーザー認証を実装しています
