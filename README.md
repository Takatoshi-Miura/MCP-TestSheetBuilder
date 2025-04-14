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

1. `generate-test`: テストシートを生成するツール
   - パラメータ:
     - `templateId`: テンプレートとなるスプレッドシートのID（必須）
     - `title`: 生成するスプレッドシートのタイトル（必須）
     - `prompt`: テスト要件を記述したプロンプト（必須）
     - `useOrthogonalArray`: 直交表を使用するかどうか（省略可）

2. `get-spreadsheet`: スプレッドシートの情報を取得するツール
   - パラメータ:
     - `id`: スプレッドシートのID（必須）
     - `range`: 取得する範囲（例: Sheet1!A1:Z100）（省略可）

## APIエンドポイント

- `GET /health`: サーバーのヘルスチェック
- `POST /generate-test-sheet`: テストシートを生成
- `GET /spreadsheet/:id`: スプレッドシートの情報を取得
- `POST /spreadsheet/:id/update`: スプレッドシートを更新

## セットアップ方法

1. Google API認証情報の設定:
   - Google Cloud Consoleで認証情報を作成し、以下のファイルを配置します：
     - `mcp-test-sheet-builder/credentials/client_secret.json`: Google APIのクライアント認証情報
   - クライアント認証情報ファイルを配置した後、初回起動時に認証フローが実行されます

2. サーバーをビルド:
   ```
   cd mcp-test-sheet-builder
   npm install
   npm run build
   ```

3. サーバーを初回起動してトークンを生成:
   ```
   node build/index.js
   ```
   - 表示されたURLにアクセスして認証を行い、コードを取得します
   - コードは http://localhost/?code= 以降の文字列全てです
   - 取得したコードをコンソールに入力するとトークンが生成されます
   - トークンファイル `mcp-test-sheet-builder/credentials/token.json` が自動的に作成されます

4. `.cursor/mcp.json`に以下の設定を追加:
   ```json
   "mcp-test-sheet-builder": {
     "command": "node",
     "args": [
       "~/MCP-TestSheetBuilder/mcp-test-sheet-builder/build/index.js"
     ],
     "env": {
       "PORT": "3005"
     }
   }
   ```

## CursorでMCPツールを使う方法

1. Cursorを起動し、「MCP Server」メニューからmcp-test-sheet-builderを選択して接続
2. 以下のような形式でツールを呼び出します：

```
// テストシートを生成する例
const result = await mcp.callTool('generate-test', {
  templateId: '1abc...xyz', // テンプレートのスプレッドシートID
  title: 'テスト計画書',
  prompt: 'Webアプリケーションのログイン機能をテストする。ブラウザ、OS、ネットワーク環境を考慮すること。'
});

// スプレッドシートの内容を取得する例
const data = await mcp.callTool('get-spreadsheet', {
  id: '1abc...xyz', // スプレッドシートID
  range: 'Sheet1!A1:D10'
});
```

## 認証情報の取得方法

1. Google Cloud Consoleでプロジェクトを作成
2. Google Drive APIとGoogle Sheets APIを有効化
3. OAuth同意画面を設定
4. OAuth 2.0クライアントIDを作成し、認証情報をダウンロード
5. ダウンロードしたJSONファイルを`credentials/client_secret.json`として保存
6. 初回実行時に認証フローを実行してトークンを生成します

## トラブルシューティング

- トークンの有効期限が切れた場合は、`credentials/token.json`を削除して再度認証フローを実行します
- 認証情報のパスを変更する場合は、環境変数`CREDENTIALS_PATH`と`TOKEN_PATH`を設定します
- ポートが既に使用されている場合（「EADDRINUSE: address already in use」エラー）の対処法：
  1. 使用中のプロセスを終了する: `lsof -i :3000` でプロセスを確認し、`kill -9 <PID>` で終了
  2. または、`src/index.ts` の `port` 変数を別のポート番号に変更し、再ビルド
- 「Client closed」エラーが発生する場合：
  1. サーバーが正常に起動しているか確認（ログを確認）
  2. サーバーコードがMCPプロトコルに対応しているか確認
