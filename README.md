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

1. `mcp_test_sheet_builder_generate_test`: テストシートを生成するツール
   - パラメータ:
     - `templateId`: テンプレートとなるスプレッドシートのID（必須）
     - `title`: 生成するスプレッドシートのタイトル（必須）
     - `prompt`: テスト要件を記述したプロンプト（必須）
     - `useOrthogonalArray`: 直交表を使用するかどうか（省略可、デフォルト: false）

2. `mcp_test_sheet_builder_get_spreadsheet`: スプレッドシートの情報を取得するツール
   - パラメータ:
     - `id`: スプレッドシートのID（必須）
     - `range`: 取得する範囲（例: Sheet1!A1:Z100）（省略可）

## セットアップ方法

1. Google API認証情報の設定:
   - Google Cloud Consoleで認証情報を作成し、以下のファイルを配置します：
     - `mcp-test-sheet-builder/credentials/client_secret.json`: Google APIのクライアント認証情報
   - 初回起動前に、事前にトークンも取得しておき、以下のパスに配置します:
     - `mcp-test-sheet-builder/credentials/token.json`: Google APIのアクセストークン

2. 依存関係のインストール:
   ```
   cd mcp-test-sheet-builder
   npm install
   ```

3. 開発サーバーの起動:
   ```
   cd mcp-test-sheet-builder
   npm run dev
   ```

## 本番環境での実行

```
cd mcp-test-sheet-builder
npm run build
npm start
```

## CursorでMCPツールを使う方法

1. Cursorを起動し、「MCP Server」メニューからmcp-test-sheet-builderを選択して接続
2. 以下のような形式でツールを呼び出します：

```
// テストシートを生成する例
const result = await mcp.callTool('mcp_test_sheet_builder_generate_test', {
  templateId: '1abc...xyz', // テンプレートのスプレッドシートID
  title: 'テスト計画書',
  prompt: 'Webアプリケーションのログイン機能をテストする。ブラウザ、OS、ネットワーク環境を考慮すること。'
});

// スプレッドシートの内容を取得する例
const data = await mcp.callTool('mcp_test_sheet_builder_get_spreadsheet', {
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
6. 認証トークンを取得するには、以下の方法があります:
   - MCP-TestSheetBuilder以外のGoogleAPI認証フローを持つプロジェクト（例: edit-google-drive）を実行してトークンを取得し、そのトークンをコピーする
   - または、OAuth認証フローを含む旧バージョンの本プロジェクトを一時的に使用してトークンを取得する

## トラブルシューティング

- トークンが期限切れになった場合は、新しいトークンを取得して`credentials/token.json`を更新してください。
- Google APIの権限が不足している場合は、Google Cloud Consoleでプロジェクトの権限を確認してください。

## 技術詳細

- このバージョンではMCPプロトコルのみをサポートしています。
- パラメータのバリデーションにZodを使用しています。
- GoogleスプレッドシートとGoogleドライブの操作にGoogle API Node.js クライアントライブラリを使用しています。
