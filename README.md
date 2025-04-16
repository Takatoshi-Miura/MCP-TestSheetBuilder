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

2. 依存関係のインストール:
   ```
   cd mcp-test-sheet-builder
   npm install
   
   # Node.js v17.3.0を使用している場合は、互換性のある古いバージョンのopenパッケージをインストール
   npm uninstall open
   npm install open@8.4.2
   ```

3. 認証トークンの生成:
   ```
   cd mcp-test-sheet-builder
   node simple-token-generator.js
   ```
   - このコマンドを実行すると、認証URLが表示されます
   - URLをコピーしてブラウザで開き、Google認証を行います
   - 認証が完了すると自動的に`mcp-test-sheet-builder/credentials/token.json`が生成されます

4. 開発サーバーの起動:
   ```
   cd mcp-test-sheet-builder
   npm run build
   node build/index.js
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
6. 以下のコマンドでトークンを生成:
   ```
   node simple-token-generator.js
   ```

## 認証トークンの詳細な生成方法

Node.jsのバージョンによっては、トークン生成時に問題が発生する場合があります。以下の方法を使用してください：

### 方法1: シンプルなトークン生成スクリプトを使用（推奨）

プロジェクトルートにある`simple-token-generator.js`を使用します：

```
node simple-token-generator.js
```

このスクリプトは固定ポート（8080）を使用するため、ランダムポートで発生する問題を回避できます。表示されるURLをブラウザにコピー＆ペーストして認証を行ってください。

### 方法2: ビルド後のトークン生成スクリプトを使用

```
npm run build
node build/generate-token.js
```

### 方法3: 手動でトークンを生成

手動でトークンを生成する場合は、次のステップに従ってください：

1. Google Cloud Consoleで「認証情報」→「OAuth 2.0 クライアント ID」を選択
2. リダイレクトURIに `http://localhost:8080` を追加
3. シンプルなトークン生成スクリプトを実行:
   ```
   node simple-token-generator.js
   ```
4. 表示されるURLをブラウザで開いて認証する
5. 認証後、リダイレクトされた画面で「このウィンドウを閉じてください」と表示されたら成功です

## Node.js v17.3.0 でのトラブルシューティング

Node.js v17.3.0では`open`パッケージの最新バージョンと互換性がない問題があります。以下のエラーが発生した場合：

```
SyntaxError: The requested module 'node:fs/promises' does not provide an export named 'constants'
```

次の対策を行ってください：

1. 互換性のある古いバージョンの`open`パッケージをインストール：

```
npm uninstall open
npm install open@8.4.2
```

2. `simple-token-generator.js`が正常に動作しない場合は、ポート番号を変更してみてください：

```javascript
// ポート番号を3000から8080に変更
const PORT = 8080;
```

## ポート競合のトラブルシューティング

特定のポートが既に使用されている場合は、以下のエラーが表示されます：

```
Error: listen EADDRINUSE: address already in use :::8080
```

この場合、以下の対策を行ってください：

1. 使用中のポートを確認する：
   ```
   lsof -i:8080 | grep LISTEN
   ```

2. 使用中のプロセスを終了する：
   ```
   kill <PID>
   ```

3. または、`simple-token-generator.js`内のポート番号を変更する：
   ```javascript
   // 別のポート番号に変更（例: 9090）
   const PORT = 9090;
   ```

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