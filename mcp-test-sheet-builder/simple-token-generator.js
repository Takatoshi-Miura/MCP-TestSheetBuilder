import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';
import { google } from 'googleapis';
import http from 'http';
import url from 'url';

// ファイルパスを設定
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const CREDENTIALS_PATH = path.join(__dirname, 'credentials/client_secret.json');
const TOKEN_PATH = path.join(__dirname, 'credentials/token.json');

// スコープを設定
const SCOPES = [
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/spreadsheets'
];

// メイン関数
async function main() {
  console.log("Google API トークン生成ツール（シンプル版）");
  console.log("=======================================");

  try {
    // クライアント認証情報の読み込み
    if (!fs.existsSync(CREDENTIALS_PATH)) {
      console.error(`認証情報ファイルが見つかりません: ${CREDENTIALS_PATH}`);
      process.exit(1);
    }

    const credentialsContent = fs.readFileSync(CREDENTIALS_PATH, 'utf8');
    const credentials = JSON.parse(credentialsContent);
    const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;

    // OAuth2クライアントを作成
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

    // ローカルサーバーを作成して認証処理
    const server = http.createServer();
    
    // ポートを指定して起動（別のポート番号を使用）
    const PORT = 8080;
    
    server.on('request', async (req, res) => {
      try {
        const queryObject = url.parse(req.url || '', true).query;
        const code = queryObject.code;
        
        if (code) {
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(`
            <html><body>
              <h1>認証に成功しました！</h1>
              <p>このウィンドウを閉じてアプリケーションに戻ってください。</p>
            </body></html>
          `);
          
          console.log("認証コードを取得しました。トークンを生成中...");
          
          try {
            // リダイレクトURIをlocalhostの固定ポートに設定
            const redirectUri = `http://localhost:${PORT}`;
            
            // コードからトークンを取得
            const { tokens } = await oAuth2Client.getToken({
              code: code,
              redirect_uri: redirectUri
            });
            
            // トークンをファイルに保存
            fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
            console.log(`トークンが保存されました: ${TOKEN_PATH}`);
          } catch (tokenError) {
            console.error("トークン取得エラー:", tokenError);
          }
          
          // サーバーを閉じる
          server.close();
          process.exit(0);
        } else {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end('認証コードが見つかりませんでした');
        }
      } catch (error) {
        console.error("リクエスト処理エラー:", error);
        res.writeHead(500, { 'Content-Type': 'text/html' });
        res.end(`エラーが発生しました: ${error.message}`);
      }
    });

    server.listen(PORT, () => {
      // リダイレクトURIを設定
      const redirectUri = `http://localhost:${PORT}`;
      
      // 認証URLを生成
      const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
        redirect_uri: redirectUri,
        prompt: 'consent'
      });
      
      console.log(`\n以下のURLをブラウザにコピー＆ペーストして開いてください:\n\n${authUrl}\n`);
      console.log(`認証が完了するとhttp://localhost:${PORT}にリダイレクトされます。`);
      console.log("サーバーがリクエストを待機中...");
    });

  } catch (error) {
    console.error("エラーが発生しました:", error);
    process.exit(1);
  }
}

// 実行
main(); 