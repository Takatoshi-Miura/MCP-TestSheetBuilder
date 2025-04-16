import { google } from "googleapis";
import fs from "fs";
import path from "path";
import http from "http";
import url from "url";
import open from "open";
import { fileURLToPath } from "url";
import { dirname } from "path";
// ESM用にファイルパスを取得
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
// 認証情報ファイルのパスを設定
const CREDENTIALS_PATH = process.env.CREDENTIALS_PATH || path.join(__dirname, "../credentials/client_secret.json");
const TOKEN_PATH = process.env.TOKEN_PATH || path.join(__dirname, "../credentials/token.json");
// Google APIのスコープ設定
const SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
];
/**
 * 認証情報からOAuth2クライアントを作成
 */
export function createOAuth2Client() {
    try {
        if (!fs.existsSync(CREDENTIALS_PATH)) {
            console.error(`認証情報ファイルが見つかりません: ${CREDENTIALS_PATH}`);
            return null;
        }
        const credentialsContent = fs.readFileSync(CREDENTIALS_PATH, "utf8");
        const credentials = JSON.parse(credentialsContent);
        const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
        return new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
    }
    catch (error) {
        console.error("OAuth2クライアント作成エラー:", error);
        return null;
    }
}
/**
 * 認証用URLを生成
 */
export function getAuthUrl(oAuth2Client, redirectUri) {
    return oAuth2Client.generateAuthUrl({
        access_type: "offline",
        scope: SCOPES,
        redirect_uri: redirectUri,
        prompt: "consent" // トークンを強制的に再取得するためにconsentを指定
    });
}
/**
 * 認証コードからトークンを取得してファイルに保存
 */
export async function getTokenFromCode(oAuth2Client, code, redirectUri) {
    try {
        const { tokens } = await oAuth2Client.getToken({
            code: code,
            redirect_uri: redirectUri
        });
        oAuth2Client.setCredentials(tokens);
        // トークンをファイルに保存
        fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
        console.log(`トークンが保存されました: ${TOKEN_PATH}`);
    }
    catch (error) {
        console.error("トークン取得エラー:", error);
        throw error;
    }
}
/**
 * ブラウザでURLを開く（Node.js v17互換版）
 */
async function openBrowser(url) {
    try {
        await open(url);
    }
    catch (error) {
        console.error("ブラウザを開く際にエラーが発生しました:", error);
        console.log("手動で次のURLをブラウザで開いてください:");
        console.log(url);
    }
}
/**
 * ローカルサーバーを起動してOAuth認証を処理
 */
export function startLocalAuthServer(oAuth2Client) {
    return new Promise((resolve, reject) => {
        // ランダムなポートでローカルサーバーを作成
        const server = http.createServer(async (req, res) => {
            try {
                // リクエストURLからコードを取得
                const queryObject = url.parse(req.url || "", true).query;
                const code = queryObject.code;
                if (code) {
                    // 成功レスポンスを返す
                    res.writeHead(200, { "Content-Type": "text/html" });
                    res.end(`
            <html>
              <body>
                <h1>認証に成功しました！</h1>
                <p>このウィンドウを閉じて、アプリケーションに戻ってください。</p>
              </body>
            </html>
          `);
                    // サーバーを閉じる
                    server.close();
                    // サーバーのアドレスを取得
                    const address = server.address();
                    if (address && typeof address !== "string") {
                        const redirectUri = `http://localhost:${address.port}`;
                        // コードからトークンを取得
                        await getTokenFromCode(oAuth2Client, code, redirectUri);
                        resolve();
                    }
                    else {
                        reject(new Error("サーバーアドレスの取得に失敗しました"));
                    }
                }
                else {
                    // エラーレスポンスを返す
                    res.writeHead(400, { "Content-Type": "text/html" });
                    res.end(`
            <html>
              <body>
                <h1>認証に失敗しました。</h1>
                <p>認証コードが見つかりませんでした。もう一度やり直してください。</p>
              </body>
            </html>
          `);
                    reject(new Error("認証コードが見つかりませんでした"));
                }
            }
            catch (error) {
                console.error("認証処理エラー:", error);
                res.writeHead(500, { "Content-Type": "text/html" });
                res.end(`
          <html>
            <body>
              <h1>エラーが発生しました。</h1>
              <p>${error}</p>
            </body>
          </html>
        `);
                reject(error);
            }
        });
        // サーバーをランダムポートでリッスン
        server.listen(0, () => {
            const address = server.address();
            if (address && typeof address !== "string") {
                const port = address.port;
                const redirectUri = `http://localhost:${port}`;
                console.log(`ローカル認証サーバーが起動しました: ${redirectUri}`);
                // 認証URLを生成して開く
                const authUrl = getAuthUrl(oAuth2Client, redirectUri);
                console.log(`ブラウザで次のURLを開いて認証してください: ${authUrl}`);
                // ブラウザを開く
                openBrowser(authUrl);
            }
            else {
                reject(new Error("サーバーアドレスの取得に失敗しました"));
            }
        });
    });
}
/**
 * OAuth認証フローを実行してトークンを取得
 */
export async function authorize() {
    const oAuth2Client = createOAuth2Client();
    if (!oAuth2Client)
        return null;
    // トークンファイルがすでに存在するか確認
    if (fs.existsSync(TOKEN_PATH)) {
        try {
            const tokenContent = fs.readFileSync(TOKEN_PATH, "utf8");
            const token = JSON.parse(tokenContent);
            oAuth2Client.setCredentials(token);
            // トークンが有効期限切れかどうか確認
            const currentTime = Date.now();
            if (token.expiry_date && token.expiry_date > currentTime) {
                console.log("既存のトークンを使用します");
                return oAuth2Client;
            }
            else if (token.refresh_token) {
                // リフレッシュトークンがある場合は更新を試みる
                try {
                    console.log("トークンを更新します...");
                    const { credentials } = await oAuth2Client.refreshAccessToken();
                    oAuth2Client.setCredentials(credentials);
                    // 更新したトークンを保存
                    fs.writeFileSync(TOKEN_PATH, JSON.stringify(credentials));
                    console.log("トークンが更新されました");
                    return oAuth2Client;
                }
                catch (refreshError) {
                    console.error("トークン更新エラー:", refreshError);
                    console.log("新しいトークンを取得します...");
                }
            }
        }
        catch (error) {
            console.error("トークン読み込みエラー:", error);
            console.log("新しいトークンを取得します...");
        }
    }
    // 新しいトークンを取得するためにOAuth認証フローを開始
    console.log("OAuth認証フローを開始します...");
    try {
        await startLocalAuthServer(oAuth2Client);
        return oAuth2Client;
    }
    catch (error) {
        console.error("OAuth認証エラー:", error);
        return null;
    }
}
/**
 * トークン生成コマンドを実行する関数
 */
export async function generateToken() {
    console.log("Google API認証トークンを生成します...");
    try {
        const authClient = await authorize();
        if (authClient) {
            console.log("トークンの生成に成功しました！");
        }
        else {
            console.error("トークンの生成に失敗しました。");
        }
    }
    catch (error) {
        console.error("トークン生成エラー:", error);
    }
}
// コマンドライン引数でトークン生成が指定された場合に実行
if (process.argv.includes("--generate-token")) {
    generateToken().catch(console.error);
}
