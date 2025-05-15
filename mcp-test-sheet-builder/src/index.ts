import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { google } from "googleapis";
import { OAuth2Client } from "google-auth-library";
import { z } from "zod";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import { authorize } from "./auth.js";
import { TestSheetBuilder } from "./services/TestSheetBuilder.js";
import { GoogleSheetService } from "./services/GoogleSheetService.js";

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

// Create server instance
const server = new McpServer({
  name: "mcp-test-sheet-builder",
  version: "1.0.0",
  capabilities: {
    resources: {},
    tools: {},
  },
});

// JSONを安全にパースする関数
function safeJsonParse(str: string): any {
  if (!str) return null;
  
  try {
    // すでにオブジェクトの場合はそのまま返す
    if (typeof str === "object") return str;
    
    // 文字列の場合はパースする
    return JSON.parse(str);
  } catch (error) {
    console.error("JSONパースエラー:", error);
    return null;
  }
}

// Google認証用クライアントの取得
async function getAuthClient(): Promise<OAuth2Client | null> {
  return authorize();
}

// TestSheetBuilder インスタンスを取得
async function getTestSheetBuilder(): Promise<TestSheetBuilder | null> {
  try {
    const credentials = safeJsonParse(fs.readFileSync(CREDENTIALS_PATH, "utf8"));
    const token = safeJsonParse(fs.readFileSync(TOKEN_PATH, "utf8"));
    
    if (!credentials || !token) {
      return null;
    }
    
    const googleSheetService = new GoogleSheetService(credentials, token);
    return new TestSheetBuilder(googleSheetService);
  } catch (error) {
    console.error("TestSheetBuilder の初期化エラー:", error);
    return null;
  }
}

// 列番号をA1形式の列文字に変換する関数
function columnIndexToLetter(index: number): string {
  let temp, letter = "";
  while (index > 0) {
    temp = (index - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    index = (index - temp - 1) / 26;
  }
  return letter;
}

// スプレッドシートの値を取得する関数
async function getSheetValues(auth: OAuth2Client, spreadsheetId: string, range: string): Promise<any[][]> {
  const sheets = google.sheets({ version: "v4", auth });
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });
    
    return response.data.values || [];
  } catch (error) {
    console.error("スプレッドシートの値取得エラー:", error);
    throw error;
  }
}

async function updateSheetValues(
  auth: OAuth2Client,
  spreadsheetId: string,
  range: string,
  values: any[][]
): Promise<void> {
  const sheets = google.sheets({ version: "v4", auth });
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values,
      },
    });
  } catch (error) {
    console.error("スプレッドシートの値更新エラー:", error);
    throw error;
  }
}

async function checkSheetExists(auth: OAuth2Client, spreadsheetId: string, sheetName: string): Promise<boolean> {
  const sheets = google.sheets({ version: "v4", auth });
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
    });
    
    return response.data.sheets?.some(sheet => sheet.properties?.title === sheetName) || false;
  } catch (error) {
    console.error("シート存在確認エラー:", error);
    throw error;
  }
}

async function addSheet(auth: OAuth2Client, spreadsheetId: string, sheetName: string): Promise<void> {
  const sheets = google.sheets({ version: "v4", auth });
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetName
              }
            }
          }
        ]
      }
    });
  } catch (error) {
    console.error("シート追加エラー:", error);
    throw error;
  }
}

async function ensureRequiredSheets(auth: OAuth2Client, spreadsheetId: string): Promise<void> {
  // 必要なシートの一覧
  const requiredSheets = ["テストケース", "因子水準"];
  
  for (const sheetName of requiredSheets) {
    const exists = await checkSheetExists(auth, spreadsheetId, sheetName);
    if (!exists) {
      await addSheet(auth, spreadsheetId, sheetName);
    }
  }
}

// スプレッドシート情報取得ツール
server.tool(
  "get_spreadsheet",
  "スプレッドシートの情報を取得します",
  {
    id: z.string().describe("スプレッドシートのID"),
    range: z.string().optional().describe("取得する範囲（例: Sheet1!A1:Z100）"),
  },
  async ({ id, range }) => {
    try {
      const auth = await getAuthClient();
      if (!auth) {
        return {
          content: [
            {
              type: "text",
              text: "Google認証に失敗しました。認証情報とトークンを確認してください。",
            },
          ],
        };
      }

      if (!range) {
        // シート情報を取得
        const sheets = google.sheets({ version: "v4", auth });
        const sheetInfo = await sheets.spreadsheets.get({
          spreadsheetId: id
        });
        
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                status: "success",
                sheetInfo: sheetInfo.data
              }, null, 2)
            }
          ],
        };
      }

      const values = await getSheetValues(auth, id, range);
      
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              status: "success", 
              values
            }, null, 2)
          }
        ],
      };
    } catch (error: any) {
      return {
        content: [
          {
            type: "text",
            text: `スプレッドシートの取得に失敗しました: ${error.message || String(error)}`
          }
        ],
        isError: true
      };
    }
  }
);

// テスト項目生成ツール
server.tool(
  "generate_test_items",
  "因子・水準シートからテスト項目を自動生成するツールです",
  {
    spreadsheetId: z.string().describe("テスト項目を生成するスプレッドシートのID"),
  },
  async ({ spreadsheetId }) => {
    try {
      // スプレッドシートIDが提供されていない場合はエラーを返す
      if (!spreadsheetId || spreadsheetId.trim() === '') {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                status: "error",
                message: "スプレッドシートIDが提供されていません。有効なスプレッドシートIDを提供してください。"
              }, null, 2)
            }
          ],
          isError: true
        };
      }

      // TestSheetBuilder インスタンスを取得
      const testSheetBuilder = await getTestSheetBuilder();
      if (!testSheetBuilder) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                status: "error",
                message: "TestSheetBuilder の初期化に失敗しました。認証情報とトークンを確認してください。"
              }, null, 2)
            }
          ],
          isError: true
        };
      }

      // テスト項目を生成
      const result = await testSheetBuilder.generateTestItemsFromFactorsAndLevels(spreadsheetId);

      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              status: result.success ? "success" : "error",
              message: result.message
            }, null, 2)
          }
        ],
        isError: !result.success
      };
    } catch (error) {
      console.error("テスト項目生成エラー:", error);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              status: "error",
              message: `テスト項目の生成中にエラーが発生しました: ${error}`
            }, null, 2)
          }
        ],
        isError: true
      };
    }
  }
);

// メイン関数
async function main() {
  try {
    // 認証情報とトークンの存在確認（起動時のみ）
    if (!fs.existsSync(CREDENTIALS_PATH)) {
      console.error(`認証情報ファイルが見つかりません: ${CREDENTIALS_PATH}`);
      process.exit(1);
    }
    if (!fs.existsSync(TOKEN_PATH)) {
      console.error(`トークンファイルが見つかりません: ${TOKEN_PATH}`);
      process.exit(1);
    }

    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("MCP Test Sheet Builder サーバーが標準入出力で実行中");
  } catch (error) {
    console.error("初期化エラー:", error);
    process.exit(1);
  }
}

main().catch((error) => {
  console.error("main()での致命的なエラー:", error);
  process.exit(1);
}); 