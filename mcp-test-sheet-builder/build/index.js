import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { google } from "googleapis";
import { z } from "zod";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";
import { authorize } from "./auth.js";
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
function safeJsonParse(str) {
    if (!str)
        return null;
    try {
        // すでにオブジェクトの場合はそのまま返す
        if (typeof str === "object")
            return str;
        // 文字列の場合はパースする
        return JSON.parse(str);
    }
    catch (error) {
        console.error("JSONパースエラー:", error);
        return null;
    }
}
// Google認証用クライアントの取得
async function getAuthClient() {
    return authorize();
}
// プロンプトから因子と水準を生成する関数
function generateFactorsFromPrompt(prompt) {
    // この実装はとても単純化されています
    const factors = [];
    // プロンプトから基本的な因子を抽出する簡易処理の例
    if (prompt.includes("ブラウザ")) {
        factors.push({
            name: "ブラウザ",
            levels: ["Chrome", "Firefox", "Safari", "Edge"]
        });
    }
    if (prompt.includes("OS")) {
        factors.push({
            name: "OS",
            levels: ["Windows", "macOS", "Linux"]
        });
    }
    if (prompt.includes("画面サイズ") || prompt.includes("解像度")) {
        factors.push({
            name: "画面サイズ",
            levels: ["スマホ", "タブレット", "デスクトップ"]
        });
    }
    if (prompt.includes("ネットワーク")) {
        factors.push({
            name: "ネットワーク状態",
            levels: ["高速", "遅延あり", "オフライン"]
        });
    }
    // もし因子が見つからなかった場合、デフォルトの因子を追加
    if (factors.length === 0) {
        factors.push({
            name: "テスト環境",
            levels: ["開発環境", "テスト環境", "本番環境"]
        });
        factors.push({
            name: "ユーザー権限",
            levels: ["一般ユーザー", "管理者"]
        });
    }
    return factors;
}
// 全因子組み合わせテストケースを生成する関数
function generateAllCombinationsTestCases(factors) {
    if (factors.length === 0) {
        return [];
    }
    // 再帰的に組み合わせを生成
    const generateCombinations = (index, currentCombination) => {
        if (index === factors.length) {
            return [currentCombination];
        }
        const factor = factors[index];
        const result = [];
        for (const level of factor.levels) {
            const newCombination = { ...currentCombination };
            newCombination[factor.name] = level;
            result.push(...generateCombinations(index + 1, newCombination));
        }
        return result;
    };
    return generateCombinations(0, {});
}
// 直交表によるテストケースを生成する関数（簡易版）
function generateOrthogonalArrayTestCases(factors) {
    // 簡易版の実装。本来は複雑な直交表アルゴリズムが必要
    if (factors.length <= 2) {
        return generateAllCombinationsTestCases(factors);
    }
    // 代表的な水準を選択して組み合わせる簡易実装
    const testCases = [];
    const mainFactors = factors.slice(0, 2); // 最初の2つの因子は全組み合わせ
    const mainCombinations = generateAllCombinationsTestCases(mainFactors);
    const otherFactors = factors.slice(2);
    for (const combination of mainCombinations) {
        const testCase = { ...combination };
        // 残りの因子はランダムに水準を選択
        for (const factor of otherFactors) {
            const randomIndex = Math.floor(Math.random() * factor.levels.length);
            testCase[factor.name] = factor.levels[randomIndex];
        }
        testCases.push(testCase);
    }
    return testCases;
}
// 列番号をA1形式の列文字に変換する関数
function columnIndexToLetter(index) {
    let temp, letter = "";
    while (index > 0) {
        temp = (index - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        index = (index - temp - 1) / 26;
    }
    return letter;
}
// スプレッドシートの値を取得する関数
async function getSheetValues(auth, spreadsheetId, range) {
    const sheets = google.sheets({ version: "v4", auth });
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range,
        });
        return response.data.values || [];
    }
    catch (error) {
        console.error("スプレッドシートの値取得エラー:", error);
        throw error;
    }
}
// スプレッドシートの値を更新する関数
async function updateSheetValues(auth, spreadsheetId, range, values) {
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
    }
    catch (error) {
        console.error("スプレッドシートの値の更新エラー:", error);
        throw error;
    }
}
// テンプレートからスプレッドシートをコピーする関数
async function copySpreadsheet(auth, templateId, title) {
    const drive = google.drive({ version: "v3", auth });
    try {
        const response = await drive.files.copy({
            fileId: templateId,
            requestBody: {
                name: title,
            },
        });
        if (!response.data.id) {
            throw new Error("コピーしたファイルのIDが取得できませんでした");
        }
        return response.data.id;
    }
    catch (error) {
        console.error("スプレッドシートのコピーエラー:", error);
        throw error;
    }
}
// テストシートを作成する関数
async function createTestSheet(auth, templateId, title, factors, testCases) {
    // テンプレートをコピー
    const newSheetId = await copySpreadsheet(auth, templateId, title);
    // 因子と水準をシートに書き込む
    const factorValues = [
        ["因子", "水準"],
        ...factors.map(factor => [factor.name, factor.levels.join(", ")])
    ];
    await updateSheetValues(auth, newSheetId, "因子水準!A1:B" + (factors.length + 1), factorValues);
    // テストケースをシートに書き込む
    if (testCases.length > 0) {
        // ヘッダー行の作成
        const headers = ["No.", ...Object.keys(testCases[0]), "結果", "備考"];
        // データ行の作成
        const rows = testCases.map((testCase, index) => {
            return [
                (index + 1).toString(),
                ...Object.values(testCase),
                "", // 結果列
                "" // 備考列
            ];
        });
        const testCaseValues = [headers, ...rows];
        await updateSheetValues(auth, newSheetId, "テストケース!A1:" + columnIndexToLetter(headers.length) + (testCases.length + 1), testCaseValues);
    }
    return newSheetId;
}
// テストシート生成ツール
server.tool("generate-test", "テストシートを生成します", {
    templateId: z.string().describe("テンプレートとなるスプレッドシートのID"),
    title: z.string().describe("生成するスプレッドシートのタイトル"),
    prompt: z.string().describe("テスト要件を記述したプロンプト"),
    useOrthogonalArray: z.boolean().optional().describe("直交表を使用するかどうか（省略時はfalse）"),
}, async ({ templateId, title, prompt, useOrthogonalArray = false }) => {
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
        // プロンプトから因子と水準を生成
        const factors = generateFactorsFromPrompt(prompt);
        // テストケースの生成
        const testCases = useOrthogonalArray
            ? generateOrthogonalArrayTestCases(factors)
            : generateAllCombinationsTestCases(factors);
        // テストシートの作成
        const newSheetId = await createTestSheet(auth, templateId, title, factors, testCases);
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify({
                        sheetId: newSheetId,
                        sheetUrl: `https://docs.google.com/spreadsheets/d/${newSheetId}`,
                        factorCount: factors.length,
                        testCaseCount: testCases.length
                    }, null, 2)
                }
            ],
        };
    }
    catch (error) {
        return {
            content: [
                {
                    type: "text",
                    text: `テストシートの生成に失敗しました: ${error.message || String(error)}`
                }
            ],
            isError: true
        };
    }
});
// スプレッドシート情報取得ツール
server.tool("get-spreadsheet", "スプレッドシートの情報を取得します", {
    id: z.string().describe("スプレッドシートのID"),
    range: z.string().optional().describe("取得する範囲（例: Sheet1!A1:Z100）"),
}, async ({ id, range = "Sheet1!A1:Z100" }) => {
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
        const values = await getSheetValues(auth, id, range);
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify({ values }, null, 2)
                }
            ],
        };
    }
    catch (error) {
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
});
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
    }
    catch (error) {
        console.error("初期化エラー:", error);
        process.exit(1);
    }
}
main().catch((error) => {
    console.error("main()での致命的なエラー:", error);
    process.exit(1);
});
