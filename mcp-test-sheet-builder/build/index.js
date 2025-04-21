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
    // プロンプトから因子と水準を抽出
    const factors = [];
    // OS関連の因子
    if (prompt.includes("OS") || prompt.includes("オペレーティングシステム")) {
        factors.push({
            name: "OS",
            levels: ["iOS", "Android"]
        });
    }
    // アプリバージョン関連
    if (prompt.includes("バージョン") || prompt.includes("version")) {
        factors.push({
            name: "アプリバージョン",
            levels: ["最新版", "旧バージョン"]
        });
    }
    // 端末関連
    if (prompt.includes("端末") || prompt.includes("デバイス") || prompt.includes("機種")) {
        factors.push({
            name: "端末",
            levels: ["iPhone", "iPad", "Android Phone", "Android Tablet"]
        });
    }
    // ネットワーク状態
    if (prompt.includes("ネットワーク") || prompt.includes("通信")) {
        factors.push({
            name: "ネットワーク状態",
            levels: ["Wi-Fi", "モバイルデータ通信", "オフライン"]
        });
    }
    // ユーザー状態
    if (prompt.includes("ユーザー") || prompt.includes("ログイン")) {
        factors.push({
            name: "ユーザー状態",
            levels: ["ログイン済み", "未ログイン"]
        });
    }
    // 画面向き
    if (prompt.includes("画面向き") || prompt.includes("orientation")) {
        factors.push({
            name: "画面向き",
            levels: ["縦向き", "横向き"]
        });
    }
    // もし因子が見つからなかった場合、デフォルトの因子を追加
    if (factors.length === 0) {
        factors.push({
            name: "OS",
            levels: ["iOS", "Android"]
        });
        factors.push({
            name: "端末",
            levels: ["スマートフォン", "タブレット"]
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
// 直交表によるテストケースを生成する関数
function generateOrthogonalArrayTestCases(factors) {
    // 直交表L4(2^3)の定義: 3因子2水準用
    const L4 = [
        [0, 0, 0],
        [0, 1, 1],
        [1, 0, 1],
        [1, 1, 0]
    ];
    // 直交表L8(2^7)の定義: 7因子2水準用
    const L8 = [
        [0, 0, 0, 0, 0, 0, 0],
        [0, 0, 0, 1, 1, 1, 1],
        [0, 1, 1, 0, 0, 1, 1],
        [0, 1, 1, 1, 1, 0, 0],
        [1, 0, 1, 0, 1, 0, 1],
        [1, 0, 1, 1, 0, 1, 0],
        [1, 1, 0, 0, 1, 1, 0],
        [1, 1, 0, 1, 0, 0, 1]
    ];
    // 2水準に変換（水準が3以上ある場合は先頭2つのみ使用）
    const factorsWith2Levels = factors.map(factor => {
        return {
            name: factor.name,
            levels: factor.levels.length >= 2 ? [factor.levels[0], factor.levels[1]] : factor.levels
        };
    });
    // 因子数に基づいて適切な直交表を選択
    let orthogonalArray;
    let usedFactors;
    if (factorsWith2Levels.length <= 3) {
        orthogonalArray = L4;
        usedFactors = factorsWith2Levels.slice(0, 3);
    }
    else if (factorsWith2Levels.length <= 7) {
        orthogonalArray = L8;
        usedFactors = factorsWith2Levels.slice(0, 7);
    }
    else {
        // 因子が多すぎる場合は全組み合わせの一部を使用
        return generateAllCombinationsTestCases(factors).slice(0, 8);
    }
    // 直交表からテストケースを生成
    const testCases = [];
    for (const row of orthogonalArray) {
        const testCase = {};
        for (let i = 0; i < usedFactors.length; i++) {
            const factor = usedFactors[i];
            if (i < row.length && factor.levels.length > row[i]) {
                testCase[factor.name] = factor.levels[row[i]];
            }
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
// シートが存在するか確認する関数
async function checkSheetExists(auth, spreadsheetId, sheetName) {
    const sheets = google.sheets({ version: "v4", auth });
    try {
        const response = await sheets.spreadsheets.get({
            spreadsheetId
        });
        if (response.data.sheets) {
            return response.data.sheets.some(sheet => sheet.properties?.title === sheetName);
        }
        return false;
    }
    catch (error) {
        console.error("シート存在確認エラー:", error);
        throw error;
    }
}
// シートを追加する関数
async function addSheet(auth, spreadsheetId, sheetName) {
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
    }
    catch (error) {
        console.error("シート追加エラー:", error);
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
// 指定されたフォルダにファイルを移動する関数
async function moveFileToFolder(auth, fileId, folderId) {
    const drive = google.drive({ version: "v3", auth });
    try {
        // 現在の親フォルダを取得
        const file = await drive.files.get({
            fileId: fileId,
            fields: 'parents'
        });
        // 親フォルダのリスト
        const previousParents = file.data.parents?.join(',') || '';
        // ファイルを新しいフォルダに移動（元の親から削除して新しい親を追加）
        await drive.files.update({
            fileId: fileId,
            addParents: folderId,
            removeParents: previousParents,
            fields: 'id, parents'
        });
    }
    catch (error) {
        console.error("ファイル移動エラー:", error);
        throw error;
    }
}
// フォルダIDの有効性を確認する関数
async function validateFolderId(auth, folderId) {
    const drive = google.drive({ version: "v3", auth });
    try {
        const response = await drive.files.get({
            fileId: folderId,
            fields: 'mimeType'
        });
        // Google DriveのフォルダのmimeTypeを確認
        return response.data.mimeType === 'application/vnd.google-apps.folder';
    }
    catch (error) {
        console.error("フォルダID検証エラー:", error);
        return false;
    }
}
// 必要なシートを確認・作成する関数
async function ensureRequiredSheets(auth, spreadsheetId) {
    // テスト項目シートの確認と作成
    const testSheetExists = await checkSheetExists(auth, spreadsheetId, "テスト項目");
    if (!testSheetExists) {
        await addSheet(auth, spreadsheetId, "テスト項目");
        // テスト項目シートの初期化
        await updateSheetValues(auth, spreadsheetId, "テスト項目!A1:G2", [
            ["№", "テスト項目", "入力値・前提条件など", "操作など", "想定される結果", "テスター１：アプリVer：OS：機種：", "【再テスト用】テスター１：アプリVer：OS：機種："],
            ["実施日", "判定", "不具合チケットNO", "実施日", "判定", "", ""]
        ]);
    }
    // 因子・水準シートの確認と作成
    const factorSheetExists = await checkSheetExists(auth, spreadsheetId, "因子・水準");
    if (!factorSheetExists) {
        await addSheet(auth, spreadsheetId, "因子・水準");
        // 因子・水準シートの初期化
        await updateSheetValues(auth, spreadsheetId, "因子・水準!A1:B1", [
            ["因子", "水準"]
        ]);
    }
}
// テストシートを作成する関数
async function createTestSheet(auth, templateId, title, prompt, useOrthogonalArray) {
    // テンプレートをコピー
    const newSheetId = await copySpreadsheet(auth, templateId, title);
    // 必要なシートの確認と作成
    await ensureRequiredSheets(auth, newSheetId);
    // プロンプトから因子と水準を生成
    const factors = generateFactorsFromPrompt(prompt);
    // 因子と水準をシートに書き込む
    const factorValues = [
        ["因子", "水準"],
        ...factors.map(factor => [factor.name, factor.levels.join(", ")])
    ];
    await updateSheetValues(auth, newSheetId, `因子・水準!A1:B${factors.length + 1}`, factorValues);
    // テストケースの生成
    const testCases = useOrthogonalArray
        ? generateOrthogonalArrayTestCases(factors)
        : generateAllCombinationsTestCases(factors);
    // テストケースをシートに書き込む
    if (testCases.length > 0) {
        const rows = testCases.map((testCase, index) => {
            const testItem = Object.entries(testCase).map(([key, val]) => `${key}: ${val}`).join("\n");
            return [
                (index + 1).toString(), // No.
                `${title}のテスト ${index + 1}`, // テスト項目
                testItem, // 入力値・前提条件など
                "", // 操作など
                "", // 想定される結果
                "", // テスター1
                "" // 再テスト用
            ];
        });
        await updateSheetValues(auth, newSheetId, `テスト項目!A3:G${testCases.length + 3}`, rows);
    }
    return newSheetId;
}
// テストシート生成ツール
server.tool("generate_test", "テストシートを生成します", {
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
        // テストシートの作成
        const newSheetId = await createTestSheet(auth, templateId, title, prompt, useOrthogonalArray);
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify({
                        status: "success",
                        sheetId: newSheetId,
                        sheetUrl: `https://docs.google.com/spreadsheets/d/${newSheetId}`,
                        message: "テストシートを作成しました。"
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
// スプレッドシートコピーツール
server.tool("copy_spreadsheet", "指定したスプレッドシートを単純にコピーするだけのツールです", {
    sourceId: z.string().describe("コピー元のスプレッドシートのID"),
    title: z.string().describe("コピー先のスプレッドシートのタイトル"),
    folderId: z.string().optional().describe("コピー先のフォルダID（省略時は同じ階層にコピー）"),
}, async ({ sourceId, title, folderId }) => {
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
        // スプレッドシートをコピー
        const newSheetId = await copySpreadsheet(auth, sourceId, title);
        // フォルダIDが指定されていれば、そのフォルダにファイルを移動
        if (folderId) {
            // フォルダIDの有効性を確認
            const isValidFolder = await validateFolderId(auth, folderId);
            if (!isValidFolder) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                status: "partial_success",
                                sheetId: newSheetId,
                                sheetUrl: `https://docs.google.com/spreadsheets/d/${newSheetId}`,
                                message: "スプレッドシートをコピーしましたが、指定されたフォルダIDが無効なため移動できませんでした。"
                            }, null, 2)
                        }
                    ],
                };
            }
            // ファイルを指定されたフォルダに移動
            await moveFileToFolder(auth, newSheetId, folderId);
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "success",
                            sheetId: newSheetId,
                            sheetUrl: `https://docs.google.com/spreadsheets/d/${newSheetId}`,
                            folderId: folderId,
                            message: "スプレッドシートをコピーして指定されたフォルダに移動しました。"
                        }, null, 2)
                    }
                ],
            };
        }
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify({
                        status: "success",
                        sheetId: newSheetId,
                        sheetUrl: `https://docs.google.com/spreadsheets/d/${newSheetId}`,
                        message: "スプレッドシートをコピーしました。"
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
                    text: `スプレッドシートのコピーに失敗しました: ${error.message || String(error)}`
                }
            ],
            isError: true
        };
    }
});
// スプレッドシート情報取得ツール
server.tool("get_spreadsheet", "スプレッドシートの情報を取得します", {
    id: z.string().describe("スプレッドシートのID"),
    range: z.string().optional().describe("取得する範囲（例: Sheet1!A1:Z100）"),
}, async ({ id, range }) => {
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
