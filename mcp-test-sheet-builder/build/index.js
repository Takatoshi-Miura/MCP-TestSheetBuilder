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
        // テンプレートIDが提供されていない場合はエラーを返す
        if (!templateId || templateId.trim() === '') {
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "error",
                            message: "テンプレートとなるスプレッドシートが提供されていません。テンプレートのIDを提供してください。"
                        }, null, 2)
                    }
                ],
                isError: true
            };
        }
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
        // テンプレートIDの有効性を確認
        try {
            const sheets = google.sheets({ version: "v4", auth });
            await sheets.spreadsheets.get({
                spreadsheetId: templateId
            });
        }
        catch (error) {
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "error",
                            message: "提供されたテンプレートIDが無効です。有効なスプレッドシートIDを提供してください。"
                        }, null, 2)
                    }
                ],
                isError: true
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
        // テンプレートIDが提供されていない場合はエラーを返す
        if (!sourceId || sourceId.trim() === '') {
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "error",
                            message: "テンプレートとなるスプレッドシートが提供されていません。テンプレートのIDを提供してください。"
                        }, null, 2)
                    }
                ],
                isError: true
            };
        }
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
        // テンプレートIDの有効性を確認
        try {
            const sheets = google.sheets({ version: "v4", auth });
            await sheets.spreadsheets.get({
                spreadsheetId: sourceId
            });
        }
        catch (error) {
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "error",
                            message: "提供されたテンプレートIDが無効です。有効なスプレッドシートIDを提供してください。"
                        }, null, 2)
                    }
                ],
                isError: true
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
function generateFeatureFactorsFromPrompt(prompt) {
    const featureFactors = [];
    // バッジ表示に関する因子・水準
    if (prompt.includes('バッジ') || prompt.includes('件数表示') || prompt.includes('未承認伝票')) {
        featureFactors.push({
            name: 'バッジ表示テスト',
            factors: [
                {
                    name: '対象画面',
                    levels: ['未承認一覧画面', 'ワークフロー一覧画面']
                },
                {
                    name: '伝票件数',
                    levels: ['0件', '1件', '98件', '99件', '100件以上']
                },
                {
                    name: 'バッジ表示',
                    levels: ['非表示', '件数表示(1-98)', '99表示', '99+表示']
                },
                {
                    name: 'デバイス',
                    levels: ['PC', 'タブレット', 'スマートフォン']
                },
                {
                    name: '画面サイズ',
                    levels: ['大', '中', '小']
                }
            ]
        });
    }
    // 画面に関する因子・水準
    if (prompt.includes('画面') || prompt.includes('ページ') || prompt.includes('UI')) {
        featureFactors.push({
            name: '画面テスト',
            factors: [
                {
                    name: '画面',
                    levels: ['一覧画面', '詳細画面', '編集画面', '登録画面']
                },
                {
                    name: 'コンポーネント',
                    levels: ['ヘッダー', 'フッター', 'サイドメニュー', 'メインコンテンツ']
                },
                {
                    name: '表示状態',
                    levels: ['初期表示', 'データあり', 'データなし', 'エラー表示']
                }
            ]
        });
    }
    // ワークフローに関する因子・水準
    if (prompt.includes('ワークフロー') || prompt.includes('承認') || prompt.includes('申請')) {
        featureFactors.push({
            name: 'ワークフローテスト',
            factors: [
                {
                    name: '画面',
                    levels: ['未承認一覧画面', 'ワークフロー一覧画面', '承認詳細画面', '申請画面']
                },
                {
                    name: '承認ステータス',
                    levels: ['未申請', '申請中', '承認済み', '差戻し', '却下']
                },
                {
                    name: 'ユーザー権限',
                    levels: ['申請者', '承認者', '管理者']
                }
            ]
        });
    }
    // データ操作に関する因子・水準
    if (prompt.includes('データ') || prompt.includes('登録') || prompt.includes('編集') || prompt.includes('削除')) {
        featureFactors.push({
            name: 'データ操作テスト',
            factors: [
                {
                    name: '操作種別',
                    levels: ['登録', '編集', '削除', '参照']
                },
                {
                    name: 'データ状態',
                    levels: ['新規データ', '既存データ', '不正データ']
                },
                {
                    name: '権限',
                    levels: ['一般ユーザー', '管理者']
                }
            ]
        });
    }
    // 検索機能に関する因子・水準
    if (prompt.includes('検索') || prompt.includes('フィルター') || prompt.includes('ソート')) {
        featureFactors.push({
            name: '検索機能テスト',
            factors: [
                {
                    name: '検索条件',
                    levels: ['キーワード検索', '詳細検索', '空検索', '該当なし検索']
                },
                {
                    name: 'ソート条件',
                    levels: ['昇順', '降順', 'なし']
                },
                {
                    name: 'データ量',
                    levels: ['少量', '大量', 'なし']
                }
            ]
        });
    }
    // 帳票出力に関する因子・水準
    if (prompt.includes('帳票') || prompt.includes('出力') || prompt.includes('印刷') || prompt.includes('PDF')) {
        featureFactors.push({
            name: '帳票出力テスト',
            factors: [
                {
                    name: '帳票種類',
                    levels: ['一覧表', '詳細表', 'サマリーレポート']
                },
                {
                    name: '出力形式',
                    levels: ['PDF', 'Excel', 'CSV']
                },
                {
                    name: 'データ量',
                    levels: ['1件', '複数件', '大量データ']
                }
            ]
        });
    }
    // 何も検出できなかった場合のデフォルト
    if (featureFactors.length === 0) {
        featureFactors.push({
            name: '基本機能テスト',
            factors: [
                {
                    name: '画面',
                    levels: ['一覧画面', '詳細画面', '編集画面']
                },
                {
                    name: '操作',
                    levels: ['参照', '登録', '編集', '削除']
                },
                {
                    name: 'ユーザー',
                    levels: ['一般ユーザー', '管理者']
                }
            ]
        });
    }
    return featureFactors;
}
// 機能ごとの因子・水準生成ツール
server.tool("generate_feature_factors", "与えられた要件に基づいて機能ごとの因子・水準を生成し、テンプレートシートに記載します", {
    templateId: z.string().describe("テンプレートとなるスプレッドシートのID"),
    title: z.string().optional().describe("スプレッドシートのタイトル（指定しない場合は元のタイトルを維持）"),
    prompt: z.string().describe("テスト要件を記述したプロンプト"),
    folderId: z.string().optional().describe("生成したシートを保存するフォルダID"),
    confirmed: z.boolean().optional().describe("因子・水準の確認結果（true: 確認済みで記載実行、false: 確認前または未確認、undefined: 初回呼び出し）"),
    featureFactorsJson: z.string().optional().describe("確認された因子・水準のJSON文字列（confirmedがtrueの場合に使用）"),
}, async ({ templateId, title, prompt, folderId, confirmed, featureFactorsJson }) => {
    try {
        // テンプレートIDが提供されていない場合はエラーを返す
        if (!templateId || templateId.trim() === '') {
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "error",
                            message: "テンプレートとなるスプレッドシートが提供されていません。テンプレートのIDを提供してください。"
                        }, null, 2)
                    }
                ],
                isError: true
            };
        }
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
        // テンプレートIDの有効性を確認
        let spreadsheetInfo;
        try {
            const sheets = google.sheets({ version: "v4", auth });
            const response = await sheets.spreadsheets.get({
                spreadsheetId: templateId
            });
            spreadsheetInfo = response.data;
        }
        catch (error) {
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "error",
                            message: "提供されたテンプレートIDが無効です。有効なスプレッドシートIDを提供してください。"
                        }, null, 2)
                    }
                ],
                isError: true
            };
        }
        // タイトル変更が必要な場合のみコピーする
        let targetSheetId = templateId;
        if (title) {
            // タイトルが指定されている場合のみシートをコピー
            targetSheetId = await copySpreadsheet(auth, templateId, title);
            // フォルダが指定されていればそのフォルダに移動
            if (folderId) {
                const isValidFolder = await validateFolderId(auth, folderId);
                if (isValidFolder) {
                    await moveFileToFolder(auth, targetSheetId, folderId);
                }
                else {
                    return {
                        content: [
                            {
                                type: "text",
                                text: JSON.stringify({
                                    status: "partial_success",
                                    sheetId: targetSheetId,
                                    sheetUrl: `https://docs.google.com/spreadsheets/d/${targetSheetId}`,
                                    message: "因子・水準シートを作成しましたが、指定されたフォルダIDが無効なため移動できませんでした。"
                                }, null, 2)
                            }
                        ]
                    };
                }
            }
        }
        // 確認済みの場合は記載処理を実行
        if (confirmed === true && featureFactorsJson) {
            // JSON文字列からオブジェクトに戻す
            const parsedFeatureFactors = safeJsonParse(featureFactorsJson);
            if (!parsedFeatureFactors) {
                return {
                    content: [
                        {
                            type: "text",
                            text: JSON.stringify({
                                status: "error",
                                message: "因子・水準データの形式が不正です。"
                            }, null, 2)
                        }
                    ],
                    isError: true
                };
            }
            // 因子・水準シートが存在するか確認し、なければ作成
            const hasFactorSheet = await checkSheetExists(auth, targetSheetId, "因子・水準");
            if (!hasFactorSheet) {
                await addSheet(auth, targetSheetId, "因子・水準");
                // 初期ヘッダー設定
                await updateSheetValues(auth, targetSheetId, "因子・水準!A1:B1", [["因子", "水準"]]);
            }
            // 現在のシートの内容を確認
            const currentValues = await getSheetValues(auth, targetSheetId, "因子・水準!A:Z");
            // 既存の因子・水準テーブルを探す
            let startRow = -1;
            let headerRow = -1;
            // 「因子」「水準」ヘッダーを探す
            for (let i = 0; i < currentValues.length; i++) {
                if (currentValues[i] &&
                    currentValues[i][0] === '因子' &&
                    currentValues[i][1] === '水準') {
                    headerRow = i;
                    startRow = i + 1; // ヘッダーの次の行から書き始める
                    break;
                }
            }
            // 既存のテーブルが見つからない場合は新規作成
            if (startRow === -1) {
                // スペースを空けて最終行の後に記載
                startRow = currentValues.length + 2;
                headerRow = startRow - 1;
                // 因子・水準テーブルのヘッダー
                await updateSheetValues(auth, targetSheetId, `因子・水準!A${headerRow}:Z${headerRow}`, [['因子', '水準', '', '', '', '', '', '', '', '', '', '', '', '', '']]);
            }
            // 既存の表に空の行がある場合はそこから書き始める
            let writeRow = startRow;
            for (let i = startRow; i < currentValues.length; i++) {
                if (!currentValues[i] || !currentValues[i][0]) {
                    writeRow = i;
                    break;
                }
                // 既存のデータの最後まで走査して空行がなければ、最後の行の次から始める
                if (i === currentValues.length - 1) {
                    writeRow = currentValues.length;
                }
            }
            // 機能ごとの表を書き込み
            let lastRow = writeRow;
            // 各機能ごとに因子・水準表を書き込む
            for (const feature of parsedFeatureFactors) {
                // 機能名を書き込む（既存のテーブルの一部として追加するため、ヘッダーは書き込まない）
                // 各因子と水準を書き込む
                for (const factor of feature.factors) {
                    // 水準を横に並べて表示するために配列を作成
                    const rowData = [factor.name];
                    // 水準を配列に追加
                    factor.levels.forEach((level) => {
                        rowData.push(level);
                    });
                    // 足りない部分を空白で埋める
                    while (rowData.length < 15) { // 列数に応じて調整
                        rowData.push('');
                    }
                    // シートに書き込む
                    await updateSheetValues(auth, targetSheetId, `因子・水準!A${lastRow}:O${lastRow}`, [rowData]);
                    lastRow++;
                }
            }
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "success",
                            sheetId: targetSheetId,
                            sheetUrl: `https://docs.google.com/spreadsheets/d/${targetSheetId}`,
                            message: "クライアント確認後、因子・水準シートに記載しました。"
                        }, null, 2)
                    }
                ]
            };
        }
        else if (confirmed === false) {
            // 確認が拒否された場合
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "cancelled",
                            message: "因子・水準の記載をキャンセルしました。",
                            sheetId: targetSheetId,
                            sheetUrl: `https://docs.google.com/spreadsheets/d/${targetSheetId}`
                        }, null, 2)
                    }
                ]
            };
        }
        else {
            // 初回呼び出しの場合は因子・水準を生成して確認を求める
            // プロンプトから機能ごとの因子・水準を抽出
            const featureFactors = generateFeatureFactorsFromPrompt(prompt);
            // クライアントに因子・水準の確認を求める
            let confirmationMessage = "以下の因子・水準を生成しました。シートに追加してよろしいですか？\n\n";
            for (const feature of featureFactors) {
                confirmationMessage += `▼ ${feature.name}\n`;
                for (const factor of feature.factors) {
                    confirmationMessage += `・${factor.name}: ${factor.levels.join(', ')}\n`;
                }
                confirmationMessage += "\n";
            }
            confirmationMessage += "この因子・水準で問題なければ、シートに記載します。";
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify({
                            status: "confirmation_required",
                            sheetId: targetSheetId,
                            sheetUrl: `https://docs.google.com/spreadsheets/d/${targetSheetId}`,
                            featureFactors: featureFactors,
                            message: confirmationMessage,
                            action: "confirm_factors"
                        }, null, 2)
                    }
                ]
            };
        }
    }
    catch (error) {
        return {
            content: [
                {
                    type: "text",
                    text: `因子・水準シートの生成に失敗しました: ${error.message || String(error)}`
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
