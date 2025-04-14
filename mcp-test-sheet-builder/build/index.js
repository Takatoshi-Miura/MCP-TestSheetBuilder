"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const GoogleSheetService_1 = require("./services/GoogleSheetService");
const TestSheetBuilder_1 = require("./services/TestSheetBuilder");
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const googleapis_1 = require("googleapis");
const readline_1 = __importDefault(require("readline"));
const http_1 = __importDefault(require("http"));
// 認証情報ファイルのパスを設定
const CREDENTIALS_PATH = process.env.CREDENTIALS_PATH || path_1.default.join(__dirname, '../credentials/client_secret.json');
const TOKEN_PATH = process.env.TOKEN_PATH || path_1.default.join(__dirname, '../credentials/token.json');
// Google APIのスコープ設定
const SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
];
/**
 * トークンを取得して保存する
 * @param oAuth2Client OAuth2クライアント
 */
function getAndSaveToken(oAuth2Client) {
    return __awaiter(this, void 0, void 0, function* () {
        const authUrl = oAuth2Client.generateAuthUrl({
            access_type: 'offline',
            scope: SCOPES,
        });
        console.log('以下のURLにアクセスして認証コードを取得してください:');
        console.log(authUrl);
        const rl = readline_1.default.createInterface({
            input: process.stdin,
            output: process.stdout,
        });
        return new Promise((resolve, reject) => {
            rl.question('認証コードを入力してください: ', (code) => __awaiter(this, void 0, void 0, function* () {
                rl.close();
                try {
                    const { tokens } = yield oAuth2Client.getToken(code);
                    oAuth2Client.setCredentials(tokens);
                    // トークンをファイルに保存
                    fs_1.default.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
                    console.log(`トークンが ${TOKEN_PATH} に保存されました`);
                    resolve(tokens);
                }
                catch (error) {
                    console.error('トークンの取得に失敗しました:', error);
                    reject(error);
                }
            }));
        });
    });
}
// 認証処理とサーバー起動
function authorize() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            // credentials ファイルの読み込み
            if (!fs_1.default.existsSync(CREDENTIALS_PATH)) {
                console.error(`認証情報ファイルが見つかりません: ${CREDENTIALS_PATH}`);
                process.exit(1);
            }
            const credentialsContent = fs_1.default.readFileSync(CREDENTIALS_PATH, 'utf8');
            const credentials = JSON.parse(credentialsContent);
            const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
            const oAuth2Client = new googleapis_1.google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
            // トークンファイルの確認
            let token;
            if (!fs_1.default.existsSync(TOKEN_PATH)) {
                console.log(`トークンファイルが見つかりません: ${TOKEN_PATH}`);
                console.log('認証フローを開始します...');
                token = yield getAndSaveToken(oAuth2Client);
            }
            else {
                const tokenContent = fs_1.default.readFileSync(TOKEN_PATH, 'utf8');
                token = JSON.parse(tokenContent);
                oAuth2Client.setCredentials(token);
            }
            // サーバーの起動
            startServer(credentials, token);
        }
        catch (error) {
            console.error('認証処理中にエラーが発生しました:', error);
            process.exit(1);
        }
    });
}
// サーバー起動関数
function startServer(credentials, token) {
    // GoogleSheetServiceの初期化
    const googleSheetService = new GoogleSheetService_1.GoogleSheetService(credentials, token);
    // TestSheetBuilderの初期化
    const testSheetBuilder = new TestSheetBuilder_1.TestSheetBuilder(googleSheetService);
    // Express アプリケーションの設定
    const app = (0, express_1.default)();
    const port = process.env.PORT || 3005;
    app.use((0, cors_1.default)());
    app.use(express_1.default.json());
    // ヘルスチェックエンドポイント
    app.get('/health', (req, res) => {
        res.status(200).json({ status: 'ok' });
    });
    // テストシートを生成するエンドポイント
    app.post('/generate-test-sheet', (req, res) => __awaiter(this, void 0, void 0, function* () {
        try {
            const { templateId, title, prompt, useOrthogonalArray } = req.body;
            if (!templateId || !title || !prompt) {
                return res.status(400).json({ error: 'templateId, title, promptは必須パラメータです' });
            }
            // プロンプトから因子と水準を生成
            const factors = testSheetBuilder.generateFactorsFromPrompt(prompt);
            // テストケースの生成（直交表を使うかどうかで分岐）
            const testCases = useOrthogonalArray
                ? testSheetBuilder.generateOrthogonalArrayTestCases(factors)
                : testSheetBuilder.generateAllCombinationsTestCases(factors);
            // テストシートの作成
            const newSheetId = yield testSheetBuilder.createTestSheet(templateId, title, factors, testCases);
            res.status(200).json({
                sheetId: newSheetId,
                sheetUrl: `https://docs.google.com/spreadsheets/d/${newSheetId}`,
                factorCount: factors.length,
                testCaseCount: testCases.length
            });
        }
        catch (error) {
            console.error('テストシート生成中にエラーが発生しました:', error);
            res.status(500).json({ error: 'テストシートの生成に失敗しました' });
        }
    }));
    // スプレッドシートの情報を取得するエンドポイント
    app.get('/spreadsheet/:id', (req, res) => __awaiter(this, void 0, void 0, function* () {
        try {
            const { id } = req.params;
            const { range } = req.query;
            if (!id) {
                return res.status(400).json({ error: 'スプレッドシートIDは必須パラメータです' });
            }
            const sheetRange = typeof range === 'string' ? range : 'Sheet1!A1:Z100';
            const values = yield googleSheetService.getSheetValues(id, sheetRange);
            res.status(200).json({ values });
        }
        catch (error) {
            console.error('スプレッドシートの取得中にエラーが発生しました:', error);
            res.status(500).json({ error: 'スプレッドシートの取得に失敗しました' });
        }
    }));
    // スプレッドシートを更新するエンドポイント
    app.post('/spreadsheet/:id/update', (req, res) => __awaiter(this, void 0, void 0, function* () {
        try {
            const { id } = req.params;
            const { range, values } = req.body;
            if (!id || !range || !values) {
                return res.status(400).json({ error: 'id, range, valuesは必須パラメータです' });
            }
            yield googleSheetService.updateSheetValues(id, range, values);
            res.status(200).json({ success: true });
        }
        catch (error) {
            console.error('スプレッドシートの更新中にエラーが発生しました:', error);
            res.status(500).json({ error: 'スプレッドシートの更新に失敗しました' });
        }
    }));
    // HTTPサーバー作成
    const server = http_1.default.createServer(app);
    // MCP関連のツール定義
    const mcpTools = {
        'generate-test': {
            description: 'テストシートを生成します',
            parameters: {
                templateId: {
                    type: 'string',
                    description: 'テンプレートとなるスプレッドシートのID'
                },
                title: {
                    type: 'string',
                    description: '生成するスプレッドシートのタイトル'
                },
                prompt: {
                    type: 'string',
                    description: 'テスト要件を記述したプロンプト'
                },
                useOrthogonalArray: {
                    type: 'boolean',
                    description: '直交表を使用するかどうか（省略時はfalse）'
                }
            },
            handler: (params) => __awaiter(this, void 0, void 0, function* () {
                try {
                    const { templateId, title, prompt, useOrthogonalArray } = params;
                    if (!templateId || !title || !prompt) {
                        return { error: 'templateId, title, promptは必須パラメータです' };
                    }
                    // プロンプトから因子と水準を生成
                    const factors = testSheetBuilder.generateFactorsFromPrompt(prompt);
                    // テストケースの生成（直交表を使うかどうかで分岐）
                    const testCases = useOrthogonalArray
                        ? testSheetBuilder.generateOrthogonalArrayTestCases(factors)
                        : testSheetBuilder.generateAllCombinationsTestCases(factors);
                    // テストシートの作成
                    const newSheetId = yield testSheetBuilder.createTestSheet(templateId, title, factors, testCases);
                    return {
                        sheetId: newSheetId,
                        sheetUrl: `https://docs.google.com/spreadsheets/d/${newSheetId}`,
                        factorCount: factors.length,
                        testCaseCount: testCases.length
                    };
                }
                catch (error) {
                    console.error('テストシート生成中にエラーが発生しました:', error);
                    return { error: 'テストシートの生成に失敗しました' };
                }
            })
        },
        'get-spreadsheet': {
            description: 'スプレッドシートの情報を取得します',
            parameters: {
                id: {
                    type: 'string',
                    description: 'スプレッドシートのID'
                },
                range: {
                    type: 'string',
                    description: '取得する範囲（例: Sheet1!A1:Z100）'
                }
            },
            handler: (params) => __awaiter(this, void 0, void 0, function* () {
                try {
                    const { id, range } = params;
                    if (!id) {
                        return { error: 'スプレッドシートIDは必須パラメータです' };
                    }
                    const sheetRange = range || 'Sheet1!A1:Z100';
                    const values = yield googleSheetService.getSheetValues(id, sheetRange);
                    return { values };
                }
                catch (error) {
                    console.error('スプレッドシートの取得中にエラーが発生しました:', error);
                    return { error: 'スプレッドシートの取得に失敗しました' };
                }
            })
        }
    };
    // サーバー起動
    server.listen(port, () => {
        console.log(`MCP Test Sheet Builder サーバーが起動しました: http://localhost:${port}`);
        console.log(`認証情報ファイル: ${CREDENTIALS_PATH}`);
        console.log(`トークンファイル: ${TOKEN_PATH}`);
    });
    // MCP処理に対応するための標準入力のリスナー
    process.stdin.on('data', (data) => {
        try {
            const message = JSON.parse(data.toString());
            // 初期化メッセージ
            if (message.type === 'initialize') {
                const response = {
                    type: 'initialize_response',
                    tools: Object.entries(mcpTools).map(([name, tool]) => ({
                        name,
                        description: tool.description,
                        parameters: tool.parameters
                    }))
                };
                process.stdout.write(JSON.stringify(response) + '\n');
            }
            // ツール呼び出しメッセージ
            else if (message.type === 'call_tool') {
                const toolName = message.name;
                const tool = mcpTools[toolName];
                if (!tool) {
                    const response = {
                        type: 'call_tool_response',
                        id: message.id,
                        error: `ツール "${toolName}" が見つかりません`
                    };
                    process.stdout.write(JSON.stringify(response) + '\n');
                }
                else {
                    (() => __awaiter(this, void 0, void 0, function* () {
                        try {
                            const result = yield tool.handler(message.parameters);
                            const response = {
                                type: 'call_tool_response',
                                id: message.id,
                                result
                            };
                            process.stdout.write(JSON.stringify(response) + '\n');
                        }
                        catch (error) {
                            const response = {
                                type: 'call_tool_response',
                                id: message.id,
                                error: `ツールの実行中にエラーが発生しました: ${error}`
                            };
                            process.stdout.write(JSON.stringify(response) + '\n');
                        }
                    }))();
                }
            }
        }
        catch (error) {
            console.error('MCP処理中にエラーが発生しました:', error);
        }
    });
    // プロセス終了時のハンドリング
    process.on('SIGTERM', () => {
        console.log('SIGTERM を受信しました。サーバーをシャットダウンします...');
        server.close();
        process.exit(0);
    });
    process.on('SIGINT', () => {
        console.log('SIGINT を受信しました。サーバーをシャットダウンします...');
        server.close();
        process.exit(0);
    });
}
// 認証処理を開始してサーバーを起動
authorize();
