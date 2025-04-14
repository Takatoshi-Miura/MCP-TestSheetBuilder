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
Object.defineProperty(exports, "__esModule", { value: true });
exports.GoogleSheetService = void 0;
const googleapis_1 = require("googleapis");
class GoogleSheetService {
    constructor(credentials, token) {
        const { client_secret, client_id, redirect_uris } = credentials.installed;
        const oAuth2Client = new googleapis_1.google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
        oAuth2Client.setCredentials(token);
        this.sheets = googleapis_1.google.sheets({ version: 'v4', auth: oAuth2Client });
        this.drive = googleapis_1.google.drive({ version: 'v3', auth: oAuth2Client });
    }
    /**
     * スプレッドシートの内容を取得する
     * @param spreadsheetId スプレッドシートID
     * @param range 範囲
     */
    getSheetValues(spreadsheetId, range) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const response = yield this.sheets.spreadsheets.values.get({
                    spreadsheetId,
                    range,
                });
                return response.data.values || [];
            }
            catch (error) {
                console.error('スプレッドシートの値の取得に失敗しました:', error);
                throw error;
            }
        });
    }
    /**
     * スプレッドシートの内容を更新する
     * @param spreadsheetId スプレッドシートID
     * @param range 範囲
     * @param values 値
     */
    updateSheetValues(spreadsheetId, range, values) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                yield this.sheets.spreadsheets.values.update({
                    spreadsheetId,
                    range,
                    valueInputOption: 'USER_ENTERED',
                    requestBody: {
                        values,
                    },
                });
            }
            catch (error) {
                console.error('スプレッドシートの値の更新に失敗しました:', error);
                throw error;
            }
        });
    }
    /**
     * テンプレートからスプレッドシートをコピーする
     * @param templateId テンプレートのスプレッドシートID
     * @param title 新しいスプレッドシートのタイトル
     */
    copySpreadsheet(templateId, title) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const response = yield this.drive.files.copy({
                    fileId: templateId,
                    requestBody: {
                        name: title,
                    },
                });
                if (!response.data.id) {
                    throw new Error('コピーしたファイルのIDが取得できませんでした');
                }
                return response.data.id;
            }
            catch (error) {
                console.error('スプレッドシートのコピーに失敗しました:', error);
                throw error;
            }
        });
    }
}
exports.GoogleSheetService = GoogleSheetService;
