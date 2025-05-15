import { google } from 'googleapis';
export class GoogleSheetService {
    constructor(credentials, token) {
        const { client_secret, client_id, redirect_uris } = credentials.installed;
        const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
        oAuth2Client.setCredentials(token);
        this.sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
        this.drive = google.drive({ version: 'v3', auth: oAuth2Client });
    }
    /**
     * スプレッドシートの内容を取得する
     * @param spreadsheetId スプレッドシートID
     * @param range 範囲
     */
    async getSheetValues(spreadsheetId, range) {
        try {
            const response = await this.sheets.spreadsheets.values.get({
                spreadsheetId,
                range,
            });
            return response.data.values || [];
        }
        catch (error) {
            // console.error('スプレッドシートの値の取得に失敗しました:', error);
            throw error;
        }
    }
    /**
     * スプレッドシートの内容を取得する（別名）
     * @param spreadsheetId スプレッドシートID
     * @param range 範囲
     */
    async getValues(spreadsheetId, range) {
        try {
            const response = await this.sheets.spreadsheets.values.get({
                spreadsheetId,
                range,
            });
            return {
                values: response.data.values || []
            };
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * スプレッドシートの内容を更新する
     * @param spreadsheetId スプレッドシートID
     * @param range 範囲
     * @param values 値
     */
    async updateSheetValues(spreadsheetId, range, values) {
        try {
            await this.sheets.spreadsheets.values.update({
                spreadsheetId,
                range,
                valueInputOption: 'USER_ENTERED',
                requestBody: {
                    values,
                },
            });
        }
        catch (error) {
            // console.error('スプレッドシートの値の更新に失敗しました:', error);
            throw error;
        }
    }
    /**
     * スプレッドシートの内容を更新する（別名）
     * @param spreadsheetId スプレッドシートID
     * @param range 範囲
     * @param values 値
     */
    async updateValues(spreadsheetId, range, values) {
        return this.updateSheetValues(spreadsheetId, range, values);
    }
    /**
     * テンプレートからスプレッドシートをコピーする
     * @param templateId テンプレートのスプレッドシートID
     * @param title 新しいスプレッドシートのタイトル
     */
    async copySpreadsheet(templateId, title) {
        try {
            const response = await this.drive.files.copy({
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
            // console.error('スプレッドシートのコピーに失敗しました:', error);
            throw error;
        }
    }
    /**
     * スプレッドシートの情報を取得する
     * @param spreadsheetId スプレッドシートID
     */
    async getSpreadsheetInfo(spreadsheetId) {
        try {
            const response = await this.sheets.spreadsheets.get({
                spreadsheetId
            });
            return response.data;
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * スプレッドシートのシート一覧を取得する
     * @param spreadsheetId スプレッドシートID
     */
    async listSheets(spreadsheetId) {
        try {
            const response = await this.sheets.spreadsheets.get({
                spreadsheetId,
                fields: 'sheets.properties'
            });
            return {
                sheets: response.data.sheets || []
            };
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * シートが存在するか確認する
     * @param spreadsheetId スプレッドシートID
     * @param sheetName シート名
     */
    async checkSheetExists(spreadsheetId, sheetName) {
        try {
            const spreadsheet = await this.getSpreadsheetInfo(spreadsheetId);
            if (spreadsheet.sheets) {
                return spreadsheet.sheets.some(sheet => sheet.properties?.title === sheetName);
            }
            return false;
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * シートを追加する
     * @param spreadsheetId スプレッドシートID
     * @param sheetName シート名
     */
    async addSheet(spreadsheetId, sheetName) {
        try {
            const response = await this.sheets.spreadsheets.batchUpdate({
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
            return response.data;
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * シートの特定範囲をクリアする
     * @param spreadsheetId スプレッドシートID
     * @param range クリアする範囲
     */
    async clearSheet(spreadsheetId, range) {
        try {
            await this.sheets.spreadsheets.values.clear({
                spreadsheetId,
                range
            });
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * シート内に表を複製する
     * @param spreadsheetId スプレッドシートID
     * @param sourceRange ソースとなる表の範囲
     * @param destinationRange 複製先の範囲
     */
    async duplicateTable(spreadsheetId, sourceRange, destinationRange) {
        try {
            // ソース範囲のデータを取得
            const sourceData = await this.getSheetValues(spreadsheetId, sourceRange);
            if (!sourceData || sourceData.length === 0) {
                throw new Error('複製元の表にデータがありません');
            }
            // 宛先に書き込み
            await this.updateSheetValues(spreadsheetId, destinationRange, sourceData);
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * 指定されたフォルダにファイルを移動する
     * @param fileId ファイルID
     * @param folderId フォルダID
     */
    async moveFileToFolder(fileId, folderId) {
        try {
            // 現在の親フォルダを取得
            const file = await this.drive.files.get({
                fileId: fileId,
                fields: 'parents'
            });
            // 親フォルダのリスト
            const previousParents = file.data.parents?.join(',') || '';
            // ファイルを新しいフォルダに移動
            await this.drive.files.update({
                fileId: fileId,
                addParents: folderId,
                removeParents: previousParents,
                fields: 'id, parents'
            });
        }
        catch (error) {
            throw error;
        }
    }
    /**
     * フォルダIDの有効性を確認する
     * @param folderId フォルダID
     */
    async validateFolderId(folderId) {
        try {
            const response = await this.drive.files.get({
                fileId: folderId,
                fields: 'mimeType'
            });
            // Google DriveのフォルダのmimeTypeを確認
            return response.data.mimeType === 'application/vnd.google-apps.folder';
        }
        catch (error) {
            return false;
        }
    }
}
