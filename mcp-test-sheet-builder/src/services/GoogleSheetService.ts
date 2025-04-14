import { google, sheets_v4, drive_v3 } from 'googleapis';

export class GoogleSheetService {
  private sheets: sheets_v4.Sheets;
  private drive: drive_v3.Drive;

  constructor(credentials: any, token: any) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
      client_id,
      client_secret,
      redirect_uris[0]
    );
    
    oAuth2Client.setCredentials(token);
    
    this.sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
    this.drive = google.drive({ version: 'v3', auth: oAuth2Client });
  }

  /**
   * スプレッドシートの内容を取得する
   * @param spreadsheetId スプレッドシートID
   * @param range 範囲
   */
  async getSheetValues(spreadsheetId: string, range: string): Promise<any[][]> {
    try {
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
      });
      
      return response.data.values || [];
    } catch (error) {
      console.error('スプレッドシートの値の取得に失敗しました:', error);
      throw error;
    }
  }

  /**
   * スプレッドシートの内容を更新する
   * @param spreadsheetId スプレッドシートID
   * @param range 範囲
   * @param values 値
   */
  async updateSheetValues(spreadsheetId: string, range: string, values: any[][]): Promise<void> {
    try {
      await this.sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values,
        },
      });
    } catch (error) {
      console.error('スプレッドシートの値の更新に失敗しました:', error);
      throw error;
    }
  }

  /**
   * テンプレートからスプレッドシートをコピーする
   * @param templateId テンプレートのスプレッドシートID
   * @param title 新しいスプレッドシートのタイトル
   */
  async copySpreadsheet(templateId: string, title: string): Promise<string> {
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
    } catch (error) {
      console.error('スプレッドシートのコピーに失敗しました:', error);
      throw error;
    }
  }
} 