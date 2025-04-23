import { GoogleSheetService } from './GoogleSheetService.js';

interface Factor {
  name: string;
  levels: string[];
}

interface TestCase {
  [key: string]: string;
}

// 機能と因子・水準のインターフェース
interface FeatureFactors {
  featureName: string;  // 機能名
  factors: Factor[];    // 因子と水準のリスト
}

export class TestSheetBuilder {
  private googleSheetService: GoogleSheetService;

  constructor(googleSheetService: GoogleSheetService) {
    this.googleSheetService = googleSheetService;
  }

  /**
   * プロンプトから因子と水準を生成する
   * @param prompt プロンプト
   */
  generateFactorsFromPrompt(prompt: string): Factor[] {
    // この実装はとても単純化されています。実際の実装ではNLPやAIを使用することが考えられます。
    // 例としての簡単な実装:
    const factors: Factor[] = [];
    
    // プロンプトから基本的な因子を抽出する簡易処理の例
    if (prompt.includes('ブラウザ')) {
      factors.push({
        name: 'ブラウザ',
        levels: ['Chrome', 'Firefox', 'Safari', 'Edge']
      });
    }
    
    if (prompt.includes('OS')) {
      factors.push({
        name: 'OS',
        levels: ['Windows', 'macOS', 'Linux']
      });
    }
    
    if (prompt.includes('画面サイズ') || prompt.includes('解像度')) {
      factors.push({
        name: '画面サイズ',
        levels: ['スマホ', 'タブレット', 'デスクトップ']
      });
    }
    
    if (prompt.includes('ネットワーク')) {
      factors.push({
        name: 'ネットワーク状態',
        levels: ['高速', '遅延あり', 'オフライン']
      });
    }
    
    // もし因子が見つからなかった場合、デフォルトの因子を追加
    if (factors.length === 0) {
      factors.push({
        name: 'テスト環境',
        levels: ['開発環境', 'テスト環境', '本番環境']
      });
      
      factors.push({
        name: 'ユーザー権限',
        levels: ['一般ユーザー', '管理者']
      });
    }
    
    return factors;
  }

  /**
   * プロンプトから機能と因子・水準のリストを生成する
   * @param prompt プロンプト
   */
  generateFeatureFactorsFromPrompt(prompt: string): FeatureFactors[] {
    const featureFactors: FeatureFactors[] = [];
    
    // プロンプトから機能を抽出する処理
    // 以下は例として、実際はもっと高度なNLPやAIを使用することが考えられます
    
    // 画面に関する因子・水準
    if (prompt.includes('画面') || prompt.includes('ページ') || prompt.includes('UI')) {
      featureFactors.push({
        featureName: '画面テスト',
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
        featureName: 'ワークフローテスト',
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
        featureName: 'データ操作テスト',
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
        featureName: '検索機能テスト',
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
        featureName: '帳票出力テスト',
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
        featureName: '基本機能テスト',
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

  /**
   * 全因子組み合わせテストケースを生成する
   * @param factors 因子と水準
   */
  generateAllCombinationsTestCases(factors: Factor[]): TestCase[] {
    if (factors.length === 0) {
      return [];
    }
    
    // 再帰的に組み合わせを生成
    const generateCombinations = (index: number, currentCombination: TestCase): TestCase[] => {
      if (index === factors.length) {
        return [currentCombination];
      }
      
      const factor = factors[index];
      const result: TestCase[] = [];
      
      for (const level of factor.levels) {
        const newCombination = { ...currentCombination };
        newCombination[factor.name] = level;
        result.push(...generateCombinations(index + 1, newCombination));
      }
      
      return result;
    };
    
    return generateCombinations(0, {});
  }

  /**
   * 直交表によるテストケースを生成する（簡易版）
   * @param factors 因子と水準
   */
  generateOrthogonalArrayTestCases(factors: Factor[]): TestCase[] {
    // 簡易版の実装。本来は複雑な直交表アルゴリズムが必要
    // ここでは因子数を減らすだけの簡単な実装
    if (factors.length <= 2) {
      return this.generateAllCombinationsTestCases(factors);
    }
    
    // 代表的な水準を選択して組み合わせる簡易実装
    const testCases: TestCase[] = [];
    const mainFactors = factors.slice(0, 2); // 最初の2つの因子は全組み合わせ
    const mainCombinations = this.generateAllCombinationsTestCases(mainFactors);
    
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

  /**
   * テンプレートからテストシートを作成し、因子水準とテストケースを書き込む
   * @param templateId テンプレートスプレッドシートID
   * @param title 新しいスプレッドシートのタイトル
   * @param factors 因子と水準
   * @param testCases テストケース
   */
  async createTestSheet(templateId: string, title: string, factors: Factor[], testCases: TestCase[]): Promise<string> {
    // テンプレートをコピー
    const newSheetId = await this.googleSheetService.copySpreadsheet(templateId, title);
    
    // 因子と水準をシートに書き込む
    const factorValues = [
      ['因子', '水準'],
      ...factors.map(factor => [factor.name, factor.levels.join(', ')])
    ];
    
    await this.googleSheetService.updateSheetValues(newSheetId, '因子水準!A1:B' + (factors.length + 1), factorValues);
    
    // テストケースをシートに書き込む
    if (testCases.length > 0) {
      // ヘッダー行の作成
      const headers = ['No.', ...Object.keys(testCases[0]), '結果', '備考'];
      
      // データ行の作成
      const rows = testCases.map((testCase, index) => {
        return [
          (index + 1).toString(),
          ...Object.values(testCase),
          '', // 結果列
          ''  // 備考列
        ];
      });
      
      const testCaseValues = [headers, ...rows];
      await this.googleSheetService.updateSheetValues(newSheetId, 'テストケース!A1:' + this.columnIndexToLetter(headers.length) + (testCases.length + 1), testCaseValues);
    }
    
    return newSheetId;
  }

  /**
   * 複数の機能ごとの因子・水準表を作成する
   * @param spreadsheetId スプレッドシートID
   * @param featureFactors 機能ごとの因子と水準のリスト
   */
  async createFeatureFactorTables(spreadsheetId: string, featureFactors: FeatureFactors[]): Promise<void> {
    // 因子・水準シートが存在することを確認
    const sheetExists = await this.googleSheetService.checkSheetExists(spreadsheetId, '因子・水準');
    if (!sheetExists) {
      await this.googleSheetService.addSheet(spreadsheetId, '因子・水準');
      
      // 初期ヘッダーを設定
      await this.googleSheetService.updateSheetValues(spreadsheetId, '因子・水準!A1:B1', [['因子', '水準']]);
    }
    
    // 現在のシートの内容を確認
    const currentValues = await this.googleSheetService.getSheetValues(spreadsheetId, '因子・水準!A:B');
    
    // 最終行の位置を取得
    let lastRow = currentValues.length + 2; // 余白を入れるため+2
    
    // 各機能ごとに因子・水準表を書き込む
    for (const featureFactor of featureFactors) {
      // 機能名を書き込む
      await this.googleSheetService.updateSheetValues(
        spreadsheetId, 
        `因子・水準!A${lastRow}:B${lastRow}`, 
        [[`【${featureFactor.featureName}】`, '']]
      );
      lastRow++;
      
      // 因子・水準テーブルのヘッダー
      await this.googleSheetService.updateSheetValues(
        spreadsheetId, 
        `因子・水準!A${lastRow}:B${lastRow}`, 
        [['因子', '水準']]
      );
      lastRow++;
      
      // 各因子と水準を書き込む
      for (const factor of featureFactor.factors) {
        // 1つのセルに1つのワードのみ記載するため、水準を別々のセルに
        const levelRows = factor.levels.map(level => [factor.name, level]);
        
        // 最初の行だけ因子名を入れる
        for (let i = 1; i < levelRows.length; i++) {
          levelRows[i][0] = ''; // 因子名を空にする
        }
        
        // シートに書き込む
        await this.googleSheetService.updateSheetValues(
          spreadsheetId, 
          `因子・水準!A${lastRow}:B${lastRow + levelRows.length - 1}`, 
          levelRows
        );
        
        lastRow += levelRows.length;
      }
      
      // 機能間の区切りのために空行を追加
      lastRow += 2;
    }
  }

  /**
   * テンプレートからテストシートを作成し、複数の機能ごとの因子・水準表を作成する
   * @param templateId テンプレートスプレッドシートID
   * @param title 新しいスプレッドシートのタイトル
   * @param prompt テスト要件を記述したプロンプト
   * @param folderId (オプション) 保存先フォルダID
   */
  async createFeatureFactorSheet(templateId: string, title: string, prompt: string, folderId?: string): Promise<string> {
    // テンプレートをコピー
    const newSheetId = await this.googleSheetService.copySpreadsheet(templateId, title);
    
    // プロンプトから機能と因子・水準を生成
    const featureFactors = this.generateFeatureFactorsFromPrompt(prompt);
    
    // 機能ごとの因子・水準表を作成
    await this.createFeatureFactorTables(newSheetId, featureFactors);
    
    // 指定されたフォルダに移動（指定がある場合）
    if (folderId) {
      const isValidFolder = await this.googleSheetService.validateFolderId(folderId);
      if (isValidFolder) {
        await this.googleSheetService.moveFileToFolder(newSheetId, folderId);
      }
    }
    
    return newSheetId;
  }

  /**
   * 列番号をA1形式の列文字に変換する
   * @param index 列番号（1始まり）
   */
  private columnIndexToLetter(index: number): string {
    let temp, letter = '';
    while (index > 0) {
      temp = (index - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      index = (index - temp - 1) / 26;
    }
    return letter;
  }
} 