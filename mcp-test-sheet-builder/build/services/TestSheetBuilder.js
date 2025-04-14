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
exports.TestSheetBuilder = void 0;
class TestSheetBuilder {
    constructor(googleSheetService) {
        this.googleSheetService = googleSheetService;
    }
    /**
     * プロンプトから因子と水準を生成する
     * @param prompt プロンプト
     */
    generateFactorsFromPrompt(prompt) {
        // この実装はとても単純化されています。実際の実装ではNLPやAIを使用することが考えられます。
        // 例としての簡単な実装:
        const factors = [];
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
     * 全因子組み合わせテストケースを生成する
     * @param factors 因子と水準
     */
    generateAllCombinationsTestCases(factors) {
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
                const newCombination = Object.assign({}, currentCombination);
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
    generateOrthogonalArrayTestCases(factors) {
        // 簡易版の実装。本来は複雑な直交表アルゴリズムが必要
        // ここでは因子数を減らすだけの簡単な実装
        if (factors.length <= 2) {
            return this.generateAllCombinationsTestCases(factors);
        }
        // 代表的な水準を選択して組み合わせる簡易実装
        const testCases = [];
        const mainFactors = factors.slice(0, 2); // 最初の2つの因子は全組み合わせ
        const mainCombinations = this.generateAllCombinationsTestCases(mainFactors);
        const otherFactors = factors.slice(2);
        for (const combination of mainCombinations) {
            const testCase = Object.assign({}, combination);
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
    createTestSheet(templateId, title, factors, testCases) {
        return __awaiter(this, void 0, void 0, function* () {
            // テンプレートをコピー
            const newSheetId = yield this.googleSheetService.copySpreadsheet(templateId, title);
            // 因子と水準をシートに書き込む
            const factorValues = [
                ['因子', '水準'],
                ...factors.map(factor => [factor.name, factor.levels.join(', ')])
            ];
            yield this.googleSheetService.updateSheetValues(newSheetId, '因子水準!A1:B' + (factors.length + 1), factorValues);
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
                        '' // 備考列
                    ];
                });
                const testCaseValues = [headers, ...rows];
                yield this.googleSheetService.updateSheetValues(newSheetId, 'テストケース!A1:' + this.columnIndexToLetter(headers.length) + (testCases.length + 1), testCaseValues);
            }
            return newSheetId;
        });
    }
    /**
     * 列番号をA1形式の列文字に変換する
     * @param index 列番号（1始まり）
     */
    columnIndexToLetter(index) {
        let temp, letter = '';
        while (index > 0) {
            temp = (index - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            index = (index - temp - 1) / 26;
        }
        return letter;
    }
}
exports.TestSheetBuilder = TestSheetBuilder;
