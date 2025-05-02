export class TestSheetBuilder {
    constructor(googleSheetService) {
        this.googleSheetService = googleSheetService;
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
    /**
     * 因子水準シートからテスト項目を生成する
     * @param spreadsheetId スプレッドシートID
     * @returns 生成結果
     */
    async generateTestItemsFromFactorsAndLevels(spreadsheetId) {
        try {
            // 因子水準シートの存在確認
            const sheetList = await this.googleSheetService.listSheets(spreadsheetId);
            const factorLevelSheetInfo = sheetList.sheets.find((sheet) => sheet.properties?.title === '因子・水準');
            if (!factorLevelSheetInfo) {
                return { success: false, message: '因子・水準シートが見つかりません。' };
            }
            // 因子水準シートの内容取得
            const factorLevelData = await this.googleSheetService.getValues(spreadsheetId, '因子・水準!A1:Z100');
            // 因子と水準を解析
            const factors = this.parseFactorsAndLevels(factorLevelData.values || []);
            if (factors.length === 0) {
                return { success: false, message: '因子・水準の内容が見つかりません。' };
            }
            // テスト項目を生成
            const testItems = this.generateTestItems(factors);
            // テスト項目シートの有無を確認
            const testItemSheetInfo = sheetList.sheets.find((sheet) => sheet.properties?.title === 'テスト項目');
            let testItemSheetId;
            // テスト項目シートが存在しない場合は作成
            if (!testItemSheetInfo) {
                // テスト項目シートを作成
                const addSheetResult = await this.googleSheetService.addSheet(spreadsheetId, 'テスト項目');
                testItemSheetId = addSheetResult.replies?.[0].addSheet?.properties?.sheetId || 0;
                // ヘッダー行を設定
                await this.googleSheetService.updateValues(spreadsheetId, 'テスト項目!A1:I3', [
                    ['№', 'テスト項目', '', '', '入力値・前提条件など', '', '操作など', '想定される結果', 'テスター１：\nアプリVer：\nOS：\n機種：'],
                    ['', '', '', '', '', '', '', '', '実施日'],
                    ['', '', '', '', '', '入力値（コピペ用）']
                ]);
            }
            else {
                testItemSheetId = testItemSheetInfo.properties?.sheetId || 0;
                // 既存のテスト項目シートを一旦クリア（ヘッダー3行を残す）
                await this.googleSheetService.clearSheet(spreadsheetId, 'テスト項目!A4:Z1000');
            }
            // テスト項目を書き込み
            if (testItems.length > 0) {
                await this.googleSheetService.updateValues(spreadsheetId, `テスト項目!A4:H${3 + testItems.length}`, testItems);
            }
            // セルの書式設定（任意）
            // ...省略...
            return { success: true, message: 'テスト項目の生成が完了しました。' };
        }
        catch (error) {
            console.error('テスト項目生成エラー:', error);
            return { success: false, message: `エラーが発生しました: ${error}` };
        }
    }
    /**
     * 因子水準データを解析する
     * @param data シートから読み取ったデータ
     * @returns 解析された因子と水準の配列
     */
    parseFactorsAndLevels(data) {
        const factors = [];
        // 「因子」「水準」と書かれた行を検索
        let factorRowIndex = -1;
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length < 2)
                continue;
            const cell1 = row[0]?.toString().trim();
            const cell2 = row[1]?.toString().trim();
            if (cell1 === '因子' && cell2 === '水準') {
                factorRowIndex = i;
                break;
            }
        }
        if (factorRowIndex === -1) {
            return factors; // 因子・水準の行が見つからない
        }
        // 因子の行を特定（因子・水準の次の行）
        const factorRow = data[factorRowIndex + 1];
        if (!factorRow || factorRow.length < 2) {
            return factors;
        }
        // コンポーネント因子を処理
        const componentFactor = {
            name: factorRow[0]?.toString().trim() || 'コンポーネント',
            levels: []
        };
        // コンポーネントの水準を収集
        for (let i = 1; i < factorRow.length; i++) {
            const level = factorRow[i]?.toString().trim();
            if (level && level !== '') {
                componentFactor.levels.push(level);
            }
        }
        if (componentFactor.levels.length > 0) {
            factors.push(componentFactor);
        }
        // 期待値因子を処理
        for (let i = factorRowIndex + 2; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length < 2)
                continue;
            const headerCell = row[0]?.toString().trim();
            // 通常の期待値
            if (headerCell === '期待値') {
                const expectFactor = {
                    name: headerCell,
                    levels: []
                };
                // 期待値の水準を収集
                for (let j = 1; j < row.length; j++) {
                    const level = row[j]?.toString().trim();
                    if (level && level !== '') {
                        expectFactor.levels.push(level);
                    }
                }
                if (expectFactor.levels.length > 0) {
                    factors.push(expectFactor);
                }
            }
            // 特殊な期待値（例: 期待値: テキスト複数行のみ）
            else if (headerCell.startsWith('期待値:')) {
                const specialFactor = {
                    name: headerCell,
                    levels: []
                };
                // この特殊期待値の水準を収集
                for (let j = 1; j < row.length; j++) {
                    const level = row[j]?.toString().trim();
                    if (level && level !== '') {
                        specialFactor.levels.push(level);
                    }
                }
                // 次の行も同じ特殊期待値の水準かもしれないのでチェック
                for (let k = i + 1; k < data.length; k++) {
                    const nextRow = data[k];
                    if (!nextRow || nextRow.length < 2)
                        break;
                    const nextHeaderCell = nextRow[0]?.toString().trim();
                    if (nextHeaderCell === '' || nextHeaderCell === undefined) {
                        // 同じ特殊期待値の追加水準
                        for (let j = 1; j < nextRow.length; j++) {
                            const level = nextRow[j]?.toString().trim();
                            if (level && level !== '') {
                                specialFactor.levels.push(level);
                            }
                        }
                    }
                    else {
                        // 別の項目なので処理終了
                        break;
                    }
                }
                if (specialFactor.levels.length > 0) {
                    factors.push(specialFactor);
                }
            }
        }
        return factors;
    }
    /**
     * テスト項目を生成する
     * @param factors 因子と水準の配列
     * @returns 生成されたテスト項目の配列
     */
    generateTestItems(factors) {
        // ヘッダー行はシートに既に存在するため、編集しない
        // テスト項目は4行目から開始
        const testItems = [];
        // 因子と水準の組み合わせ生成
        let testCases = this.generateCombinations(factors);
        // テスト項目作成
        for (let i = 0; i < testCases.length; i++) {
            const testCase = testCases[i];
            const testNumber = i + 1;
            // テスト項目行の作成
            // A列: テスト番号
            // B列: テストID (B~D列で3つまでの因子水準を表示)
            // E列: 前提準備となる操作
            // G列: テスト時の操作
            // H列: G列で行った操作によって起こる期待する結果
            // テスト項目の基本情報
            let row = [
                testNumber.toString(), // A列: No.
            ];
            // B列からD列には因子水準を記載 (最大3つまで)
            // 水準値を取得
            let factorLevels = [];
            for (let j = 0; j < testCase.values.length; j++) {
                factorLevels.push(testCase.values[j]);
            }
            // 因子名と水準値を組み合わせてテスト項目名を作成
            let itemName = '';
            for (let j = 0; j < Math.min(factorLevels.length, 1); j++) {
                itemName += factorLevels[j];
            }
            row.push(itemName); // B列: 1つ目の因子水準を表示
            // C列とD列に残りの因子水準を表示（あれば）
            if (factorLevels.length > 1) {
                row.push(factorLevels[1]); // C列: 2つ目の因子水準
            }
            else {
                row.push(''); // C列: 空
            }
            if (factorLevels.length > 2) {
                row.push(factorLevels[2]); // D列: 3つ目の因子水準
            }
            else {
                row.push(''); // D列: 空
            }
            // E列: 前提準備となる操作
            const preconditions = this.generatePreconditions(testCase);
            row.push(preconditions);
            // F列: 空欄（または入力値コピペ用）
            row.push('');
            // G列: テスト時の操作
            const testProcedure = this.generateTestProcedure(testCase);
            row.push(testProcedure);
            // H列: 期待する結果
            const expectedResult = this.generateExpectedResult(testCase);
            row.push(expectedResult);
            testItems.push(row);
        }
        return testItems;
    }
    /**
     * テストの前提条件を生成
     * @param testCase テストケース
     * @returns 前提条件の文字列
     */
    generatePreconditions(testCase) {
        // 前提条件の例
        let preconditions = '・テスト用のアカウントを準備し、ログインしておく';
        // コンポーネントに基づく前提条件を追加
        const componentIndex = testCase.factors.findIndex(f => f === 'コンポーネント');
        if (componentIndex !== -1) {
            const component = testCase.values[componentIndex];
            preconditions += `\n・${component}が表示される画面を開いておく`;
        }
        return preconditions;
    }
    /**
     * テスト手順を生成
     * @param testCase テストケース
     * @returns テスト手順の文字列
     */
    generateTestProcedure(testCase) {
        let procedure = '';
        let stepNumber = 1;
        // 各因子と水準に対応する手順を生成
        for (let i = 0; i < testCase.factors.length; i++) {
            const factor = testCase.factors[i];
            const value = testCase.values[i];
            if (factor === 'コンポーネント') {
                procedure += `${stepNumber}. 「${value}」コンポーネントを表示する\n`;
                stepNumber++;
            }
            else if (factor === '期待値') {
                procedure += `${stepNumber}. 表示された「${value}」を確認する\n`;
                stepNumber++;
            }
            else if (factor.startsWith('期待値:')) {
                // 特殊な期待値の場合
                // 例: "期待値: テキスト複数行のみ" の場合
                const specificContext = factor.replace('期待値:', '').trim();
                procedure += `${stepNumber}. ${specificContext}の「${value}」を確認する\n`;
                stepNumber++;
            }
        }
        return procedure.trim();
    }
    /**
     * 期待結果を生成
     * @param testCase テストケース
     * @returns 期待結果の文字列
     */
    generateExpectedResult(testCase) {
        let result = '';
        let stepNumber = 1;
        // 各因子と水準に対応する期待結果を生成
        for (let i = 0; i < testCase.factors.length; i++) {
            const factor = testCase.factors[i];
            const value = testCase.values[i];
            if (factor === 'コンポーネント') {
                result += `${stepNumber}. 「${value}」コンポーネントが正しく表示されること\n`;
                stepNumber++;
            }
            else if (factor === '期待値') {
                result += `${stepNumber}. 「${value}」の状態が仕様通りであること\n`;
                stepNumber++;
            }
            else if (factor.startsWith('期待値:')) {
                // 特殊な期待値の場合
                const specificContext = factor.replace('期待値:', '').trim();
                result += `${stepNumber}. ${specificContext}の「${value}」が仕様通りに表示・動作すること\n`;
                stepNumber++;
            }
        }
        return result.trim();
    }
    /**
     * 因子と水準の組み合わせを生成する
     * @param factors 因子と水準の配列
     * @returns 組み合わせの配列
     */
    generateCombinations(factors) {
        if (!factors || factors.length === 0) {
            return [];
        }
        // 最初の因子の水準を初期組み合わせとして設定
        let combinations = [];
        const firstFactor = factors[0];
        for (const level of firstFactor.levels) {
            combinations.push({
                factors: [firstFactor.name],
                values: [level]
            });
        }
        // 残りの因子を1つずつ追加
        for (let i = 1; i < factors.length; i++) {
            const currentFactor = factors[i];
            const newCombinations = [];
            // 現在の組み合わせと新しい因子の水準を組み合わせる
            for (const combo of combinations) {
                for (const level of currentFactor.levels) {
                    newCombinations.push({
                        factors: [...combo.factors, currentFactor.name],
                        values: [...combo.values, level]
                    });
                }
            }
            combinations = newCombinations;
        }
        return combinations;
    }
}
