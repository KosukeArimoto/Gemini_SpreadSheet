
/**
 * inputシートのデータを読み込み、APIを使って大分類・中分類を生成し、新しいシートに出力する関数
 */
function generateCategories() {
  const ui = SpreadsheetApp.getUi(); 

  try {
    // --- 2. config情報を変数に格納する ---
    const direction = configSheet.getRange('C3').getValue(); // 今回は 'row' (行方向) の前提で処理
    const prompt1 = promptSheet.getRange(prompt1_pos).getValue();
    console.log("【INFO】direction変数は"+direction)
    console.log("【INFO】sep変数は"+sep)

    if (!direction || !sep || isNaN(sep) || sep <= 0) {
      throw new Error('configシートのC3(方向), C4(分割数)のいずれかが無効です。');
    }

    // inputシートからデータを取得
    const inputSheetName = promptSheet.getRange(inputSheetName_pos).getValue();
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) {
      throw new Error(`データシート「${inputSheetName}」が見つかりません。`);
    }
    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];
    const data = allData.slice(1); // ヘッダーを除いた実データ

    if (data.length === 0) {
      ui.alert(`${inputSheetName}シートにデータがありません。`);
      return;
    }
    
    ss.toast('分類処理を開始します...', '開始', 5);

    let result = []; // 最終的な全分類結果を格納する配列
    let previousResultJsonForPrompt = ""; // 次のプロンプトに含めるための、前回までの結果サマリー

    // --- 3 & 6. inputデータがなくなるまでループ処理 ---
    for (let i = 0; i < data.length; i += sep) {
      const chunk = data.slice(i, i + sep);
      ss.toast(`データを処理中... (${i + chunk.length} / ${data.length})`, 'API連携中', -1);
      
      const chunkWithHeader = [header].concat(chunk);
      const csvChunk = chunkWithHeader.map(row => 
        row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
      ).join('\n');

      let prompt = _replacePrompts(prompt1);
      if (previousResultJsonForPrompt) {
        prompt += `
# 前回までの分類結果の概要
以下は前回までに分類した結果です。この分類基準や粒度を参考にし、必要であれば新たな分類の追加や既存分類の再編をおこなってください。
${previousResultJsonForPrompt}
`;
      }
      prompt += `
# 今回分類するデータ (CSV形式)
---
${csvChunk}
---

上記データの分析結果をJSON配列形式で出力してください。`;

      const resultText = callGemini_(prompt);

      const jsonStringMatch = resultText.match(/```json\s*([\s\S]*?)\s*```/);
      const cleanedJsonString = jsonStringMatch ? jsonStringMatch[1] : resultText;
      result = JSON.parse(cleanedJsonString);
      // result = result.concat(newResults);

      previousResultJsonForPrompt = JSON.stringify(result, null, 2);
    }

    // --- 7. 最終的な結果を新しいシートに出力する ---
    
    // ★★★ここからが修正点：重複を削除する処理★★★
    console.log("【INFO】result変数は"+result);
    const uniqueCategoriesMap = new Map();
    result.forEach(item => {
      // 「大分類」と「中分類」を結合したユニークなキーを作成
      const key = `${item.major_category}|${item.minor_category}`;
      // Mapオブジェクトにキーと値をセット（キーが重複した場合は上書きされる）
      uniqueCategoriesMap.set(key, item);
    });
    // Mapの値だけを取り出して、重複が削除された配列を生成
    const uniqueResult = Array.from(uniqueCategoriesMap.values());
    console.log("【INFO】uniqueResult変数は"+uniqueResult);
    // ★★★修正点はここまで★★★

    ss.toast('結果を出力しています...', '最終処理中', -1);
    const resultSheetName = `分類リスト_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    const resultSheet = ss.insertSheet(resultSheetName, ss.getNumSheets() + 1);
    
    const outputHeader = ['大分類', '中分類'];
    // 重複削除済みの `uniqueResult` を使って出力データを作成
    const outputData = uniqueResult.map(item => [
      item.major_category,
      item.minor_category
    ]);

    // ヘッダーとデータをシートに書き込み
    resultSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
    if (outputData.length > 0) {
      resultSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    }
    resultSheet.autoResizeColumns(1, outputHeader.length);

    ss.toast('分類処理が完了しました！', '成功', 10);

  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message);
  }
}


/**
 * [AI利用版] 
 * 元データと分類リストをGemini APIに渡し、
 * 各データに最適な分類を判断させて付与・出力する関数
 */
function mergeCategories(resultSheetName="") {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('AIによる分類付与を開始します...', '開始', 5);

    // --- 1. 設定情報を取得 ---
    const inputSheetName = promptSheet.getRange(inputSheetName_pos).getValue();
    let categorySheetName = resultSheetName; // まず引数の値で初期化
    if (!categorySheetName) {
      // 引数が空（単体実行）の場合、C8セルから取得
      categorySheetName = promptSheet.getRange(categorySheetName_pos).getValue();
    }
    const categorySheet = ss.getSheetByName(categorySheetName);
    const prompt2 = promptSheet.getRange(prompt2_pos).getValue();

    // --- 2. 元データと分類リストを読み込む ---
    // 元データを取得
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) throw new Error(`入力シート「${inputSheetName}」が見つかりません。`);
    const allOriginalData = inputSheet.getDataRange().getValues();
    const originalHeader = allOriginalData[0];
    const originalData = allOriginalData.slice(1);

    ss.toast(`分類リスト「${categorySheet.getName()}」を使用します。`, '情報', 5);
    const categoryData = categorySheet.getDataRange().getValues();
    categoryData.shift(); // ヘッダーを除外
    const categoryListAsJson = JSON.stringify(
      categoryData.map(row => ({ major_category: row[0], minor_category: row[1] })),
      null, 2
    );

    // --- 3. 元データを分割し、ループ処理 ---
    let finalMergedData = [];
    for (let i = 0; i < originalData.length; i += sep) {
      const chunk = originalData.slice(i, i + sep);
      ss.toast(`AIがデータを分析中... (${i + chunk.length} / ${originalData.length})`, 'API連携中', -1);
      
      const csvChunk = [originalHeader].concat(chunk).map(row => 
        row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
      ).join('\n');

      // --- 4. Gemini APIに投げるプロンプトを作成 ---
      let prompt = _replacePrompts(prompt2);

      prompt += `
# 分類カテゴリのリスト (JSON形式)
利用可能な分類は以下の通りです。このリストの中から最適なものを選択してください。
---
${categoryListAsJson}
---

# 今回割り当てる入力データ (CSV形式)
以下の各データ項目に対して、上記のリストから最も適切と思われる「大分類」と「中分類」を割り当ててください。
---
${csvChunk}
---`


      // --- 5. APIを呼び出し、結果を結合 ---
      const resultText = callGemini_(prompt);

      const cleanedJsonString = resultText.match(/```json\s*([\s\S]*?)\s*```/)?.[1] || resultText;
      const newResults = JSON.parse(cleanedJsonString);
      finalMergedData = finalMergedData.concat(newResults);
    }

    // --- 6. 最終結果を指定のシートに出力 ---
    if (finalMergedData.length === 0) {
      throw new Error("AIからの処理結果が空でした。プロンプトやAPIの応答を確認してください。");
    }
    
    ss.toast('最終結果を出力しています...', '処理中', 5);
    const finalHeader = Object.keys(finalMergedData[0]);
    const outputData = finalMergedData.map(item => finalHeader.map(key => item[key]));
    
    // TODO
    const outputSheetName = `分類付与結果_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);
    // outputシートネームをスプシから取得していた時のコード
    // let outputSheet = ss.getSheetByName(outputSheetName);
    // if (outputSheet) {
    //   outputSheet.clear();
    // } else {
    //   outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);
    // }
    
    outputSheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]).setFontWeight('bold');
    outputSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    outputSheet.autoResizeColumns(1, finalHeader.length);
    
    ui.alert('成功', `シート「${outputSheetName}」にAIによる分類を付与したデータを出力しました。`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message);
  }
}


/**
 * [改善版] 分類付与済みのデータを基に、設計へのフィードバックをAIで生成し、新しいシートに出力する関数
 * ★★★同一カテゴリ内でAIの応答がなくなるまでバッチ処理を繰り返す★★★
 */
function generateFeedback() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('設計フィードバックの生成を開始します...', '開始', 5);

    // --- 1. 設定情報を取得 ---
    const inputSheetName = promptSheet.getRange(outputSheetName_pos).getValue();
    const basePrompt = promptSheet.getRange(prompt3_pos).getValue();
    const inputCategory = configSheet.getRange('C5').getValue();

    if (!inputSheetName || !basePrompt) {
      throw new Error(`promptシートの${inputSheetName}(入力シート名)またはprompt3(プロンプト)が空です。`);
    }

    // --- 2. 入力データを読み込む ---
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) throw new Error(`入力シート「${inputSheetName}」が見つかりません。`);
    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];
    const data = allData.slice(1);

    if (data.length === 0) {
      throw new Error(`入力シート「${inputSheetName}」にデータがありません。`);
    }

    // --- 3. 指定された列でデータをグループ化 ---
    const categoryIndex = header.indexOf(inputCategory);
    if (categoryIndex === -1) {
      throw new Error(`入力シートのヘッダーに「${inputCategory}」列が見つかりません。`);
    }
    const groupedData = {};
    data.forEach(row => {
      const category = row[categoryIndex];
      if (!groupedData[category]) {
        groupedData[category] = [];
      }
      groupedData[category].push(row);
    });

    // --- 4. グループ化したチャンクごとにループ処理（外側ループ） ---
    let combinedMarkdownResponse = "";
    let previousFeedbackForPrompt = "";
    
    const categories = Object.keys(groupedData);
    let processedCategories = 0;

    for (const categoryName of categories) {
      const chunk = groupedData[categoryName];
      processedCategories++;
      
      const csvChunk = [header].concat(chunk).map(row => 
        row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
      ).join('\n');
      
      let continueProcessingCategory = true;
      let batchNumber = 1;

      while (continueProcessingCategory) {
        ss.toast(`[${processedCategories}/${categories.length}] カテゴリ「${categoryName}」を分析中 (バッチ ${batchNumber})...`, 'API連携中', -1);

        let prompt = _replacePrompts(basePrompt);
        if (previousFeedbackForPrompt) {
          prompt += `\n\n---
        # 🔴 重要：既に出力済みのフィードバック
        以下は既に出力したフィードバックです。
        新たなフィードバックはこのフィードバックに追加する形式で出力してください。
        **「# 🔁 重複防止条件」**のルールに厳密に従い、これらと重複する内容は絶対に出力しないでください。
        ${previousFeedbackForPrompt}`;
                }
                prompt += `\n\n---

        # 出力形式の追加説明
        ヘッダー自体は出力しないでください。

        # 今回分析する入力データ (CSV形式)
        ${csvChunk}`;

        const resultText = callGemini_(prompt);
        console.log(resultText)
        
        combinedMarkdownResponse += resultText + "\n";
        previousFeedbackForPrompt += resultText + "\n";
        batchNumber++;

        const newFeedbackData = parseMarkdownTable_(resultText);
        if (newFeedbackData.length <= 1 || resultText.includes('続きなし')) {
          continueProcessingCategory = false;
        }
      }
      Utilities.sleep(1000); // ★★★1秒間待機する★★★
    }

    // --- 8. 最終結果をシートに出力 ---
    ss.toast('最終結果を出力しています...', '処理中', 5);
    const feedbackData = parseMarkdownTable_(combinedMarkdownResponse);

    if (feedbackData.length === 0) {
      throw new Error("AIの応答からテーブルデータを抽出できませんでした。");
    }

    // ★★★ここからが修正点：重複したヘッダー行を削除する処理★★★
    const headerRow = feedbackData[0]; // 最初の行をヘッダーとして取得
    const headerString = headerRow.join('|'); // 比較用の文字列を作成

    // 最初のヘッダー行と、ヘッダーと一致しないデータ行だけをフィルタリング
    const uniqueHeaderData = feedbackData.filter((row, index) => {
      return index === 0 || row.join('|') !== headerString;
    });
    // ★★★修正点はここまで★★★

    const outputSheetName = `設計FB_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);
    
    // 重複ヘッダーを削除した `uniqueHeaderData` を使って書き込み
    outputSheet.getRange(1, 1, uniqueHeaderData.length, uniqueHeaderData[0].length)
      .setValues(uniqueHeaderData)
      .setWrap(true)
      .setVerticalAlignment('top');
      
    outputSheet.getRange(1, 1, 1, uniqueHeaderData[0].length).setFontWeight('bold');
    outputSheet.autoResizeColumns(1, uniqueHeaderData[0].length);
    
    ui.alert('成功', `シート「${outputSheetName}」に設計FBを出力しました。`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message);
  }
}


/**
 * 生成済みの「設計フィードバック」を指定された指示で修正し、
 * 新しいシートに改訂版として出力する関数
 */
function reviseFeedback() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('設計フィードバックの改訂処理を開始します...', '開始', 5);

    // --- 1. 設定情報と入力データをすべて読み込む ---
    const revisionSheet = ss.getSheetByName('形式知修正');
    if (!revisionSheet) throw new Error('シート「形式知修正」が見つかりません。');

    // 設定値を取得
    const feedbackSheetName = revisionSheet.getRange('C6').getValue();
    const rawDataSheetName = promptSheet.getRange(inputSheetName_pos).getValue();
    const feedbackRule = promptSheet.getRange(prompt3_pos).getValue();

    // 修正対象のリストを取得 (B12, C12から最終行まで)
    const revisionList = revisionSheet.getRange('B12:C' + revisionSheet.getLastRow()).getValues()
      .filter(row => row[0] && row[1]); // 番号と指示の両方が入力されている行のみを対象

    if (revisionList.length === 0) {
      ui.alert('「形式知修正」シートに、修正対象のフィードバック番号と変更要望が入力されていません。');
      return;
    }

    // --- 2. データを高速に参照できるよう、Mapオブジェクトに変換 ---
    // 対象フィードバックシートのデータを読み込み
    const feedbackSheet = ss.getSheetByName(feedbackSheetName);
    if (!feedbackSheet) throw new Error(`対象フィードバックシート「${feedbackSheetName}」が見つかりません。`);
    const feedbackData = feedbackSheet.getDataRange().getValues();
    const feedbackHeader = feedbackData.shift();
    const feedbackMap = new Map(feedbackData.map(row => [String(row[0]), row])); // Map<フィードバック番号, 行データ>

    // 元の入力データを読み込み
    const rawDataSheet = ss.getSheetByName(rawDataSheetName);
    if (!rawDataSheet) throw new Error(`大元の入力シート「${rawDataSheetName}」が見つかりません。`);
    const rawData = rawDataSheet.getDataRange().getValues();
    const rawDataHeader = rawData.shift();
    const rawDataMap = new Map(rawData.map(row => [String(row[0]), row])); // Map<通し番号, 行データ>

    // --- 3. 修正リストを順番に処理するループ ---
    let revisedFeedbackResults = [];
    let processCount = 0;
    for (const revision of revisionList) {
      const feedbackNumber = String(revision[0]);
      const revisionPrompt = revision[1];
      processCount++;
      ss.toast(`[${processCount}/${revisionList.length}] フィードバック番号「${feedbackNumber}」を修正中...`, 'API連携中', -1);

      // Mapから元のフィードバックデータを取得
      const originalFeedbackRow = feedbackMap.get(feedbackNumber);
      if (!originalFeedbackRow) {
        console.warn(`フィードバック番号「${feedbackNumber}」が見つかりませんでした。スキップします。`);
        continue;
      }
      const baseSerialNumbers = String(originalFeedbackRow[4]).split(/[\n,]/).map(s => s.trim()); // ベース通し番号を取得

      // 元の入力データをMapから取得
      let referencedRawData = "";
      baseSerialNumbers.forEach(serialNumber => {
        const rawRow = rawDataMap.get(serialNumber);
        if (rawRow) {
          referencedRawData += rawDataHeader.join(',') + '\n' + rawRow.join(',') + '\n\n';
        }
      });
      
      // --- 4. AIへのプロンプトを構築 ---
      const finalPrompt = `
# あなたの役割
あなたは「自動車向けワイヤーハーネス設計のシニアエンジニア」です。一度作成した設計フィードバックを、追加の指示に基づき、より高品質なものに改訂する専門家として振る舞ってください。

# 元の設計フィードバック
以下は今回修正する対象のフィードバックです。
- フィードバック番号: ${feedbackNumber}
- フィードバックタイトル: ${originalFeedbackRow[1]}
- フィードバック概要: ${originalFeedbackRow[2]}
- フィードバック詳細: ${originalFeedbackRow[3]}

# 修正指示
以下の指示に従って、上記のフィードバックを改訂してください。
「${revisionPrompt}」

フィードバック生成ルールは以下に記載の内容に従うこと。
「${feedbackRule}」

# 参照情報
このフィードバックの元となったデータは以下の通りです。この内容をよく読んだ上で、修正指示を反映してください。
${referencedRawData}

# 出力形式
改訂後のフィードバックを、以下のJSONオブジェクト形式で出力してください。キーの名前と順番は厳密に守ってください。
{
  "フィードバック番号": "${feedbackNumber}",
  "フィードバックタイトル": "（改訂後のタイトル）",
  "フィードバック概要": "（改訂後の概要）",
  "フィードバック詳細": "（改訂後の詳細）",
  "ベース通し番号": "${originalFeedbackRow[4]}",
  "ベース概要（管理番号）": "（改訂後のベース概要）"
}`;

      // --- 5. APIを呼び出し、結果を格納 ---
      const resultText = callGemini_(finalPrompt);

      console.log("resultTextの内容は"+resultText)
      const cleanedJsonString = resultText.match(/```json\s*([\s\S]*?)\s*```/)?.[1] || resultText;
      console.log("cleanedJsonStringの内容は"+cleanedJsonString)
      const revisedFeedback = JSON.parse(cleanedJsonString);
      console.log("revisedFeedbackの内容は"+revisedFeedback)
      revisedFeedbackResults.push(revisedFeedback);
    }

    // --- 6. 最終結果を新しいシートに出力 ---
    if (revisedFeedbackResults.length === 0) {
      ui.alert('改訂されたフィードバックがありませんでした。');
      return;
    }

    ss.toast('最終結果を出力しています...', '処理中', 5);
    const outputSheetName = `改訂版FB_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);

    const outputHeader = Object.keys(revisedFeedbackResults[0]);
    const outputData = revisedFeedbackResults.map(item => outputHeader.map(key => item[key]));
    
    outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
    outputSheet.getRange(2, 1, outputData.length, outputData[0].length)
      .setValues(outputData)
      .setWrap(true)
      .setVerticalAlignment('top');
      
    outputSheet.autoResizeColumns(1, outputHeader.length);
    
    ui.alert('成功', `シート「${outputSheetName}」に改訂版FBを出力しました。`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message);
  }
}


/**
 * [改善版] inputシートのデータを読み込み、AIを使って分析・抽出し、
 * ★★★プロンプトで定義された自由な形式で★★★新しいシートに出力する関数
 */
function freePrompt() {
  const ui = SpreadsheetApp.getUi();

  try {
    // --- 1. 設定情報を取得 ---
    const inputSheetName = freePromptSheet.getRange("C6").getValue();
    const basePrompt = freePromptSheet.getRange("C25").getValue(); // 出力形式を定義するプロンプト

    if (!sep || isNaN(sep) || !inputSheetName || !basePrompt) {
      throw new Error('configシート(C4)またはfree promptシート(C6, C25)の設定が不足しています。');
    }

    // --- 2. 入力データを読み込む ---
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) throw new Error(`データシート「${inputSheetName}」が見つかりません。`);
    
    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];
    const data = allData.slice(1);

    if (data.length === 0) {
      ui.alert(`${inputSheetName}シートにデータがありません。`);
      return;
    }
    
    ss.toast('分析処理を開始します...', '開始', 5);

    let allResults = []; // 全てのチャンクからの結果を格納する配列
    let previousResultJsonForPrompt = "";

    // --- 3. ループ処理 ---
    for (let i = 0; i < data.length; i += sep) {
      const chunk = data.slice(i, i + sep);
      ss.toast(`データを処理中... (${i + chunk.length} / ${data.length})`, 'API連携中', -1);
      
      const chunkWithHeader = [header].concat(chunk);
      const csvChunk = chunkWithHeader.map(row => 
        row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
      ).join('\n');

      // --- 4. プロンプトを構築 ---
      let prompt = basePrompt; // ユーザーが定義したプロンプトをベースにする
      if (previousResultJsonForPrompt) {
        prompt += `\n\n# 前回までの出力結果の概要\n以下は前回までに出力した結果です。この内容や形式を参考に、一貫性を保ってください。\n${previousResultJsonForPrompt}`;
      }
      prompt += `\n\n# 今回分析するデータ (CSV形式)\n---\n${csvChunk}`;

// --- 5. APIを呼び出し、結果を安全に解析・結合 ---
      const resultText = callGemini_(prompt);
      let jsonToParse = "";

      try {
        // 戦略1: まず ```json ... ``` のマークダウンブロックを探す
        const jsonStringMatch = resultText.match(/```json\s*([\s\S]*?)\s*```/);
        
        if (jsonStringMatch && jsonStringMatch[1]) {
          // マークダウンブロックが見つかった場合
          jsonToParse = jsonStringMatch[1];
        } else {
          // 戦略2: マークダウンがない場合、「承知しました」などの前置きを無視し、
          // 応答から { または [ で始まる最初のJSON部分を探す
          
          const startIndex = resultText.indexOf('{');
          const arrayStartIndex = resultText.indexOf('[');
          
          let jsonStartIndex = -1;

          // 最初に出現する { か [ を見つける
          if (startIndex !== -1 && arrayStartIndex !== -1) {
            jsonStartIndex = Math.min(startIndex, arrayStartIndex);
          } else if (startIndex !== -1) {
            jsonStartIndex = startIndex;
          } else if (arrayStartIndex !== -1) {
            jsonStartIndex = arrayStartIndex;
          }

          if (jsonStartIndex !== -1) {
            // JSONの開始文字が見つかった場合
            // 対応する最後の } または ] を探す
            const startChar = resultText[jsonStartIndex];
            const endChar = (startChar === '{') ? '}' : ']';
            
            const jsonEndIndex = resultText.lastIndexOf(endChar);
            
            if (jsonEndIndex > jsonStartIndex) {
              // 最初 {/[ から 最後 }/] までを切り出す
              jsonToParse = resultText.substring(jsonStartIndex, jsonEndIndex + 1);
            } else {
              // 開始文字しか見つからなかった場合 (異常だが念のため)
              jsonToParse = resultText.substring(jsonStartIndex);
            }
          } else {
            // { も [ も ```json も見つからなかった
            // この場合は解析エラーになるが、catchで処理される
            jsonToParse = resultText;
          }
        }
        
        // 抽出した文字列を解析
        if (jsonToParse.trim() !== "") {
          const newResults = JSON.parse(jsonToParse);
          allResults = allResults.concat(newResults);
        }

      } catch (e) {
        // ログが大きすぎるとエラーになるため、API応答を短縮して記録する
        const truncatedResponse = resultText.substring(0, 5000);
        console.error(`JSON解析エラー。このチャンクをスキップします。API応答(先頭5000文字): ${truncatedResponse}`, e);
        continue;
      }

      // ログが大きすぎる問題を避けるため、次回のプロンプトに渡す概要を短縮する
      if (allResults.length > 0) {
        // 最新の5件だけを概要として渡す (例)
        const recentResults = allResults.slice(-5);
        previousResultJsonForPrompt = JSON.stringify(recentResults, null, 2);
      } else {
        previousResultJsonForPrompt = "";
      }

      // --- 6. 最終結果を動的に解釈してシートに出力 ---
      if (allResults.length === 0) {
        throw new Error("AIからの有効な応答がありませんでした。");
      }

      ss.toast('結果を出力しています...', '最終処理中', -1);

      // ★★★ここからが改善点：結果から動的にヘッダーとデータを作成★★★
      const outputHeader = Object.keys(allResults[0]); // 最初の結果オブジェクトからキーを取得してヘッダーにする
      const outputData = allResults.map(item => {
        return outputHeader.map(key => item[key] || ""); // ヘッダーの順に値を取得。存在しない場合は空文字
      });
      // ★★★改善点はここまで★★★

      const resultSheetName = `分析結果_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
      const resultSheet = ss.insertSheet(resultSheetName, ss.getNumSheets() + 1);
      
      resultSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
      if (outputData.length > 0) {
        resultSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
      }
      resultSheet.autoResizeColumns(1, outputHeader.length);

      ss.toast('処理が完了しました！', '成功', 10);

    }
  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message);
  }
}

/**
 * 「設計フィードバック」シートの内容を基に、イラスト作成用の
 * 「OK事例」「NG事例」をAIで生成し、新しいシートに出力する関数
 */
function createIllustrationPrompts() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('イラスト用プロンプトの生成を開始します...', '開始', 5);

    // --- 1. 設定情報を取得 ---
    const feedbackSheetName = promptSheet.getRange(feedbackSheetName_pos).getValue();
    const prompt4 = promptSheet.getRange(prompt4_pos).getValue();
    
    // ★★★ C10セルから読み込む列指定を追加 ★★★
    const columnsString = promptSheet.getRange('C10').getValue(); // 例: "A, B, E"

    // --- 2. 入力データを読み込む ---
    const feedbackSheet = ss.getSheetByName(feedbackSheetName);
    if (!feedbackSheet) throw new Error(`対象フィードバックシート「${feedbackSheetName}」が見つかりません。`);
    
    const allData = feedbackSheet.getDataRange().getValues();
    const header = allData[0];
    const data = allData.slice(1);

    // const testData = data.slice(0, 4); // テスト用に2行に絞る
    const testData = data; // 本番用

    if (testData.length === 0) {
      throw new Error(`入力シート「${feedbackSheetName}」にデータがありません。`);
    }

    // --- 3. 処理に必要な列のインデックスと列名を特定 ★★★ここから修正★★★
    let columnIndices;
    if (columnsString) {
      // C10に指定がある場合：新しいヘルパー関数で解析
      columnIndices = _parseColumnRangeString(columnsString);
      if (columnIndices.length === 0) {
        throw new Error('promptシートC10セルの列指定が有効ではありませんでした。');
      }
    } else {
      // C10が空白の場合：全ての列を対象とする
      columnIndices = header.map((_, index) => index); // [0, 1, 2, ..., n-1] を生成
    }
    
    // 取得したインデックスを使って、ヘッダー（列名）のリストを作成
    const columnsToUse = columnIndices.map(index => {
        if (index < 0 || index >= header.length) {
            throw new Error(`列指定 ${index + 1} がシートの範囲外です。`); // 無効なインデックスがあればエラー
        }
        return header[index];
    });
    // ★★★修正点はここまで★★★

    // --- 4. 1行ずつループ処理 ---
    let finalOutputRows = [];
    let processCount = 0;
    
    // ベースとなるプロンプトを準備 (静的な置換はここで完了)
    const basePrompt = _replacePrompts(prompt4);

    for (const row of testData) {
      processCount++;
      ss.toast(`[${processCount}/${testData.length}] フィードバックを処理中...`, 'API連携中', -1);

      // --- 5. プロンプトに含めるフィードバック内容を動的に構築 ★★★ここから修正★★★
      let feedbackContent = "";
      columnsToUse.forEach((colName, i) => {
        const dataIndex = columnIndices[i]; // 取得するデータのインデックス
        feedbackContent += `- ${colName}: ${row[dataIndex]}\n`;
      });
      
      // プロンプト内の動的なプレースホルダーを置換
      let finalPrompt = basePrompt + feedbackContent;

      console.log(finalPrompt);
      // --- 6. APIを呼び出し、結果を解析 ---
      const resultText = callGemini_(finalPrompt);

      const parsedTable = parseMarkdownTable_(resultText);
      let okCase = "（生成失敗）";
      let ngCase = "（生成失敗）";
      if (parsedTable.length > 1) { 
        okCase = parsedTable[1][1] || okCase;
        ngCase = parsedTable[1][2] || ngCase;
      }

      finalOutputRows.push(row.concat([okCase, ngCase]));
    } // --- ループここまで ---

    // --- 7. 最終結果を新しいシートに出力 ---
    ss.toast('最終結果を出力しています...', '処理中', 5);
    
    const outputSheetName = `イラストプロンプト案_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);
    
    const outputHeader = header.concat(['OK事例', 'NG事例']);
    
    outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
    if (finalOutputRows.length > 0) {
      outputSheet.getRange(2, 1, finalOutputRows.length, finalOutputRows[0].length)
        .setValues(finalOutputRows)
        .setWrap(true)
        .setVerticalAlignment('top');
    }
    
    outputSheet.autoResizeColumns(1, outputHeader.length);
    
    ui.alert('成功', `シート「${outputSheetName}」にイラスト用プロンプト案を出力しました。`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message);
  }
}

/**
 * [改善版] 「イラストプロンプト案」シートを基に、イラストを一括生成し、
 * ★★★指定されたGoogle Driveフォルダに画像を保存しつつ、シートにも画像を挿入する★★★関数
 */
function createImages() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('イラストの一括生成を開始します...', '開始', 5);

    // --- 1. 設定情報を取得 ---
    const imagePromptSheetName = promptSheet.getRange(imagePromptSheetName_pos).getValue();
    const promt5 = promptSheet.getRange(prompt5_pos).getValue();
    const outputFolderUrl = promptSheet.getRange(imageSaveDir_pos).getValue(); // 保存先フォルダURL

    const numberOfGenerations = parseInt(configSheet.getRange('C6').getValue(), 10) || 1;
    if (numberOfGenerations <= 0) {
      ui.alert('エラー', 'configシートC6セルの「生成枚数」は1以上の数値を入力してください。', ui.ButtonSet.OK);
      return;
    }

    // --- 1b. 保存先フォルダを特定 ---
    let outputFolder;
    if (outputFolderUrl) {
      const folderId = _extractFolderIdFromUrl(outputFolderUrl);
      if (folderId) {
        try {
          outputFolder = DriveApp.getFolderById(folderId);
          Logger.log(`保存先フォルダを指定: ${outputFolder.getName()} (ID: ${folderId})`);
        } catch (e) {
          throw new Error(`指定されたフォルダURL(ID: ${folderId})が見つからないかアクセスできません。処理を中止します。`);
        }
      } else {
        throw new Error(`promptシートC13セルのURLから有効なフォルダIDを取得できませんでした。処理を中止します。`);
      }
    } else {
      throw new Error(`promptシートC13セルに保存先フォルダのURLが指定されていません。処理を中止します。`);
    }

    // --- 2. 入力シートを準備 ---
    const sheet = ss.getSheetByName(imagePromptSheetName);
    if (!sheet) throw new Error(`シート「${imagePromptSheetName}」が見つかりません。`);

    const allData = sheet.getDataRange().getValues();
    let header = allData[0];
    const dataRows = allData.slice(1);

    const filterString = promptSheet.getRange(imageTargetNum_pos).getValue();
    let dataToProcess = [];

    if (filterString) {
      const targetNumbers = new Set(_parseNumberRangeString(filterString));
      const serialNumberIndex = 0;
      dataRows.forEach((row, index) => {
        const serialNumber = parseInt(row[serialNumberIndex], 10);
        if (targetNumbers.has(serialNumber)) {
          dataToProcess.push({ rowData: row, rowIndex: index + 2, serialNumber: String(row[serialNumberIndex]) });
        }
      });
    } else {
      dataToProcess = dataRows.map((row, index) => ({
        rowData: row,
        rowIndex: index + 2,
        serialNumber: String(row[0])
      }));
    }

    if (dataToProcess.length === 0) {
      ui.alert('処理対象のデータが見つかりませんでした。');
      return;
    }

    // --- 2b. 新しい列（画像 + URL）を準備 --- ★ ここから修正 ★
    const existingImageCols = header.filter(h => h.toString().startsWith('生成画像'));
    const firstNewColIndex = header.length;
    let newHeaders = [];
    let newHeaderIndices = {}; // { '生成画像_1': index, '生成画像URL_1': index, ... }

    for (let i = 0; i < numberOfGenerations; i++) {
      const colNumber = existingImageCols.length / 2 + i + 1; // 画像とURLのペアで数える
      const imageHeaderName = colNumber === 1 ? '生成画像' : `生成画像_${colNumber}`;
      newHeaders.push(imageHeaderName);
      newHeaderIndices[imageHeaderName] = firstNewColIndex + (i * 2); // 画像列のインデックス
    }

    if (newHeaders.length > 0) {
      sheet.getRange(1, firstNewColIndex + 1, 1, newHeaders.length).setValues([newHeaders]).setFontWeight('bold');
      header = header.concat(newHeaders); // メモリ上のヘダー配列も更新
    }
    // ★ ここまで修正 ★

    const okCaseIndex = header.indexOf('OK事例');
    const ngCaseIndex = header.indexOf('NG事例');
    if (okCaseIndex === -1 || ngCaseIndex === -1) {
      throw new Error('入力シートに「OK事例」または「NG事例」の列が見つかりません。');
    }

    const testData = dataToProcess; // 全データ対象に変更

    // --- 3. ループ処理 ---
    const basePrompt = _replacePrompts(promt5);
    let processCount = 0;

    for (const item of testData) {
      const { rowData, rowIndex, serialNumber } = item;
      processCount++;

      const okCase = rowData[okCaseIndex];
      const ngCase = rowData[ngCaseIndex];

      let finalPrompt = basePrompt
        .replace('<NG_Image>', ngCase)
        .replace('<OK_Image>', okCase);

      // --- 4. 指定された回数だけAPIを呼び出し、画像をDriveに保存 & シートに挿入 --- ★ ここから修正 ★
      for (let j = 0; j < numberOfGenerations; j++) {
        const colNumber = existingImageCols.length / 2 + j + 1;
        const imageHeaderName = colNumber === 1 ? '生成画像' : `生成画像_${colNumber}`;
        const currentImageColIndex = newHeaderIndices[imageHeaderName]; // 画像を挿入する列インデックス

        ss.toast(`[${processCount}/${testData.length}] 画像 ${colNumber} を生成中 (No.${serialNumber})...`, 'API連携中', -1);

        const base64Image = callGPTApi_(finalPrompt); // DALL·E APIを呼び出し

        // (1) Driveに保存
        const imageName = `${imagePromptSheetName}_No${serialNumber}_${imageHeaderName}.png`;
        let savedFileUrl = '';
        try {
          const decodedBytes = Utilities.base64Decode(base64Image);
          const imageBlob = Utilities.newBlob(decodedBytes, MimeType.PNG, imageName);
          const savedFile = outputFolder.createFile(imageBlob);
          savedFileUrl = savedFile.getUrl(); // 保存したファイルのURLを取得
          Logger.log(`画像を保存しました: ${savedFile.getName()} (URL: ${savedFileUrl})`);
        } catch (saveError) {
          Logger.log(`エラー: No.${serialNumber} の画像 ${colNumber} の保存に失敗しました - ${saveError}`);
          savedFileUrl = '保存失敗'; // エラー情報をURL列に記録
        }

        // (2) シートに画像を挿入
        const dataUrl = `data:image/png;base64,${base64Image}`;
        const cellImage = SpreadsheetApp.newCellImage().setSourceUrl(dataUrl).build();
        sheet.getRange(rowIndex, currentImageColIndex + 1).setValue(cellImage);


        if (numberOfGenerations > 1) {
          Utilities.sleep(1000);
        }
      }
      sheet.setRowHeight(rowIndex, 200); // 行高さを調整
      // ★ ここまで修正 ★
    }

    ss.toast('すべてのイラスト生成・保存が完了しました。', '完了', 5);
    ui.alert('成功', `イラストの一括生成とDriveフォルダへの保存が完了しました。\n保存先: ${outputFolder.getName()}`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message, ui.ButtonSet.OK);
  }
}

function _replacePrompts(originalPrompt) {
  // B14からC22までの置換リストを一度に取得
  const replacements = promptSheet.getRange('B20:C28').getValues();

  let finalPrompt = originalPrompt;

  // 取得したリストを1行ずつループ処理
  for (const row of replacements) {
    const wordToReplace = row[0]; // B列の値
    const replacementValue = row[1]; // C列の値

    // B列に置換する単語が入力されている場合のみ処理を実行
    if (wordToReplace) {
      // {word} の形式のプレースホルダーを全て置換する (RegExpの'g'フラグ)
      const placeholder = new RegExp(`{${wordToReplace}}`, 'g');
      finalPrompt = finalPrompt.replace(placeholder, replacementValue);
    }
  }

  return finalPrompt;
}

/**
* [補助関数] AIが生成したMarkdownテーブル形式のテキストを解析し、
* スプレッドシート用の2次元配列に変換する
*/
function parseMarkdownTable_(markdownText) {
  const lines = markdownText.split('\n');
  const tableData = [];

  for (const line of lines) {
    // "|" を含み、ヘッダーの区切り線 "---" を含まない行をテーブルの行とみなす
      if (line.includes('|') && !line.includes('---')) {
        const cells = line.split('|')
        .map(cell => cell.trim().replace(/<br>/g, '\n'))  // 各セルの前後の空白を削除。セル内改行するように置換
        .slice(1, -1); // 先頭と末尾の空の要素を削除

        if (cells.length > 0) {
          tableData.push(cells);
      }
    }
  }
  return tableData;
}

/**
 * [新規] 列指定文字列（例: "A, C, E-G"）を0ベースのインデックス配列（例: [0, 2, 4, 5, 6]）に変換する
 * @param {string} rangeString - 列指定文字列
 * @return {number[]} - 0ベースの列インデックスの配列
 */
function _parseColumnRangeString(rangeString) {
  const indices = new Set(); // 重複を自動で除く
  const parts = rangeString.split(',');

  for (const part of parts) {
    const trimmedPart = part.trim().toUpperCase(); // 大文字に統一
    if (trimmedPart.includes('-')) {
      const [startLetter, endLetter] = trimmedPart.split('-');
      const startIndex = _columnToIndex(startLetter);
      const endIndex = _columnToIndex(endLetter);
      if (startIndex !== -1 && endIndex !== -1 && startIndex <= endIndex) {
        for (let i = startIndex; i <= endIndex; i++) {
          indices.add(i);
        }
      } else {
        Logger.log(`警告: 無効な列範囲 "${trimmedPart}" は無視されました。`);
      }
    } else {
      const index = _columnToIndex(trimmedPart);
      if (index !== -1) {
        indices.add(index);
      } else {
         Logger.log(`警告: 無効な列指定 "${trimmedPart}" は無視されました。`);
      }
    }
  }
  // Setをソートされた数値配列に変換して返す
  return Array.from(indices).sort((a, b) => a - b);
}

// _columnToIndex 関数も少し修正（無効な文字の場合 -1 を返すように）
function _columnToIndex(columnLetter) {
  let index = 0;
  columnLetter = columnLetter.toUpperCase();
  if (!/^[A-Z]+$/.test(columnLetter)) { // アルファベット以外は無効
      return -1;
  }
  for (let i = 0; i < columnLetter.length; i++) {
    index = index * 26 + (columnLetter.charCodeAt(i) - 64);
  }
  return index - 1;
}
/**
 * [新規] カンマ区切りとハイフンつなぎの数字の文字列（例: "1, 3, 5-9"）を
 * 数値の配列（例: [1, 3, 5, 6, 7, 8, 9]）に変換するヘルパー関数
 * @param {string} rangeString - 変換対象の文字列
 * @return {number[]} - 数値の配列
 */
function _parseNumberRangeString(rangeString) {
  const numbers = new Set(); // 重複を自動で除くためにSetを使用
  const parts = rangeString.split(',');

  for (const part of parts) {
    const trimmedPart = part.trim();
    if (trimmedPart.includes('-')) {
      const [start, end] = trimmedPart.split('-').map(Number);
      if (!isNaN(start) && !isNaN(end) && start <= end) {
        for (let i = start; i <= end; i++) {
          numbers.add(i);
        }
      }
    } else {
      const num = Number(trimmedPart);
      if (!isNaN(num)) {
        numbers.add(num);
      }
    }
  }
  return Array.from(numbers); // Setを配列に変換して返す
}

/**
 * [ヘルパー関数] Google DriveのフォルダURLからフォルダIDを抽出する
 * @param {string} folderUrl - Google DriveのフォルダURL
 * @return {string | null} - フォルダID、見つからない場合はnull
 */
function _extractFolderIdFromUrl(folderUrl) {
  if (!folderUrl || typeof folderUrl !== 'string') return null;
  let id = null;
  // 標準的なフォルダURL (.../folders/ID)
  let match = folderUrl.match(/folders\/([a-zA-Z0-9_-]{25,})/);
  if (match && match[1]) {
    id = match[1];
  } else {
    // 共有リンクURL (...?id=ID)
    match = folderUrl.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
    if (match && match[1]) {
      id = match[1];
    }
  }
  // Google DriveのIDは通常25文字以上
  return (id && id.length >= 25) ? id : null;
}

/**
 * 「カテゴリごとに知見生成」シートの設定に基づいて、
 * 指定されたシートの各行を画像生成用プロンプトとして使用し、
 * 生成した画像を各行の最右列に挿入する関数
 *
 * 設定:
 * - C33セル: 画像生成用のベースプロンプト
 * - C18セル: 画像生成対象のシート名
 */
function generateRowImages() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('行ごとの画像生成を開始します...', '開始', 5);

    // --- 1. 設定情報を取得 ---
    const knowledgeSheet = ss.getSheetByName('カテゴリごとに知見生成');
    if (!knowledgeSheet) {
      throw new Error('シート「カテゴリごとに知見生成」が見つかりません。');
    }

    // C33: 画像生成用のベースプロンプト
    const basePrompt = knowledgeSheet.getRange('C33').getValue();
    if (!basePrompt) {
      throw new Error('C33セルに画像生成用のプロンプトが設定されていません。');
    }

    // C18: 画像生成対象のシート名
    const targetSheetName = knowledgeSheet.getRange('C18').getValue();
    if (!targetSheetName) {
      throw new Error('C18セルに画像生成対象のシート名が設定されていません。');
    }

    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      throw new Error(`画像生成対象シート「${targetSheetName}」が見つかりません。`);
    }

    // 保存先フォルダURL（オプション）- promptシートのC13セルから取得
    const outputFolderUrl = promptSheet.getRange(imageSaveDir_pos).getValue();
    let outputFolder = null;
    if (outputFolderUrl) {
      const folderId = _extractFolderIdFromUrl(outputFolderUrl);
      if (folderId) {
        try {
          outputFolder = DriveApp.getFolderById(folderId);
          Logger.log(`保存先フォルダを指定: ${outputFolder.getName()} (ID: ${folderId})`);
        } catch (e) {
          Logger.log(`警告: 指定されたフォルダにアクセスできません。画像はシートにのみ挿入されます。`);
        }
      }
    }

    // --- 2. 対象シートのデータを読み込む ---
    const allData = targetSheet.getDataRange().getValues();
    if (allData.length === 0) {
      throw new Error(`シート「${targetSheetName}」にデータがありません。`);
    }

    const header = allData[0];
    const dataRows = allData.slice(1);

    if (dataRows.length === 0) {
      throw new Error(`シート「${targetSheetName}」にデータ行がありません（ヘッダーのみ）。`);
    }

    // --- 3. 画像を挿入する列を特定（最右列の次）---
    const imageColumnIndex = header.length + 1; // 1-indexed
    const imageHeaderName = '生成画像';

    // ヘッダーに「生成画像」列を追加（まだ存在しない場合）
    const existingImageHeader = targetSheet.getRange(1, imageColumnIndex).getValue();
    if (!existingImageHeader || existingImageHeader !== imageHeaderName) {
      targetSheet.getRange(1, imageColumnIndex).setValue(imageHeaderName).setFontWeight('bold');
    }

    // --- 4. 各行をループ処理して画像を生成 ---
    let processCount = 0;
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      const rowIndex = i + 2; // シート上の行番号（1-indexed、ヘッダー分+1）
      processCount++;

      // 行データをCSV形式の文字列に変換
      const rowCsvString = row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',');
      const rowWithHeaderCsv = header.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',') + '\n' + rowCsvString;

      // プロンプトを構築
      const finalPrompt = `${basePrompt}

# 入力データ（CSV形式）
以下のデータを基に画像を生成してください。
---
${rowWithHeaderCsv}
---`;

      ss.toast(`[${processCount}/${dataRows.length}] 行${rowIndex}の画像を生成中...`, 'API連携中', -1);

      try {
        // 画像生成APIを呼び出し
        const base64Image = callGPTApi_(finalPrompt);

        // (1) Google Driveに保存（フォルダが指定されている場合）
        if (outputFolder) {
          try {
            const imageName = `${targetSheetName}_行${rowIndex}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMddHHmmss')}.png`;
            const decodedBytes = Utilities.base64Decode(base64Image);
            const imageBlob = Utilities.newBlob(decodedBytes, 'image/png', imageName);
            const savedFile = outputFolder.createFile(imageBlob);
            Logger.log(`画像を保存: ${savedFile.getName()} (URL: ${savedFile.getUrl()})`);
          } catch (saveError) {
            Logger.log(`警告: 行${rowIndex}の画像保存に失敗しました - ${saveError}`);
          }
        }

        // (2) シートに画像を挿入
        const dataUrl = `data:image/png;base64,${base64Image}`;
        const cellImage = SpreadsheetApp.newCellImage().setSourceUrl(dataUrl).build();
        targetSheet.getRange(rowIndex, imageColumnIndex).setValue(cellImage);

        // 行の高さを調整
        targetSheet.setRowHeight(rowIndex, 200);

        // APIレート制限対策のため待機
        if (i < dataRows.length - 1) {
          Utilities.sleep(1000);
        }

      } catch (imageError) {
        Logger.log(`エラー: 行${rowIndex}の画像生成に失敗しました - ${imageError}`);
        targetSheet.getRange(rowIndex, imageColumnIndex).setValue('生成失敗');
      }
    }

    // --- 5. 完了メッセージ ---
    ss.toast('すべての画像生成が完了しました。', '完了', 5);
    const folderMsg = outputFolder ? `\n保存先: ${outputFolder.getName()}` : '';
    ui.alert('成功', `シート「${targetSheetName}」の各行に対する画像生成が完了しました。${folderMsg}`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ss.toast('エラーが発生しました。', '失敗', 10);
    ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message, ui.ButtonSet.OK);
  }
}
