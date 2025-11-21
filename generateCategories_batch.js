// ===================================================================
// カテゴリ生成処理: バッチ処理用の関数群
// ===================================================================

// 作業シート名
const GENERATE_CATEGORIES_WORK_LIST_SHEET_NAME = "_分類リスト生成作業リスト";
const MERGE_CATEGORIES_WORK_LIST_SHEET_NAME = "_分類付与作業リスト";
const REVISE_FEEDBACK_WORK_LIST_SHEET_NAME = "_形式知修正作業リスト";
const ILLUSTRATION_PROMPTS_WORK_LIST_SHEET_NAME = "_イラストプロンプト作業リスト";
const CREATE_IMAGES_WORK_LIST_SHEET_NAME = "_画像生成作業リスト";

/**
 * [SETUP] generateCategories のセットアップ
 * inputシートのデータを分割してタスクを作成します
 */
function generateCategories_SETUP() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('分類リスト生成のセットアップを開始します...', '開始', 10);

    // --- 1. 設定情報を取得 ---
    const direction = configSheet.getRange('C3').getValue();
    const prompt1 = promptSheet.getRange(prompt1_pos).getValue();

    if (!direction || !sep || isNaN(sep) || sep <= 0) {
      throw new Error('configシートのC3(方向), C4(分割数)のいずれかが無効です。');
    }

    // --- 2. 入力データを読み込む ---
    const inputSheetName = promptSheet.getRange(inputSheetName_pos).getValue();
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) {
      throw new Error(`データシート「${inputSheetName}」が見つかりません。`);
    }

    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];
    const data = allData.slice(1);

    if (data.length === 0) {
      ui.alert(`${inputSheetName}シートにデータがありません。`);
      return;
    }

    // --- 3. 作業シート作成 & タスク書き込み ---
    const workSheet = _createGenerateCategoriesWorkSheet(inputSheetName, prompt1, JSON.stringify(header));
    const workListData = [];

    // データをsep件ずつのチャンクに分割してタスク化
    for (let i = 0; i < data.length; i += sep) {
      const chunk = data.slice(i, Math.min(i + sep, data.length));
      workListData.push([
        `Chunk_${i}_${i + chunk.length - 1}`, // TaskKey
        JSON.stringify(chunk), // TaskData (チャンクデータをJSON形式)
        STATUS_EMPTY, // Status
        `${i + 1}-${i + chunk.length}` // 範囲（参照用）
      ]);
    }

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 4).setValues(workListData);
    }

    _showSetupCompletionDialog();

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] generateCategories バッチ処理ワーカー
 * これまでの分類結果を引き継ぎながら、順次処理します
 */
function generateCategories_PROCESS() {
  const startTime = new Date().getTime();

  const workSheet = ss.getSheetByName(GENERATE_CATEGORIES_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  // --- 1. 共通設定を作業シートから取得 ---
  const inputSheetName = workSheet.getRange("E1").getValue();
  const basePromptTemplate = workSheet.getRange("F1").getValue();
  const headerJson = workSheet.getRange("G1").getValue();

  // これまでの分類結果を取得（L1セルに保存）
  let previousResultJsonForPrompt = workSheet.getRange("L1").getValue() || "";

  if (!inputSheetName || !basePromptTemplate) {
    Logger.log("作業シート E1, F1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  const header = JSON.parse(headerJson);
  const basePrompt = _replacePrompts(basePromptTemplate);

  // --- 2. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 4);
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;
  let currentResult = previousResultJsonForPrompt ? JSON.parse(previousResultJsonForPrompt) : [];

  // --- 3. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][2]; // C列: Status

    if (currentStatus === STATUS_EMPTY) {
      // 実行時間が上限に近づいたら、自主的に終了
      const currentTime = new Date().getTime();
      if (currentTime - startTime > MAX_EXECUTION_TIME_MS) {
        Logger.log(`時間上限 (${MAX_EXECUTION_TIME_MS / 60000}分) に近づいたため、処理を中断します。`);
        // これまでの結果をL1セルに保存
        workSheet.getRange("L1").setValue(JSON.stringify(currentResult, null, 2));
        break;
      }

      const sheetRow = i + 2;
      const taskKey = workValues[i][0];
      const range = workValues[i][3];

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        // タスクデータを解析
        const chunk = JSON.parse(workValues[i][1]);

        Logger.log(`[${processedCountInThisRun + 1}] データ範囲 ${range} を分類中...`);

        // CSVに変換
        const chunkWithHeader = [header].concat(chunk);
        const csvChunk = chunkWithHeader.map(row =>
          row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
        ).join('\n');

        // プロンプトを構築
        let prompt = basePrompt;
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

        // APIを呼び出し
        const resultText = callGemini_(prompt);
        const jsonStringMatch = resultText.match(/```json\s*([\s\S]*?)\s*```/);
        const cleanedJsonString = jsonStringMatch ? jsonStringMatch[1] : resultText;
        currentResult = JSON.parse(cleanedJsonString);

        // 次のプロンプト用に更新
        previousResultJsonForPrompt = JSON.stringify(currentResult, null, 2);

        // 結果を作業シートに一時保存（E列以降）
        workSheet.getRange(sheetRow, 5).setValue(JSON.stringify(currentResult));

        // 待機（API制限対策）
        Utilities.sleep(1000);

        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;
        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);
      }
    }
  }

  Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);
  SpreadsheetApp.flush();

  // --- 4. 完了チェック ---
  const lastRow = workSheet.getLastRow();
  let remainingTasks = 0;
  if (lastRow >= 2) {
    const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    remainingTasks = newStatusValues.filter(
      row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
    ).length;
  }

  if (remainingTasks === 0) {
    Logger.log("✅ すべてのタスクが完了しました！");

    // 完了時に最終結果を新しいシートに出力
    _outputGenerateCategoriesResults(workSheet, currentResult);

    // L1セルの一時データをクリア
    workSheet.getRange("L1").clearContent();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'すべての分類リスト生成が完了し、結果を出力しました。',
      '✅ 完了',
      10
    );
  } else {
    // 未完了の場合、現在の結果をL1に保存
    workSheet.getRange("L1").setValue(JSON.stringify(currentResult, null, 2));

    Logger.log(`残りタスク数: ${remainingTasks}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `処理中... 残り ${remainingTasks} 件`,
      '分類リスト生成中',
      5
    );
  }
}

/**
 * [ヘルパー関数] generateCategories用の作業シートを作成
 */
function _createGenerateCategoriesWorkSheet(inputSheetName, prompt1, headerJson) {
  let workSheet = ss.getSheetByName(GENERATE_CATEGORIES_WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(GENERATE_CATEGORIES_WORK_LIST_SHEET_NAME, 0);
  }

  const workHeader = ["TaskKey", "TaskData", "Status", "Range", "Result"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1, F1, G1 に実行時に必要な情報を保存
  workSheet.getRange("E1").setValue(inputSheetName);
  workSheet.getRange("F1").setValue(prompt1);
  workSheet.getRange("G1").setValue(headerJson);

  // L1: これまでの分類結果を保存（継続実行用）
  workSheet.getRange("L1").setValue("");

  workSheet.autoResizeColumn(1);
  return workSheet;
}

/**
 * [ヘルパー関数] 完了時に分類結果を新しいシートに出力
 */
function _outputGenerateCategoriesResults(workSheet, result) {
  if (!result || result.length === 0) {
    Logger.log("出力する分類結果がありません。");
    return;
  }

  // 重複削除処理
  const uniqueCategoriesMap = new Map();
  result.forEach(item => {
    const key = `${item.major_category}_${item.minor_category}`;
    if (!uniqueCategoriesMap.has(key)) {
      uniqueCategoriesMap.set(key, item);
    }
  });

  const uniqueCategories = Array.from(uniqueCategoriesMap.values());

  // 新しいシートに出力
  const outputSheetName = `分類リスト_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);

  const outputHeader = Object.keys(uniqueCategories[0]);
  const outputData = uniqueCategories.map(item => outputHeader.map(key => item[key]));

  outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
  outputSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  outputSheet.autoResizeColumns(1, outputHeader.length);

  Logger.log(`シート「${outputSheetName}」に分類リストを出力しました。`);
}

/**
 * [SETUP] mergeCategories のセットアップ
 * 元データと分類リストを基に、分類付与タスクを作成します
 */
function mergeCategories_SETUP() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('分類付与のセットアップを開始します...', '開始', 10);

    // --- 1. 設定情報を取得 ---
    const inputSheetName = promptSheet.getRange(inputSheetName_pos).getValue();
    const categorySheetName = promptSheet.getRange(categorySheetName_pos).getValue();
    const prompt2 = promptSheet.getRange(prompt2_pos).getValue();

    const inputSheet = ss.getSheetByName(inputSheetName);
    const categorySheet = ss.getSheetByName(categorySheetName);

    if (!inputSheet) throw new Error(`入力シート「${inputSheetName}」が見つかりません。`);
    if (!categorySheet) throw new Error(`分類シート「${categorySheetName}」が見つかりません。`);

    // --- 2. 入力データを読み込む ---
    const allOriginalData = inputSheet.getDataRange().getValues();
    const originalHeader = allOriginalData[0];
    const originalData = allOriginalData.slice(1);

    if (originalData.length === 0) {
      ui.alert('入力シートにデータがありません。');
      return;
    }

    // 分類リストを読み込む
    const categoryData = categorySheet.getDataRange().getValues();
    categoryData.shift(); // ヘッダーを除外
    const categoryListAsJson = JSON.stringify(
      categoryData.map(row => ({ major_category: row[0], minor_category: row[1] })),
      null, 2
    );

    // --- 3. 作業シート作成 & タスク書き込み ---
    const workSheet = _createMergeCategoriesWorkSheet(inputSheetName, categorySheetName, prompt2, JSON.stringify(originalHeader), categoryListAsJson);
    const workListData = [];

    // データをsep件ずつのチャンクに分割してタスク化
    for (let i = 0; i < originalData.length; i += sep) {
      const chunk = originalData.slice(i, Math.min(i + sep, originalData.length));
      workListData.push([
        `Chunk_${i}_${i + chunk.length - 1}`, // TaskKey
        JSON.stringify(chunk), // TaskData
        STATUS_EMPTY, // Status
        `${i + 1}-${i + chunk.length}` // 範囲
      ]);
    }

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 4).setValues(workListData);
    }

    _showSetupCompletionDialog();

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] mergeCategories バッチ処理ワーカー
 */
function mergeCategories_PROCESS() {
  const startTime = new Date().getTime();

  const workSheet = ss.getSheetByName(MERGE_CATEGORIES_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  // --- 1. 共通設定を作業シートから取得 ---
  const inputSheetName = workSheet.getRange("E1").getValue();
  const categorySheetName = workSheet.getRange("F1").getValue();
  const basePromptTemplate = workSheet.getRange("G1").getValue();
  const headerJson = workSheet.getRange("H1").getValue();
  const categoryListAsJson = workSheet.getRange("I1").getValue();

  if (!inputSheetName || !basePromptTemplate) {
    Logger.log("作業シート E1, G1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  const header = JSON.parse(headerJson);
  const basePrompt = _replacePrompts(basePromptTemplate);

  // --- 2. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 4);
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;

  // --- 3. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][2]; // C列: Status

    if (currentStatus === STATUS_EMPTY) {
      // 実行時間が上限に近づいたら、自主的に終了
      const currentTime = new Date().getTime();
      if (currentTime - startTime > MAX_EXECUTION_TIME_MS) {
        Logger.log(`時間上限 (${MAX_EXECUTION_TIME_MS / 60000}分) に近づいたため、処理を中断します。`);
        break;
      }

      const sheetRow = i + 2;
      const taskKey = workValues[i][0];
      const range = workValues[i][3];

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        // タスクデータを解析
        const chunk = JSON.parse(workValues[i][1]);

        Logger.log(`[${processedCountInThisRun + 1}] データ範囲 ${range} に分類を付与中...`);

        // CSVに変換
        const csvChunk = [header].concat(chunk).map(row =>
          row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
        ).join('\n');

        // プロンプトを構築
        let prompt = basePrompt;
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
---`;

        // APIを呼び出し
        const resultText = callGemini_(prompt);
        const cleanedJsonString = resultText.match(/```json\s*([\s\S]*?)\s*```/)?.[1] || resultText;
        const newResults = JSON.parse(cleanedJsonString);

        // 結果を作業シートに保存（E列）
        workSheet.getRange(sheetRow, 5).setValue(JSON.stringify(newResults));

        // 待機（API制限対策）
        Utilities.sleep(1000);

        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;
        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);
      }
    }
  }

  Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);
  SpreadsheetApp.flush();

  // --- 4. 完了チェック ---
  const lastRow = workSheet.getLastRow();
  let remainingTasks = 0;
  if (lastRow >= 2) {
    const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    remainingTasks = newStatusValues.filter(
      row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
    ).length;
  }

  if (remainingTasks === 0) {
    Logger.log("✅ すべてのタスクが完了しました！");

    // 完了時に結果を新しいシートに出力
    _outputMergeCategoriesResults(workSheet, inputSheetName);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'すべての分類付与が完了し、結果を出力しました。',
      '✅ 完了',
      10
    );
  } else {
    Logger.log(`残りタスク数: ${remainingTasks}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `処理中... 残り ${remainingTasks} 件`,
      '分類付与中',
      5
    );
  }
}

/**
 * [ヘルパー関数] mergeCategories用の作業シートを作成
 */
function _createMergeCategoriesWorkSheet(inputSheetName, categorySheetName, prompt2, headerJson, categoryListAsJson) {
  let workSheet = ss.getSheetByName(MERGE_CATEGORIES_WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(MERGE_CATEGORIES_WORK_LIST_SHEET_NAME, 0);
  }

  const workHeader = ["TaskKey", "TaskData", "Status", "Range", "Result"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1〜I1 に実行時に必要な情報を保存
  workSheet.getRange("E1").setValue(inputSheetName);
  workSheet.getRange("F1").setValue(categorySheetName);
  workSheet.getRange("G1").setValue(prompt2);
  workSheet.getRange("H1").setValue(headerJson);
  workSheet.getRange("I1").setValue(categoryListAsJson);

  workSheet.autoResizeColumn(1);
  return workSheet;
}

/**
 * [ヘルパー関数] 完了時に分類付与結果を新しいシートに出力
 */
function _outputMergeCategoriesResults(workSheet, inputSheetName) {
  const lastRow = workSheet.getLastRow();
  if (lastRow < 2) return;

  // 結果データを読み込む（E列）
  const resultsRange = workSheet.getRange(2, 5, lastRow - 1, 1);
  const resultsData = resultsRange.getValues();
  const statusRange = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();

  // 完了したデータのみを結合
  let finalMergedData = [];
  for (let i = 0; i < resultsData.length; i++) {
    if (statusRange[i][0] === STATUS_DONE && resultsData[i][0]) {
      const chunkResults = JSON.parse(resultsData[i][0]);
      finalMergedData = finalMergedData.concat(chunkResults);
    }
  }

  if (finalMergedData.length === 0) {
    Logger.log("出力するデータがありません。");
    return;
  }

  // 新しいシートに出力
  const outputSheetName = `分類付与済_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);

  const finalHeader = Object.keys(finalMergedData[0]);
  const outputData = finalMergedData.map(item => finalHeader.map(key => item[key]));

  outputSheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]).setFontWeight('bold');
  outputSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  outputSheet.autoResizeColumns(1, finalHeader.length);

  Logger.log(`シート「${outputSheetName}」に分類付与済データを出力しました。`);
}

/**
 * [SETUP] reviseFeedback のセットアップ
 * 「形式知修正」シートの設定に基づいて、修正タスクを作成します
 */
function reviseFeedback_SETUP() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('形式知修正のセットアップを開始します...', '開始', 10);

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

    // --- 2. 作業シート作成 & タスク書き込み ---
    const workSheet = _createReviseFeedbackWorkSheet(feedbackSheetName, rawDataSheetName, feedbackRule);
    const workListData = [];

    revisionList.forEach((revision, index) => {
      const feedbackNumber = String(revision[0]);
      const revisionPrompt = revision[1];
      workListData.push([
        `Feedback_${feedbackNumber}`, // TaskKey
        JSON.stringify({ feedbackNumber, revisionPrompt }), // TaskData (JSON形式)
        STATUS_EMPTY, // Status
        feedbackNumber // 参照用
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 4).setValues(workListData);
    }

    _showSetupCompletionDialog();

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] reviseFeedback バッチ処理ワーカー
 * この関数を繰り返し実行して、タスクを順次処理します
 */
function reviseFeedback_PROCESS() {
  const startTime = new Date().getTime();

  const workSheet = ss.getSheetByName(REVISE_FEEDBACK_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  // --- 1. 共通設定を作業シートから取得 ---
  const feedbackSheetName = workSheet.getRange("E1").getValue();
  const rawDataSheetName = workSheet.getRange("F1").getValue();
  const feedbackRule = workSheet.getRange("G1").getValue();

  if (!feedbackSheetName || !rawDataSheetName || !feedbackRule) {
    Logger.log("作業シート E1, F1, G1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  // --- 2. 必要なデータを事前に読み込む ---
  let feedbackSheet, feedbackData, feedbackHeader, feedbackMap;
  let rawDataSheet, rawData, rawDataHeader, rawDataMap;

  try {
    feedbackSheet = ss.getSheetByName(feedbackSheetName);
    if (!feedbackSheet) throw new Error(`対象フィードバックシート「${feedbackSheetName}」が見つかりません。`);
    feedbackData = feedbackSheet.getDataRange().getValues();
    feedbackHeader = feedbackData.shift();
    feedbackMap = new Map(feedbackData.map(row => [String(row[0]), row]));

    rawDataSheet = ss.getSheetByName(rawDataSheetName);
    if (!rawDataSheet) throw new Error(`大元の入力シート「${rawDataSheetName}」が見つかりません。`);
    rawData = rawDataSheet.getDataRange().getValues();
    rawDataHeader = rawData.shift();
    rawDataMap = new Map(rawData.map(row => [String(row[0]), row]));
  } catch (e) {
    Logger.log(`必須リソースが開けません: ${e}`);
    return;
  }

  // --- 3. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 4);
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;
  let revisedFeedbackResults = [];

  // --- 4. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][2]; // C列: Status

    if (currentStatus === STATUS_EMPTY) {
      // 実行時間が上限に近づいたら、自主的に終了
      const currentTime = new Date().getTime();
      if (currentTime - startTime > MAX_EXECUTION_TIME_MS) {
        Logger.log(`時間上限 (${MAX_EXECUTION_TIME_MS / 60000}分) に近づいたため、処理を中断します。`);
        break;
      }

      const sheetRow = i + 2; // 作業シートの行番号
      const taskKey = workValues[i][0];
      const taskDataJson = workValues[i][1];
      const feedbackNumber = workValues[i][3];

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        // タスクデータを解析
        const taskData = JSON.parse(taskDataJson);
        const revisionPrompt = taskData.revisionPrompt;

        Logger.log(`[${processedCountInThisRun + 1}] フィードバック番号「${feedbackNumber}」を修正中...`);

        // Mapから元のフィードバックデータを取得
        const originalFeedbackRow = feedbackMap.get(feedbackNumber);
        if (!originalFeedbackRow) {
          throw new Error(`フィードバック番号「${feedbackNumber}」が見つかりませんでした。`);
        }

        const baseSerialNumbers = String(originalFeedbackRow[4]).split(/[\n,]/).map(s => s.trim());

        // 元の入力データをMapから取得
        let referencedRawData = "";
        baseSerialNumbers.forEach(serialNumber => {
          const rawRow = rawDataMap.get(serialNumber);
          if (rawRow) {
            referencedRawData += rawDataHeader.join(',') + '\n' + rawRow.join(',') + '\n\n';
          }
        });

        // --- AIへのプロンプトを構築 ---
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

        // --- APIを呼び出し、結果を格納 ---
        const resultText = callGemini_(finalPrompt);
        const cleanedJsonString = resultText.match(/```json\s*([\s\S]*?)\s*```/)?.[1] || resultText;
        const revisedFeedback = JSON.parse(cleanedJsonString);

        // 結果を作業シートのD列以降に書き込み（一時保存）
        const resultRow = [
          revisedFeedback["フィードバック番号"],
          revisedFeedback["フィードバックタイトル"],
          revisedFeedback["フィードバック概要"],
          revisedFeedback["フィードバック詳細"],
          revisedFeedback["ベース通し番号"],
          revisedFeedback["ベース概要（管理番号）"]
        ];
        workSheet.getRange(sheetRow, 5, 1, resultRow.length).setValues([resultRow]);

        // 待機（API制限対策）
        Utilities.sleep(1000);

        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;
        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);
      }
    }
  }

  Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);
  SpreadsheetApp.flush();

  // --- 5. 完了チェック ---
  const lastRow = workSheet.getLastRow();
  let remainingTasks = 0;
  if (lastRow >= 2) {
    const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    remainingTasks = newStatusValues.filter(
      row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
    ).length;
  }

  if (remainingTasks === 0) {
    Logger.log("✅ すべてのタスクが完了しました！");

    // 完了時に結果を新しいシートに出力
    _outputRevisedFeedbackResults(workSheet);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'すべての形式知修正が完了し、結果を出力しました。',
      '✅ 完了',
      10
    );
  } else {
    Logger.log(`残りタスク数: ${remainingTasks}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `処理中... 残り ${remainingTasks} 件`,
      '形式知修正中',
      5
    );
  }
}

/**
 * [ヘルパー関数] reviseFeedback用の作業シートを作成
 */
function _createReviseFeedbackWorkSheet(feedbackSheetName, rawDataSheetName, feedbackRule) {
  let workSheet = ss.getSheetByName(REVISE_FEEDBACK_WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(REVISE_FEEDBACK_WORK_LIST_SHEET_NAME, 0);
  }

  const workHeader = ["TaskKey", "TaskData", "Status", "FeedbackNumber", "結果_番号", "結果_タイトル", "結果_概要", "結果_詳細", "結果_ベース通し番号", "結果_ベース概要"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1, F1, G1 に実行時に必要な情報を保存
  workSheet.getRange("E1").setValue(feedbackSheetName);
  workSheet.getRange("F1").setValue(rawDataSheetName);
  workSheet.getRange("G1").setValue(feedbackRule);

  workSheet.autoResizeColumn(1);
  return workSheet;
}

/**
 * [ヘルパー関数] 完了時に結果を新しいシートに出力
 */
function _outputRevisedFeedbackResults(workSheet) {
  const lastRow = workSheet.getLastRow();
  if (lastRow < 2) return;

  // 結果データを読み込む（E列以降）
  const resultsRange = workSheet.getRange(2, 5, lastRow - 1, 6);
  const resultsData = resultsRange.getValues();

  // 完了したデータのみをフィルタリング
  const completedResults = [];
  const statusRange = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();

  for (let i = 0; i < resultsData.length; i++) {
    if (statusRange[i][0] === STATUS_DONE && resultsData[i][0]) {
      completedResults.push(resultsData[i]);
    }
  }

  if (completedResults.length === 0) return;

  // 新しいシートに出力
  const outputSheetName = `改訂版FB_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);

  const outputHeader = ["フィードバック番号", "フィードバックタイトル", "フィードバック概要", "フィードバック詳細", "ベース通し番号", "ベース概要（管理番号）"];

  outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
  outputSheet.getRange(2, 1, completedResults.length, completedResults[0].length)
    .setValues(completedResults)
    .setWrap(true)
    .setVerticalAlignment('top');

  outputSheet.autoResizeColumns(1, outputHeader.length);

  Logger.log(`シート「${outputSheetName}」に改訂版FBを出力しました。`);
}

/**
 * [SETUP] createIllustrationPrompts のセットアップ
 * 「設計フィードバック」シートの各行について、イラスト用プロンプト生成タスクを作成します
 */
function createIllustrationPrompts_SETUP() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('イラストプロンプト生成のセットアップを開始します...', '開始', 10);

    // --- 1. 設定情報を取得 ---
    const feedbackSheetName = promptSheet.getRange(feedbackSheetName_pos).getValue();
    const prompt4 = promptSheet.getRange(prompt4_pos).getValue();
    const columnsString = promptSheet.getRange('C10').getValue();

    if (!feedbackSheetName || !prompt4) {
      throw new Error('promptシートの設定（フィードバックシート名またはプロンプト）が不足しています。');
    }

    // --- 2. 入力データを読み込む ---
    const feedbackSheet = ss.getSheetByName(feedbackSheetName);
    if (!feedbackSheet) throw new Error(`対象フィードバックシート「${feedbackSheetName}」が見つかりません。`);

    const allData = feedbackSheet.getDataRange().getValues();
    const header = allData[0];
    const data = allData.slice(1);

    if (data.length === 0) {
      throw new Error(`入力シート「${feedbackSheetName}」にデータがありません。`);
    }

    // --- 3. 処理に必要な列のインデックスを特定 ---
    let columnIndices;
    if (columnsString) {
      columnIndices = _parseColumnRangeString(columnsString);
      if (columnIndices.length === 0) {
        throw new Error('promptシートC10セルの列指定が有効ではありませんでした。');
      }
    } else {
      columnIndices = header.map((_, index) => index);
    }

    const columnsToUse = columnIndices.map(index => {
      if (index < 0 || index >= header.length) {
        throw new Error(`列指定 ${index + 1} がシートの範囲外です。`);
      }
      return header[index];
    });

    // --- 4. 作業シート作成 & タスク書き込み ---
    const workSheet = _createIllustrationPromptsWorkSheet(feedbackSheetName, prompt4, JSON.stringify(columnIndices), JSON.stringify(columnsToUse));
    const workListData = [];

    data.forEach((row, index) => {
      const rowIndex = index + 2;
      workListData.push([
        `Row_${rowIndex}`, // TaskKey
        JSON.stringify(row), // TaskData (行データをJSON形式)
        STATUS_EMPTY, // Status
        rowIndex // 参照用
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 4).setValues(workListData);
    }

    _showSetupCompletionDialog();

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] createIllustrationPrompts バッチ処理ワーカー
 */
function createIllustrationPrompts_PROCESS() {
  const startTime = new Date().getTime();

  const workSheet = ss.getSheetByName(ILLUSTRATION_PROMPTS_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  // --- 1. 共通設定を作業シートから取得 ---
  const feedbackSheetName = workSheet.getRange("E1").getValue();
  const basePromptTemplate = workSheet.getRange("F1").getValue();
  const columnIndices = JSON.parse(workSheet.getRange("G1").getValue());
  const columnsToUse = JSON.parse(workSheet.getRange("H1").getValue());

  if (!feedbackSheetName || !basePromptTemplate) {
    Logger.log("作業シート E1, F1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  const basePrompt = _replacePrompts(basePromptTemplate);

  // --- 2. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 4);
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;

  // --- 3. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][2]; // C列: Status

    if (currentStatus === STATUS_EMPTY) {
      // 実行時間が上限に近づいたら、自主的に終了
      const currentTime = new Date().getTime();
      if (currentTime - startTime > MAX_EXECUTION_TIME_MS) {
        Logger.log(`時間上限 (${MAX_EXECUTION_TIME_MS / 60000}分) に近づいたため、処理を中断します。`);
        break;
      }

      const sheetRow = i + 2;
      const taskKey = workValues[i][0];
      const rowIndex = workValues[i][3];

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        // タスクデータを解析
        const row = JSON.parse(workValues[i][1]);

        Logger.log(`[${processedCountInThisRun + 1}] 行${rowIndex}のイラストプロンプトを生成中...`);

        // プロンプトに含めるフィードバック内容を構築
        let feedbackContent = "";
        columnsToUse.forEach((colName, idx) => {
          const dataIndex = columnIndices[idx];
          feedbackContent += `- ${colName}: ${row[dataIndex]}\n`;
        });

        const finalPrompt = basePrompt + feedbackContent;

        // APIを呼び出し
        const resultText = callGemini_(finalPrompt);
        const parsedTable = parseMarkdownTable_(resultText);

        let okCase = "（生成失敗）";
        let ngCase = "（生成失敗）";
        if (parsedTable.length > 1) {
          okCase = parsedTable[1][1] || okCase;
          ngCase = parsedTable[1][2] || ngCase;
        }

        // 結果を作業シートに書き込み
        workSheet.getRange(sheetRow, 5, 1, 2).setValues([[okCase, ngCase]]);

        // 待機（API制限対策）
        Utilities.sleep(1000);

        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;
        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);
      }
    }
  }

  Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);
  SpreadsheetApp.flush();

  // --- 4. 完了チェック ---
  const lastRow = workSheet.getLastRow();
  let remainingTasks = 0;
  if (lastRow >= 2) {
    const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    remainingTasks = newStatusValues.filter(
      row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
    ).length;
  }

  if (remainingTasks === 0) {
    Logger.log("✅ すべてのタスクが完了しました！");

    // 完了時に結果を新しいシートに出力
    _outputIllustrationPromptsResults(workSheet, feedbackSheetName);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'すべてのイラストプロンプト生成が完了し、結果を出力しました。',
      '✅ 完了',
      10
    );
  } else {
    Logger.log(`残りタスク数: ${remainingTasks}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `処理中... 残り ${remainingTasks} 件`,
      'イラストプロンプト生成中',
      5
    );
  }
}

/**
 * [ヘルパー関数] createIllustrationPrompts用の作業シートを作成
 */
function _createIllustrationPromptsWorkSheet(feedbackSheetName, prompt4, columnIndices, columnsToUse) {
  let workSheet = ss.getSheetByName(ILLUSTRATION_PROMPTS_WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(ILLUSTRATION_PROMPTS_WORK_LIST_SHEET_NAME, 0);
  }

  const workHeader = ["TaskKey", "TaskData", "Status", "RowIndex", "結果_OK事例", "結果_NG事例"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1, F1, G1, H1 に実行時に必要な情報を保存
  workSheet.getRange("E1").setValue(feedbackSheetName);
  workSheet.getRange("F1").setValue(prompt4);
  workSheet.getRange("G1").setValue(columnIndices);
  workSheet.getRange("H1").setValue(columnsToUse);

  workSheet.autoResizeColumn(1);
  return workSheet;
}

/**
 * [ヘルパー関数] 完了時にイラストプロンプト結果を新しいシートに出力
 */
function _outputIllustrationPromptsResults(workSheet, feedbackSheetName) {
  const lastRow = workSheet.getLastRow();
  if (lastRow < 2) return;

  // 元のフィードバックシートのデータを取得
  const feedbackSheet = ss.getSheetByName(feedbackSheetName);
  if (!feedbackSheet) return;

  const allData = feedbackSheet.getDataRange().getValues();
  const header = allData[0];
  const data = allData.slice(1);

  // 結果データを読み込む（E, F列）
  const resultsRange = workSheet.getRange(2, 5, lastRow - 1, 2);
  const resultsData = resultsRange.getValues();

  // 完了したデータのみをマージ
  const outputRows = [];
  const statusRange = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();

  for (let i = 0; i < data.length && i < resultsData.length; i++) {
    if (statusRange[i][0] === STATUS_DONE) {
      outputRows.push(data[i].concat(resultsData[i]));
    }
  }

  if (outputRows.length === 0) return;

  // 新しいシートに出力
  const outputSheetName = `イラストプロンプト案_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);

  const outputHeader = header.concat(['OK事例', 'NG事例']);

  outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
  outputSheet.getRange(2, 1, outputRows.length, outputRows[0].length)
    .setValues(outputRows)
    .setWrap(true)
    .setVerticalAlignment('top');

  outputSheet.autoResizeColumns(1, outputHeader.length);

  Logger.log(`シート「${outputSheetName}」にイラスト用プロンプト案を出力しました。`);
}

/**
 * [SETUP] createImages のセットアップ
 * 「イラストプロンプト案」シートの設定に基づいて、画像生成タスクを作成します
 */
function createImages_SETUP() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('画像生成のセットアップを開始します...', '開始', 10);

    // --- 1. 設定情報を取得 ---
    const imagePromptSheetName = promptSheet.getRange(imagePromptSheetName_pos).getValue();
    const promt5 = promptSheet.getRange(prompt5_pos).getValue();
    const outputFolderUrl = promptSheet.getRange(imageSaveDir_pos).getValue();

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

    // --- 2b. ヘッダー列を準備 ---
    const existingImageCols = header.filter(h => h.toString().startsWith('生成画像'));
    const firstNewColIndex = header.length;
    let newHeaders = [];

    for (let i = 0; i < numberOfGenerations; i++) {
      const colNumber = existingImageCols.length / 2 + i + 1;
      const imageHeaderName = colNumber === 1 ? '生成画像' : `生成画像_${colNumber}`;
      newHeaders.push(imageHeaderName);
    }

    if (newHeaders.length > 0) {
      sheet.getRange(1, firstNewColIndex + 1, 1, newHeaders.length).setValues([newHeaders]).setFontWeight('bold');
      header = header.concat(newHeaders);
    }

    const okCaseIndex = header.indexOf('OK事例');
    const ngCaseIndex = header.indexOf('NG事例');
    if (okCaseIndex === -1 || ngCaseIndex === -1) {
      throw new Error('入力シートに「OK事例」または「NG事例」の列が見つかりません。');
    }

    // --- 3. 作業シート作成 & タスク書き込み ---
    const workSheet = _createImagesWorkSheet(
      imagePromptSheetName,
      promt5,
      outputFolderUrl,
      numberOfGenerations,
      okCaseIndex,
      ngCaseIndex,
      firstNewColIndex
    );
    const workListData = [];

    dataToProcess.forEach(item => {
      workListData.push([
        `Row_${item.rowIndex}`, // TaskKey
        JSON.stringify(item.rowData), // TaskData (行データをJSON形式)
        STATUS_EMPTY, // Status
        item.serialNumber // 参照用
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 4).setValues(workListData);
    }

    _showSetupCompletionDialog();

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] createImages バッチ処理ワーカー
 */
function createImages_PROCESS() {
  const startTime = new Date().getTime();

  const workSheet = ss.getSheetByName(CREATE_IMAGES_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  // --- 1. 共通設定を作業シートから取得 ---
  const imagePromptSheetName = workSheet.getRange("E1").getValue();
  const basePromptTemplate = workSheet.getRange("F1").getValue();
  const outputFolderUrl = workSheet.getRange("G1").getValue();
  const numberOfGenerations = parseInt(workSheet.getRange("H1").getValue(), 10);
  const okCaseIndex = parseInt(workSheet.getRange("I1").getValue(), 10);
  const ngCaseIndex = parseInt(workSheet.getRange("J1").getValue(), 10);
  const firstNewColIndex = parseInt(workSheet.getRange("K1").getValue(), 10);

  if (!imagePromptSheetName || !basePromptTemplate || !outputFolderUrl) {
    Logger.log("作業シート E1, F1, G1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  const basePrompt = _replacePrompts(basePromptTemplate);

  // --- 2. 必要なリソースを取得 ---
  let sheet, outputFolder;

  try {
    sheet = ss.getSheetByName(imagePromptSheetName);
    if (!sheet) throw new Error(`シート「${imagePromptSheetName}」が見つかりません。`);

    const folderId = _extractFolderIdFromUrl(outputFolderUrl);
    if (!folderId) throw new Error('フォルダIDを取得できませんでした。');
    outputFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log(`必須リソースが開けません: ${e}`);
    return;
  }

  // --- 3. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 4);
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;

  // --- 4. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][2]; // C列: Status

    if (currentStatus === STATUS_EMPTY) {
      // 実行時間が上限に近づいたら、自主的に終了
      const currentTime = new Date().getTime();
      if (currentTime - startTime > MAX_EXECUTION_TIME_MS) {
        Logger.log(`時間上限 (${MAX_EXECUTION_TIME_MS / 60000}分) に近づいたため、処理を中断します。`);
        break;
      }

      const sheetRow = i + 2;
      const taskKey = workValues[i][0];
      const serialNumber = workValues[i][3];

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        // タスクデータを解析
        const rowData = JSON.parse(workValues[i][1]);
        const rowIndex = parseInt(taskKey.split('_')[1], 10);

        const okCase = rowData[okCaseIndex];
        const ngCase = rowData[ngCaseIndex];

        let finalPrompt = basePrompt
          .replace('<NG_Image>', ngCase)
          .replace('<OK_Image>', okCase);

        Logger.log(`[${processedCountInThisRun + 1}] No.${serialNumber} の画像生成中 (${numberOfGenerations}枚)...`);

        // 指定された回数だけAPIを呼び出し、画像を生成
        for (let j = 0; j < numberOfGenerations; j++) {
          const currentImageColIndex = firstNewColIndex + j;

          const base64Image = callGPTApi_(finalPrompt);

          // (1) Driveに保存
          const colNumber = j + 1;
          const imageHeaderName = colNumber === 1 ? '生成画像' : `生成画像_${colNumber}`;
          const imageName = `${imagePromptSheetName}_No${serialNumber}_${imageHeaderName}.png`;
          let savedFileUrl = '';

          try {
            const decodedBytes = Utilities.base64Decode(base64Image);
            const imageBlob = Utilities.newBlob(decodedBytes, MimeType.PNG, imageName);
            const savedFile = outputFolder.createFile(imageBlob);
            savedFileUrl = savedFile.getUrl();
            Logger.log(`画像を保存: ${savedFile.getName()}`);
          } catch (saveError) {
            Logger.log(`警告: No.${serialNumber} の画像 ${colNumber} の保存に失敗 - ${saveError}`);
            savedFileUrl = '保存失敗';
          }

          // (2) シートに画像を挿入
          const dataUrl = `data:image/png;base64,${base64Image}`;
          const cellImage = SpreadsheetApp.newCellImage().setSourceUrl(dataUrl).build();
          sheet.getRange(rowIndex, currentImageColIndex + 1).setValue(cellImage);

          if (j < numberOfGenerations - 1) {
            Utilities.sleep(1000);
          }
        }

        sheet.setRowHeight(rowIndex, 200);

        // 待機（API制限対策）
        Utilities.sleep(1000);

        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;
        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);
      }
    }
  }

  Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);
  SpreadsheetApp.flush();

  // --- 5. 完了チェック ---
  const lastRow = workSheet.getLastRow();
  let remainingTasks = 0;
  if (lastRow >= 2) {
    const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    remainingTasks = newStatusValues.filter(
      row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
    ).length;
  }

  if (remainingTasks === 0) {
    Logger.log("✅ すべてのタスクが完了しました！");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'すべての画像生成が完了しました。',
      '✅ 完了',
      10
    );
  } else {
    Logger.log(`残りタスク数: ${remainingTasks}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `処理中... 残り ${remainingTasks} 件`,
      '画像生成中',
      5
    );
  }
}

/**
 * [ヘルパー関数] createImages用の作業シートを作成
 */
function _createImagesWorkSheet(imagePromptSheetName, promt5, outputFolderUrl, numberOfGenerations, okCaseIndex, ngCaseIndex, firstNewColIndex) {
  let workSheet = ss.getSheetByName(CREATE_IMAGES_WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(CREATE_IMAGES_WORK_LIST_SHEET_NAME, 0);
  }

  const workHeader = ["TaskKey", "TaskData", "Status", "SerialNumber"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1〜K1 に実行時に必要な情報を保存
  workSheet.getRange("E1").setValue(imagePromptSheetName);
  workSheet.getRange("F1").setValue(promt5);
  workSheet.getRange("G1").setValue(outputFolderUrl);
  workSheet.getRange("H1").setValue(numberOfGenerations);
  workSheet.getRange("I1").setValue(okCaseIndex);
  workSheet.getRange("J1").setValue(ngCaseIndex);
  workSheet.getRange("K1").setValue(firstNewColIndex);

  workSheet.autoResizeColumn(1);
  return workSheet;
}

// ===================================================================
// 注: 以下の共通ヘルパー関数は commonHelpers.js に移動しました
// - _showSetupCompletionDialog()
// - _parseColumnRangeString()
// - _parseNumberRangeString()
// - _extractFolderIdFromUrl()
// - _replacePrompts()
// - parseMarkdownTable_()
// ===================================================================
