// ===================================================================
// カテゴリ生成処理: バッチ処理用の関数群
// ===================================================================

// 作業シート名
const GENERATE_CATEGORIES_WORK_LIST_SHEET_NAME = "_分類リスト生成作業リスト";
const MERGE_CATEGORIES_WORK_LIST_SHEET_NAME = "_分類付与作業リスト";
const GENERATE_FEEDBACK_WORK_LIST_SHEET_NAME = "_設計FB生成作業リスト";
const GENERATE_FEEDBACK_TEMP_RESULTS_SHEET_NAME = "_設計FB中間結果"; // 50,000文字制限回避用
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

    _showSetupCompletionDialog({
      workSheetName: GENERATE_CATEGORIES_WORK_LIST_SHEET_NAME,
      menuItemName: '📊 分類・整理 > ①-2 分類リストを生成 (実行)',
      processFunctionName: 'generateCategories_PROCESS',
      useManualExecution: true
    });

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
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

  const workSheet = ss.getSheetByName(GENERATE_CATEGORIES_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  _showProgress('分類リスト生成処理を開始します...', '📊 分類生成', 3);

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
      // 動的タイムアウトチェック：次のタスクを実行可能かを判定
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        // これまでの結果をL1セルに保存
        workSheet.getRange("L1").setValue(JSON.stringify(currentResult, null, 2));
        break;
      }

      const taskStartTime = new Date().getTime(); // このタスクの開始時刻
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

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        // 進捗表示（手動実行時のみ）
        if (processedCountInThisRun % 3 === 0) {
          const totalTasks = workValues.length;
          _showProgress(
            `${processedCountInThisRun} / ${totalTasks} 件完了`,
            '📊 分類生成中',
            2
          );
        }

        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

        // エラーの場合も実行時間を記録（次回の予測精度向上のため）
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
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

    _showProgress(
      'すべての分類リスト生成が完了し、結果を出力しました。',
      '✅ 完了',
      10
    );
  } else {
    // 未完了の場合、現在の結果をL1に保存
    workSheet.getRange("L1").setValue(JSON.stringify(currentResult, null, 2));

    Logger.log(`残りタスク数: ${remainingTasks}`);
    _showProgress(
      `今回 ${processedCountInThisRun} 件処理。残り ${remainingTasks} 件`,
      '⏸️ 一時停止',
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

  // タブの色をグレーに設定
  workSheet.setTabColor('#999999');

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

    _showSetupCompletionDialog({
      workSheetName: MERGE_CATEGORIES_WORK_LIST_SHEET_NAME,
      menuItemName: '📊 分類・整理 > ②-2 データに分類を付与 (実行)',
      processFunctionName: 'mergeCategories_PROCESS',
      useManualExecution: true
    });

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
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

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
      // 動的タイムアウトチェック：次のタスクを実行可能かを判定
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        break;
      }

      const taskStartTime = new Date().getTime();
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

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

        // エラーの場合も実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
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
    workSheet = ss.insertSheet(MERGE_CATEGORIES_WORK_LIST_SHEET_NAME, ss.getNumSheets() + 1);
  }

  const workHeader = ["TaskKey", "TaskData", "Status", "Range", "Result"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1〜I1 に実行時に必要な情報を保存
  workSheet.getRange("E1").setValue(inputSheetName);
  workSheet.getRange("F1").setValue(categorySheetName);
  workSheet.getRange("G1").setValue(prompt2);
  workSheet.getRange("H1").setValue(headerJson);
  workSheet.getRange("I1").setValue(categoryListAsJson);

  // タブの色をグレーに設定
  workSheet.setTabColor('#999999');

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
 * [SETUP] generateFeedback のセットアップ
 * カテゴリごとにグループ化してタスクを作成します
 */
function generateFeedback_SETUP() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('設計FB生成のセットアップを開始します...', '開始', 10);

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

    // --- 4. 作業シート作成 & タスク書き込み ---
    // TaskData列を削除し、カテゴリ名のみ保存（50,000文字制限回避）
    const workSheet = _createGenerateFeedbackWorkSheet(inputSheetName, basePrompt, JSON.stringify(header), inputCategory);
    const workListData = [];

    const categories = Object.keys(groupedData);
    categories.forEach((categoryName, index) => {
      workListData.push([
        `Category_${index}_${categoryName}`, // TaskKey
        STATUS_EMPTY, // Status
        categoryName // カテゴリ名（PROCESS時に入力シートから該当データを抽出）
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 3).setValues(workListData);
    }

    // --- 5. 中間結果シートを作成（50,000文字制限回避用：複数行構造）---
    let tempResultsSheet = ss.getSheetByName(GENERATE_FEEDBACK_TEMP_RESULTS_SHEET_NAME);
    if (tempResultsSheet) {
      tempResultsSheet.clear();
    } else {
      tempResultsSheet = ss.insertSheet(GENERATE_FEEDBACK_TEMP_RESULTS_SHEET_NAME, 0);
    }

    // ヘッダーを設定（複数行形式）
    const tempHeader = ["カテゴリ名", "バッチ番号", "フィードバック内容", "処理済み"];
    tempResultsSheet.getRange(1, 1, 1, 4).setValues([tempHeader]).setFontWeight('bold');
    tempResultsSheet.setTabColor('#cccccc'); // グレー
    tempResultsSheet.setColumnWidth(3, 500); // フィードバック内容列を広く

    Logger.log(`中間結果シート「${GENERATE_FEEDBACK_TEMP_RESULTS_SHEET_NAME}」を作成しました（複数行構造）。`);

    _showSetupCompletionDialog({
      workSheetName: GENERATE_FEEDBACK_WORK_LIST_SHEET_NAME,
      menuItemName: '📝 設計FB > ③-2 設計FBを生成 (実行)',
      processFunctionName: 'generateFeedback_PROCESS',
      useManualExecution: true,
      tempResultsSheetName: GENERATE_FEEDBACK_TEMP_RESULTS_SHEET_NAME
    });

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] generateFeedback バッチ処理ワーカー
 * カテゴリごとに処理し、前回までのフィードバック結果を引き継ぎます
 */
function generateFeedback_PROCESS() {
  const startTime = new Date().getTime();
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

  const workSheet = ss.getSheetByName(GENERATE_FEEDBACK_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  // 中間結果シートを取得
  const tempResultsSheet = ss.getSheetByName(GENERATE_FEEDBACK_TEMP_RESULTS_SHEET_NAME);
  if (!tempResultsSheet) {
    Logger.log("中間結果シートが見つかりません。SETUPを先に実行してください。");
    return;
  }

  // --- 1. 共通設定を作業シートから取得 ---
  const inputSheetName = workSheet.getRange("D1").getValue();
  const basePromptTemplate = workSheet.getRange("E1").getValue();
  const headerJson = workSheet.getRange("F1").getValue();
  const inputCategoryColumn = workSheet.getRange("G1").getValue();

  // これまでのフィードバック結果を中間結果シートから取得
  let previousFeedbackForPrompt = _loadPreviousFeedbackFromTempSheet(tempResultsSheet);

  if (!inputSheetName || !basePromptTemplate) {
    Logger.log("作業シート D1, E1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  const header = JSON.parse(headerJson);
  const basePrompt = _replacePrompts(basePromptTemplate);

  // --- 1.5. 入力シートからデータを読み込む（都度読み込み方式）---
  const inputSheet = ss.getSheetByName(inputSheetName);
  if (!inputSheet) {
    Logger.log(`入力シート「${inputSheetName}」が見つかりません。`);
    return;
  }
  const allInputData = inputSheet.getDataRange().getValues();
  const inputHeader = allInputData[0];
  const inputData = allInputData.slice(1);
  const categoryColumnIndex = inputHeader.indexOf(inputCategoryColumn);

  if (categoryColumnIndex === -1) {
    Logger.log(`入力シートのヘッダーに「${inputCategoryColumn}」列が見つかりません。`);
    return;
  }

  // --- 2. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 3);
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;
  let combinedMarkdownResponse = previousFeedbackForPrompt;

  // --- 3. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][1]; // B列: Status（列が変わった）

    if (currentStatus === STATUS_EMPTY) {
      // 動的タイムアウトチェック：次のタスクを実行可能かを判定
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        // 中間結果は既にtempResultsSheetに保存済みなので、ここでは何もしない
        break;
      }

      const taskStartTime = new Date().getTime();
      const sheetRow = i + 2;
      const taskKey = workValues[i][0];
      const categoryName = workValues[i][2]; // C列: カテゴリ名（列が変わった）

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 2).setValue(STATUS_PROCESSING); // B列に変更

        // 入力シートから該当カテゴリのデータを抽出（都度読み込み）
        const chunk = inputData.filter(row => row[categoryColumnIndex] === categoryName);

        if (chunk.length === 0) {
          Logger.log(`カテゴリ「${categoryName}」のデータが見つかりません。スキップします。`);
          workSheet.getRange(sheetRow, 2).setValue(STATUS_DONE);
          continue;
        }

        Logger.log(`[${processedCountInThisRun + 1}] カテゴリ「${categoryName}」を分析中... (${chunk.length}行)`);

        // CSVに変換
        const csvChunk = [header].concat(chunk).map(row =>
          row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
        ).join('\n');

        // カテゴリ内で複数回API呼び出しを行う可能性がある
        let continueProcessingCategory = true;
        let batchNumber = 1;

        while (continueProcessingCategory) {
          // 時間チェック（whileループ内も動的チェック）
          if (!_shouldContinueProcessing(startTime, taskExecutionTimes, 2.0)) {
            Logger.log(`時間上限に近づいたため、カテゴリ「${categoryName}」の処理を中断します。`);
            // 中間結果は既にtempResultsSheetに保存済み
            throw new Error("時間制限により中断");
          }

          let prompt = basePrompt;
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
          combinedMarkdownResponse += resultText + "\n";
          previousFeedbackForPrompt += resultText + "\n";

          // 🔥 各API呼び出しの直後に中間結果を保存（バッチごとに行を追加）
          _saveCategoryResultToTempSheet(tempResultsSheet, categoryName, batchNumber, resultText);
          Logger.log(`  バッチ ${batchNumber} の結果を中間シートに保存しました`);

          batchNumber++;

          const newFeedbackData = parseMarkdownTable_(resultText);
          if (newFeedbackData.length <= 1 || resultText.includes('続きなし')) {
            continueProcessingCategory = false;
          }

          Utilities.sleep(1000);
        }

        // whileループ完了 = カテゴリの処理が正常終了
        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 2).setValue(STATUS_DONE); // B列に変更
        processedCountInThisRun++;

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 2).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`); // B列に変更

        // エラーの場合も実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);

        // エラー発生時は処理を中断（中間結果は既にtempResultsSheetに保存済み）
        break;
      }
    }
  }

  Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);
  SpreadsheetApp.flush();

  // --- 4. 完了チェック ---
  const lastRowForCheck = workSheet.getLastRow();
  let remainingTasks = 0;
  if (lastRowForCheck >= 2) {
    const newStatusValues = workSheet.getRange(2, 2, lastRowForCheck - 1, 1).getValues(); // B列に変更
    remainingTasks = newStatusValues.filter(
      row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
    ).length;
  }

  if (remainingTasks === 0) {
    Logger.log("✅ すべてのタスクが完了しました！");

    // 完了時に中間結果シートから全データを読み込んで最終結果を出力
    const allResults = _loadAllResultsFromTempSheet(tempResultsSheet);
    _outputGenerateFeedbackResults(workSheet, allResults);

    // 中間結果シートを削除（または保持する場合はコメントアウト）
    ss.deleteSheet(tempResultsSheet);
    Logger.log("中間結果シートを削除しました。");

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'すべての設計FB生成が完了し、結果を出力しました。',
      '✅ 完了',
      10
    );
  } else {
    Logger.log(`残りタスク数: ${remainingTasks}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `処理中... 残り ${remainingTasks} 件`,
      '設計FB生成中',
      5
    );
  }
}

/**
 * [ヘルパー関数] generateFeedback用の作業シートを作成
 */
function _createGenerateFeedbackWorkSheet(inputSheetName, prompt3, headerJson, inputCategoryColumn) {
  let workSheet = ss.getSheetByName(GENERATE_FEEDBACK_WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(GENERATE_FEEDBACK_WORK_LIST_SHEET_NAME, 0);
  }

  // TaskData列を削除し、カテゴリ名のみ保存（50,000文字制限回避）
  const workHeader = ["TaskKey", "Status", "Category"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // D1, E1, F1, G1 に実行時に必要な情報を保存
  workSheet.getRange("D1").setValue(inputSheetName);
  workSheet.getRange("E1").setValue(prompt3);
  workSheet.getRange("F1").setValue(headerJson);
  workSheet.getRange("G1").setValue(inputCategoryColumn); // カテゴリ列名を保存

  // タブの色をグレーに設定
  workSheet.setTabColor('#999999');

  workSheet.autoResizeColumn(1);
  return workSheet;
}

/**
 * [ヘルパー関数] 完了時に設計FB結果を新しいシートに出力
 */
function _outputGenerateFeedbackResults(workSheet, combinedMarkdownResponse) {
  if (!combinedMarkdownResponse) {
    Logger.log("出力する設計FB結果がありません。");
    return;
  }

  // Markdownテーブルをパース
  const feedbackData = parseMarkdownTable_(combinedMarkdownResponse);

  if (feedbackData.length === 0) {
    Logger.log("Markdownテーブルのパースに失敗しました。");
    return;
  }

  // 重複したヘッダー行を削除
  const headerRow = feedbackData[0];
  const headerString = headerRow.join('|');
  const uniqueHeaderData = feedbackData.filter((row, index) => {
    return index === 0 || row.join('|') !== headerString;
  });

  // 新しいシートに出力
  const outputSheetName = `設計FB_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
  const outputSheet = ss.insertSheet(outputSheetName, ss.getNumSheets() + 1);

  outputSheet.getRange(1, 1, uniqueHeaderData.length, uniqueHeaderData[0].length)
    .setValues(uniqueHeaderData)
    .setWrap(true)
    .setVerticalAlignment('top');

  outputSheet.getRange(1, 1, 1, uniqueHeaderData[0].length).setFontWeight('bold');
  outputSheet.autoResizeColumns(1, uniqueHeaderData[0].length);

  Logger.log(`シート「${outputSheetName}」に設計FBを出力しました。`);
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

    _showSetupCompletionDialog({
      workSheetName: REVISE_FEEDBACK_WORK_LIST_SHEET_NAME,
      menuItemName: '📝 設計FB > ④-2 FBを個別に修正 (実行)',
      processFunctionName: 'reviseFeedback_PROCESS',
      useManualExecution: true
    });

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
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

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
      // 動的タイムアウトチェック：次のタスクを実行可能かを判定
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        break;
      }

      const taskStartTime = new Date().getTime();
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

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

        // エラーの場合も実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
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

  // タブの色をグレーに設定
  workSheet.setTabColor('#999999');

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

    _showSetupCompletionDialog({
      workSheetName: ILLUSTRATION_PROMPTS_WORK_LIST_SHEET_NAME,
      menuItemName: '🎨 イラスト生成 > ⑤-2 イラスト用プロンプト案を生成 (実行)',
      processFunctionName: 'createIllustrationPrompts_PROCESS',
      useManualExecution: true
    });

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
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

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
      // 動的タイムアウトチェック：次のタスクを実行可能かを判定
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        break;
      }

      const taskStartTime = new Date().getTime();
      const sheetRow = i + 2;
      const taskKey = workValues[i][0];
      const rowIndex = workValues[i][3];

      try{
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

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

        // エラーの場合も実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
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

  // タブの色をグレーに設定
  workSheet.setTabColor('#999999');

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
    // 「生成画像」で始まり「URL」を含まない列のみカウント（生成画像, 生成画像_2, 生成画像_3...）
    const existingImageCols = header.filter(h => h.toString().startsWith('生成画像') && !h.toString().includes('URL'));
    const firstNewColIndex = header.length;
    let newHeaders = [];

    for (let i = 0; i < numberOfGenerations; i++) {
      const colNumber = existingImageCols.length + i + 1; // 既存の画像列数 + 新規インデックス
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

    _showSetupCompletionDialog({
      workSheetName: CREATE_IMAGES_WORK_LIST_SHEET_NAME,
      menuItemName: '🎨 イラスト生成 > ⑥-2 イラストを一括生成 (実行)',
      processFunctionName: 'createImages_PROCESS',
      useManualExecution: true
    });

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
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

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
      // 動的タイムアウトチェック：次のタスクを実行可能かを判定
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        break;
      }

      const taskStartTime = new Date().getTime();
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

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

        // エラーの場合も実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
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

  // タブの色をグレーに設定
  workSheet.setTabColor('#999999');

  workSheet.autoResizeColumn(1);
  return workSheet;
}

// ===================================================================
// 設計FB生成用ヘルパー関数（50,000文字制限対応）
// ===================================================================

/**
 * 中間結果シートからこれまでのフィードバック結果を読み込む（複数行形式対応）
 * @param {Sheet} tempResultsSheet - 中間結果シート
 * @return {string} - 前回までのフィードバック結果（Markdown形式）
 */
function _loadPreviousFeedbackFromTempSheet(tempResultsSheet) {
  const lastRow = tempResultsSheet.getLastRow();
  if (lastRow < 2) {
    return ""; // ヘッダーのみの場合は空
  }

  const data = tempResultsSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const processedResults = data.filter(row => row[3] === true); // 処理済みのみ（D列）

  if (processedResults.length === 0) {
    return "";
  }

  // フィードバック内容を結合（C列：フィードバック内容）
  return processedResults.map(row => row[2]).join('\n\n');
}

/**
 * カテゴリの処理結果を中間結果シートに保存（複数行形式）
 * @param {Sheet} tempResultsSheet - 中間結果シート
 * @param {string} categoryName - カテゴリ名
 * @param {number} batchNumber - バッチ番号
 * @param {string} markdown - フィードバック内容（Markdown形式、このバッチ分のみ）
 */
function _saveCategoryResultToTempSheet(tempResultsSheet, categoryName, batchNumber, markdown) {
  const lastRow = tempResultsSheet.getLastRow();

  // 同じカテゴリ・バッチ番号の既存行を検索
  let targetRow = -1;
  if (lastRow >= 2) {
    const data = tempResultsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === categoryName && data[i][1] === batchNumber) {
        targetRow = i + 2; // 実際のシート行番号
        break;
      }
    }
  }

  if (targetRow !== -1) {
    // 既存のバッチを更新（通常は発生しないが念のため）
    tempResultsSheet.getRange(targetRow, 3).setValue(markdown);
    tempResultsSheet.getRange(targetRow, 4).setValue(true);
    Logger.log(`カテゴリ「${categoryName}」バッチ ${batchNumber} の結果を更新しました（行${targetRow}）`);
  } else {
    // 新しいバッチを追加
    tempResultsSheet.appendRow([categoryName, batchNumber, markdown, true]);
    Logger.log(`カテゴリ「${categoryName}」バッチ ${batchNumber} の結果を追加しました`);
  }
}

/**
 * 中間結果シートから全結果を読み込む（複数行形式対応）
 * @param {Sheet} tempResultsSheet - 中間結果シート
 * @return {string} - 全フィードバック結果（Markdown形式）
 */
function _loadAllResultsFromTempSheet(tempResultsSheet) {
  const lastRow = tempResultsSheet.getLastRow();
  if (lastRow < 2) {
    return "";
  }

  const data = tempResultsSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // カテゴリ名でグループ化してソート、バッチ番号順に結合
  const categoryMap = {};
  data.forEach(row => {
    const categoryName = row[0];
    const batchNumber = row[1];
    const feedback = row[2];

    if (!categoryMap[categoryName]) {
      categoryMap[categoryName] = [];
    }
    categoryMap[categoryName].push({ batchNumber, feedback });
  });

  // 各カテゴリ内でバッチ番号順にソート
  const result = [];
  Object.keys(categoryMap).forEach(categoryName => {
    const batches = categoryMap[categoryName];
    batches.sort((a, b) => a.batchNumber - b.batchNumber);
    const categoryFeedback = batches.map(b => b.feedback).join('\n\n');
    result.push(categoryFeedback);
  });

  return result.join('\n\n');
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
