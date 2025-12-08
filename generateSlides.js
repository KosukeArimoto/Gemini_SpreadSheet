
// ===================================================================
// STEP 1: SETUP関数
// ===================================================================

/**
 * [SETUP] 1行1スライド (DetailTR) のセットアップ
 */
function createSlideDetailTR_SETUP() {
  const ui = SpreadsheetApp.getUi();
  try {
    ss.toast('セットアップ (DetailTR) を開始します...', '開始', 10);
    // --- 元の設定項目 ---
    const SLIDES_TEMPLATE_ID_TR = '1NYkmHwG4hHm8sadB_n15N6knXNGXtX3ZpLibePXfKS8';
    const TEMPLATE_SLIDE_INDEX_TR = 1;
    const ALT_TEXT_TITLE_MAP_TR = {
      "placeholder_equip":0, "placeholder_line":1, "placeholder_process":2,
      "placeholder_title":3, "placeholder_point":4, "placeholder_detail":5,
      "placeholder_check":6, "placeholder_id":7, "placeholder_place":8,
      "placeholder_point_rough":9, "placeholder_equip_num":11,
      "placeholder_original_num":12,
    };
    const IMAGE_ALT_TEXT_TITLE_TR = 'placeholder_image'; // 画像プレースホルダーのタイトル
    const ILLUSTRATION_COLUMN_INDEX_TR = 13; // N列（0-indexed）
    const combineRows = false;
    const mode = 'DetailTR';
    const groupingColumns = ["設備名称", "工程", "異常現象"];

    // --- 1. 対象シート取得 (元のロジック) ---
    const targetSheetName = tokaiPromptSheet.getRange("C12").getValue();
    if (!targetSheetName) throw new Error(`対象シート名が入力されていません。`);
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`データシート "${targetSheetName}" が見つかりません。`);

        // --- 2. ID採番 ---
    try {
      const masterSheetName = tokaiPromptSheet.getRange("C14").getValue();
      const id_col=8;
      const ID_PREFIX="DC-TY-";
      assignPersistentGroupIds_(sheet, masterSheetName, id_col, ID_PREFIX, groupingColumns); 
      SpreadsheetApp.getActiveSpreadsheet().toast('グループIDをA列に採番・更新しました。', 'ID採番完了', 3);
    } catch (e) {
      throw new Error(`ID採番中にエラーが発生しました: ${e.message}`);
    }

    // --- 2. 新規プレゼンテーション作成 ---
    const newPresentationTitle = `詳細事例スライド_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
    const presentationId = _createAndMovePresentation(newPresentationTitle);

    // --- 3. データ行取得 ---
    const allData = sheet.getDataRange().getValues();
    const dataRows = allData.slice(1);
    if (dataRows.length === 0) throw new Error('シートにデータが見つかりません（ヘッダーを除く）。');

    // --- 4. 作業シート作成 & タスク書き込み ---
    const workSheet = _createWorkSheet(presentationId, targetSheetName);
    const workListData = [];

    dataRows.forEach((row, index) => {
      const rowNum = index + 2; // 実際のシート行番号
      workListData.push([
        `Row_${rowNum}`, // TaskKey
        rowNum, // TaskData (行番号)
        STATUS_EMPTY, // Status
        mode, // Mode
        presentationId, SLIDES_TEMPLATE_ID_TR, TEMPLATE_SLIDE_INDEX_TR, combineRows,
        JSON.stringify(ALT_TEXT_TITLE_MAP_TR),
        IMAGE_ALT_TEXT_TITLE_TR,
        ILLUSTRATION_COLUMN_INDEX_TR
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 11).setValues(workListData);
    }

    _showSetupCompletionDialog({
      workSheetName: WORK_LIST_SHEET_NAME,
      menuItemName: '🌡️ 東海理科用 > 1-6 スライド生成(詳細情報)（実行）',
      processFunctionName: 'createSlides_PROCESS',
      useManualExecution: true
    });

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー (DetailTR):\n${e.message}`);
  }
}


/**
 * [SETUP] 複数行1スライド (SummaryTR) のセットアップ
 */
function createSlideSummaryTR_SETUP() {
  const ui = SpreadsheetApp.getUi();
  try {
    ss.toast('セットアップ (SummaryTR) を開始します...', '開始', 10);
    // --- 元の設定項目 ---
    const SLIDES_TEMPLATE_ID_TR = '1NYkmHwG4hHm8sadB_n15N6knXNGXtX3ZpLibePXfKS8';
    const TEMPLATE_SLIDE_INDEX_TR = 2;
    const ALT_TEXT_TITLE_MAP_TR = {
      "placeholder_equip": 3, "placeholder_line": 6, "placeholder_process": 8,
      "placeholder_trouble": 9, "placeholder_id": 0, "placeholder_place": 1,
      "placeholder_point_rough": 7, "placeholder_equip_num": 5, "placeholder_original_nums": 2,
      "placeholder_date": 4, "placeholder_title": 10, "placeholder_detail": 11,
      "placeholder_issue": 12, "placeholder_fix": 13, "placeholder_name": 14, "placeholder_original_num" : 2
    };
    const IMAGE_ALT_TEXT_TITLE_TR = false;
    const ILLUSTRATION_COLUMN_INDEX_TR = false;
    const combineRows = true;
    const mode = 'SummaryTR';
    const chunkSize = 5; // 1スライドにまとめる最大行数
    const groupingColumns = ["設備名称", "工程ブロック/資産No", "異常現象"];

    // --- 1. 対象シート取得 (元のロジック) ---
    const targetSheetName = tokaiPromptSheet.getRange("C15").getValue();
    if (!targetSheetName) throw new Error(`対象シート名が入力されていません。`);
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`データシート "${targetSheetName}" が見つかりません。`);

    // --- 2. ID採番 ---
    try {
      const masterSheetName = tokaiPromptSheet.getRange("C17").getValue();
      const id_col=1;
      const ID_PREFIX="EC-TY-";
      assignPersistentGroupIds_(sheet, masterSheetName, id_col, ID_PREFIX, groupingColumns);
      SpreadsheetApp.getActiveSpreadsheet().toast('グループIDをA列に採番・更新しました。', 'ID採番完了', 3);
    } catch (e) {
      throw new Error(`ID採番中にエラーが発生しました: ${e.message}`);
    }

    // --- 3. 新規プレゼンテーション作成 ---
    const newPresentationTitle = `事例一覧スライド_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
    const presentationId = _createAndMovePresentation(newPresentationTitle);

    // --- 4. データ取得 & グルーピング (createSlidesMainFuncから移植) ---
    const allData = sheet.getDataRange().getValues();
    const header = allData[0];
    const dataRows = allData.slice(1);

    const groupIndices = groupingColumns.map(colName => {
      const index = header.indexOf(colName);
      if (index === -1) throw new Error(`データシートのヘッダーに列名「${colName}」が見つかりません。`);
      return index;
    });

    // ★重要: グルーピングロジックを変更。行番号(index + 2)を格納する
    const groupedData = new Map(); // Map<グループキー, { rowNumbers: number[] }>
    dataRows.forEach((row, index) => {
      const groupKey = groupIndices.map(idx => row[idx]).join('|');
      
      // グループ化のキーが空欄の場合はスキップ (元のロジックにはなかったが、ID採番ロジックに合わせて追加)
      if (groupIndices.map(idx => row[idx]).some(val => val === null || val === "")) {
        return; 
      }

      if (!groupedData.has(groupKey)) {
        groupedData.set(groupKey, { rowNumbers: [] });
      }
      groupedData.get(groupKey).rowNumbers.push(index + 2); // 実際のシート行番号を格納
    });

    if (groupedData.size === 0) throw new Error('グルーピング対象のデータが0件です。');

    // --- 5. 作業シート作成 & タスク書き込み (チャンク単位) ---
    const workSheet = _createWorkSheet(presentationId, targetSheetName);
    const workListData = [];

    for (const [groupKey, groupInfo] of groupedData.entries()) {
      const groupRowNumbers = groupInfo.rowNumbers; // [2, 5, 10, 11, 12, 15]

      // チャンキング
      for (let i = 0; i < groupRowNumbers.length; i += chunkSize) {
        const chunkRowNumbers = groupRowNumbers.slice(i, i + chunkSize); // [2, 5, 10, 11, 12]
        
        workListData.push([
          `${groupKey}|Chunk${i}`, // TaskKey (一意にする)
          JSON.stringify(chunkRowNumbers), // TaskData (行番号配列)
          STATUS_EMPTY, // Status
          mode, // Mode
          presentationId, SLIDES_TEMPLATE_ID_TR, TEMPLATE_SLIDE_INDEX_TR, combineRows,
          JSON.stringify(ALT_TEXT_TITLE_MAP_TR), // AltTextMap
          IMAGE_ALT_TEXT_TITLE_TR, // ImageAltText
          ILLUSTRATION_COLUMN_INDEX_TR  // ImageColIndex
        ]);
      }
    }

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 11).setValues(workListData);
    }

    _showSetupCompletionDialog({
      workSheetName: WORK_LIST_SHEET_NAME,
      menuItemName: '🌡️ 東海理科用 > 2-2 スライド生成(まとめ一覧)（実行）',
      processFunctionName: 'createSlides_PROCESS',
      useManualExecution: true
    });

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー (SummaryTR):\n${e.message}`);
  }
}

// ===================================================================
// STEP 2: PROCESS関数 (ワーカー)
// ===================================================================

/**
 * [PROCESS] スライド生成バッチ処理ワーカー
 * この関数を5分ごとなどの時間ベーストリガーで実行します。
 */
function createSlides_PROCESS() {
  const startTime = new Date().getTime();
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

  const workSheet = ss.getSheetByName(WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  _showProgress('スライド生成処理を開始します...', '📽️ スライド生成', 3);

  // --- 1. 共通設定（プレゼンテーションID、対象シート名）を作業シートから取得 ---
  // (D1セル、E1セルに保存したと仮定)
  const presentationId = workSheet.getRange("D1").getValue();
  const targetSheetName = workSheet.getRange("E1").getValue();

  if (!presentationId || !targetSheetName) {
    Logger.log("作業シート D1 または E1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  let presentation;
  let inputSheet;
  let allData;
  try {
    presentation = SlidesApp.openById(presentationId);
    inputSheet = ss.getSheetByName(targetSheetName);
    if (!inputSheet) throw new Error(`入力シート ${targetSheetName} が見つかりません。`);
    allData = inputSheet.getDataRange().getValues(); // ★全データを一度だけ読み込む
  } catch (e) {
    Logger.log(`必須リソース（プレゼンテーション, 入力シート）が開けません: ${e}`);
    return; // 処理不可
  }

  // --- 2. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 11); // 11列分取得
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;

  // --- 3. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][2]; // C列: Status

    // 未処理のタスクか？
    if (currentStatus === STATUS_EMPTY) {

      // 動的タイムアウトチェック：次のタスクを実行可能かを判定
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        break; // 次のトリガー実行に任せる
      }

      const taskStartTime = new Date().getTime();
      const sheetRow = i + 2; // スプレッドシートの実際の行番号

      // タスク情報を取得
      const taskKey = workValues[i][0];
      const taskDataJson = workValues[i][1];
      // const mode = workValues[i][3]; // (参考用)
      const templateId = workValues[i][5];
      const templateIndex = workValues[i][6];
      const combineRows = workValues[i][7];
      const altTextMap = JSON.parse(workValues[i][8]);
      const imageAltText = workValues[i][9];
      const imageColIndex = workValues[i][10];

      let templateSlide; // テンプレートスライドはタスクごとに取得

      try {
        // 3a. ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);
        
        templateSlide = SlidesApp.openById(templateId).getSlides()[templateIndex];
        if (!templateSlide) {
          throw new Error(`テンプレートスライド (ID: ${templateId}, Index: ${templateIndex}) が見つかりません。`);
        }

        // 3b. タスク実行 (combineRows フラグに基づいて処理を分岐)
        if (combineRows === false) {
          // --- 1行1スライド (Tomy, DetailTR) ---
          const rowNum = JSON.parse(taskDataJson); // 行番号 (e.g. 3)
          const row = allData[rowNum - 1]; // allData (0-indexed) から行データを復元
          
          _transferSingleRowToSlide(
            presentation,
            templateSlide,
            row,
            rowNum,
            altTextMap,
            imageAltText,
            imageColIndex
          );

        } else {
          // --- 複数行1スライド (SummaryTR) ---
          const chunkRowNumbers = JSON.parse(taskDataJson); // 行番号配列 (e.g. [2, 5, 10])
          const chunk = chunkRowNumbers.map(rowNum => allData[rowNum - 1]); // allDataからチャンクデータを復元
          const startRowNumForLog = chunkRowNumbers[0] || (i + 2);

          // SummaryTR の Map を再構築 (元のロジック)
          const entries = Object.entries(altTextMap);
          const inputOnceMap = Object.fromEntries(entries.slice(0, 4));
          const combinedMap = Object.fromEntries(entries.slice(4, 9));
          const detailMap = Object.fromEntries(entries.slice(9,));

          // ★元の _transferChunkToSlide_ 関数をそのまま呼び出す
          _transferChunkToSlide_(
            presentation,
            templateSlide,
            chunk,
            startRowNumForLog,
            inputOnceMap,
            combinedMap,
            detailMap
          );
        }

        // 3c. 待機 (元のロジック)
        Utilities.sleep(SLEEP_MS_PER_SLIDE);

        // 3d. ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        // 3件ごとに進捗を表示（スライド生成は時間がかかるため頻度を下げる）
        if (processedCountInThisRun % 3 === 0) {
          const totalTasks = workValues.length;
          _showProgress(
            `${processedCountInThisRun} / ${totalTasks} 件完了`,
            '📽️ スライド生成中',
            2
          );
        }

        SpreadsheetApp.flush();

      } catch (e) {
        // 3e. エラー処理
        Logger.log(`タスク "${taskKey}" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

        // エラーの場合も実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
      }
    } // End if (status_empty)
  } // End for loop

  Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);
  SpreadsheetApp.flush(); // シートへの書き込みを強制的に反映させる

  // --- 4. 完了チェック ---
  const lastRow = workSheet.getLastRow();
  console.log("last row is "+ lastRow);
  let remainingTasks = 0;
  if (lastRow >= 2) {
    const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    remainingTasks = newStatusValues.filter(
      row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
    ).length;
  }
  console.log("remaining tasks are "+remainingTasks);

  // 「今回の実行で処理したタスクがあり」かつ「（最新のステータスで）残タスクが0になった」場合
  if (remainingTasks === 0 && processedCountInThisRun > 0) {
    Logger.log("すべてのタスクが完了しました。");

    try {
      // 4a. 最初の空スライドを削除 (元のロジック)
      const finalPresentation = SlidesApp.openById(presentationId); // 再度開く
      const initialSlide = finalPresentation.getSlides()[0];
      if (initialSlide && finalPresentation.getSlides().length > 1) {
        initialSlide.remove();
        Logger.log("最初の空スライドを削除しました。");
      }
      
      // 4b. 完了通知
      const presentationUrl = finalPresentation.getUrl();
      Logger.log(`処理完了。プレゼンテーションURL: ${presentationUrl}`);
      _showProgress('すべてのスライド生成が完了しました！', '✅ 完了', 10);

      // 手動実行時のみアラート表示
      if (_isManualExecution()) {
        ui.alert('成功', `プレゼンテーションを作成しました: ${finalPresentation.getName()}\nURL: ${presentationUrl}`, ui.ButtonSet.OK);
      }

      // 4c. トリガーを停止
      stopTriggers_('createSlides_PROCESS');

    } catch (e) {
      Logger.log(`完了処理（空スライド削除、トリガー停止）中にエラー: ${e}`);
    }
  }
}

// ===================================================================
// STEP 3: ヘルパー関数 (新規・変更・流用)
// ===================================================================

/**
 * [新規] 1行1スライドの転記処理 (createSlidesMainFunc の
 * * * ブロックから移植)
 */
function _transferSingleRowToSlide(presentation, templateSlide, row, rowNumForLog, altTextMap, imageAltText, imageColIndex) {
  
  // この関数内は、元の createSlidesMainFunc の `else` (1行1スライド) ブロックの
  // `try...catch` の中身とほぼ同じ
  
  const newSlide = presentation.insertSlide(presentation.getSlides().length, templateSlide);
  const pageElements = newSlide.getPageElements();

  // --- 日付挿入 ---
  try {
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const datePlaceholder = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === "placeholder_created_date")?.asShape();
    if (datePlaceholder && datePlaceholder.getText) {
      datePlaceholder.getText().setText(formattedDate);
    } else {
      Logger.log(`情報(行 ${rowNumForLog}): 代替テキスト "placeholder_created_date" が見つかりません。`);
    }
  } catch (e) {
    Logger.log(`警告(行 ${rowNumForLog}): 日付挿入処理でエラー - ${e}`);
  }

  // --- テキスト置換 ---
  for (const altTextTitle in altTextMap) {
    const colIndex = altTextMap[altTextTitle];
    if (colIndex >= 0 && colIndex < row.length) {
      let replacementValue = row[colIndex];
      if (replacementValue instanceof Date) {
        replacementValue = Utilities.formatDate(replacementValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      }
      const shape = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === altTextTitle)?.asShape();
      if (shape && shape.getText) {
        shape.getText().setText(String(replacementValue || ''));
      } else {
        Logger.log(`警告: 行 ${rowNumForLog}: 代替テキスト "${altTextTitle}" が見つかりません。`);
      }
    } else if (colIndex !== -1) {
      Logger.log(`警告: 行 ${rowNumForLog}: 代替テキスト "${altTextTitle}" の列インデックス ${colIndex} が範囲外です。`);
    }
  }

  // --- 画像置換 (imageAltTextが指定されている場合のみ) ---
  if (imageAltText && imageColIndex !== false && imageColIndex >= 0) {
    const imageSource = row[imageColIndex];
    let imageBlob = null;

    if (typeof imageSource === 'string' && imageSource.toLowerCase().startsWith('http')) {
      const fileId = extractGoogleDriveId_(imageSource);
      if (fileId) { try { imageBlob = DriveApp.getFileById(fileId).getBlob(); } catch (e) { Logger.log(`警告: 行 ${rowNumForLog}: Driveファイル取得失敗 - ${e}`); } }
      else { try { imageBlob = UrlFetchApp.fetch(imageSource).getBlob(); } catch (e) { Logger.log(`警告: 行 ${rowNumForLog}: URL画像取得失敗 - ${e}`); } }
    } else if (typeof imageSource === 'object' && imageSource !== null && imageSource.toString() === 'CellImage') {
      try { const imageUrl = imageSource.getContentUrl(); if (imageUrl) { imageBlob = UrlFetchApp.fetch(imageUrl).getBlob(); } else { Logger.log(`警告: 行 ${rowNumForLog}: CellImage URL取得不可`); } }
      catch(e) { Logger.log(`警告: 行 ${rowNumForLog}: CellImage処理エラー - ${e}`); }
    }

    if (imageBlob) {
        const imagePlaceholder = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.IMAGE && el.getTitle() === imageAltText)?.asImage();
        if (imagePlaceholder) {
          imagePlaceholder.replace(imageBlob);
          Logger.log(`行 ${rowNumForLog}: 画像(タイトル: ${imageAltText})を置換しました。`);
        } else {
          Logger.log(`警告: 行 ${rowNumForLog}: 代替テキスト "${imageAltText}" を持つ画像が見つかりません。`);
        }
    } else if (imageSource){
      Logger.log(`警告: 行 ${rowNumForLog}: 列 ${imageColIndex + 1} の画像ソースを処理できませんでした。ソース: ${imageSource}`);
    }
  }
}


/**
 * [新規] 作業シート（_SlideWorkList）を作成するヘルパー関数
 * @param {string} presentationId - 新規作成したスライドのID
 * @param {string} targetSheetName - 読み込み元のシート名
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成またはクリアされた作業シート
 */
function _createWorkSheet(presentationId, targetSheetName) {
  let workSheet = ss.getSheetByName(WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear(); // 既存のシートをクリア
  } else {
    workSheet = ss.insertSheet(WORK_LIST_SHEET_NAME, 0);
  }
  
  const workHeader = [
    "TaskKey", "TaskData (JSON or RowNum)", "Status", "Mode",
    "PresentationID", "TemplateID", "TemplateIndex", "CombineRows",
    "AltTextMap (JSON)", "ImageAltText", "ImageColIndex"
  ];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');
  
  // D1, E1 にトリガー実行で必要な情報を保存
  workSheet.getRange("D1").setValue(presentationId);
  workSheet.getRange("E1").setValue(targetSheetName);

  // タブの色をグレーに設定
  workSheet.setTabColor('#999999');

  workSheet.autoResizeColumn(1);
  return workSheet;
}

/**
 * [新規] 新規プレゼンテーションを作成し、指定フォルダに移動するヘルパー関数
 * (元の createSlidesMainFunc の冒頭部分)
 * @param {string} newPresentationTitle - 新規スライドのタイトル
 * @return {string} 新規作成されたプレゼンテーションのID
 */
function _createAndMovePresentation(newPresentationTitle) {
  // --- 保存先フォルダの指定 (元のロジック) ---
  // (promptSheet と slideSaveDir_pos はグローバル定義されている前提)
  const outputFolderUrl = promptSheet.getRange(slideSaveDir_pos).getValue();
  let outputFolder = null; 

  if (outputFolderUrl) {
    const folderId = _extractFolderIdFromUrl(outputFolderUrl);
    if (folderId) {
      try {
        outputFolder = DriveApp.getFolderById(folderId);
      } catch (e) {
        Logger.log(`警告: 指定されたフォルダ(ID: ${folderId})が見つからないかアクセスできません。ルートに保存します。`);
        ui.alert('警告', `指定された保存先フォルダが見つからないかアクセスできません。\nマイドライブのルートに保存します。`, ui.ButtonSet.OK);
      }
    } else {
      Logger.log(`警告: ${slideSaveDir_pos}セルのURLからフォルダIDを取得できませんでした。ルートに保存します。`);
      ui.alert('警告', `${slideSaveDira_pos}セルのURLが正しくありません。\nマイドライブのルートに保存します。`, ui.ButtonSet.OK);
    }
  } else {
    Logger.log("保存先フォルダの指定がないため、マイドライブのルートに保存します。");
  }

  // --- プレゼンテーションの作成 & 移動 (元のロジック) ---
  const tempPresentation = SlidesApp.create(newPresentationTitle);
  const presentationId = tempPresentation.getId();
  const presentationFile = DriveApp.getFileById(presentationId);
  
  if (outputFolder) {
    try {
      presentationFile.moveTo(outputFolder);
      Logger.log(`プレゼンテーションをフォルダ「${outputFolder.getName()}」に移動しました。`);
    } catch (moveError) {
       Logger.log(`警告: フォルダへの移動に失敗。ルートに残ります。エラー: ${moveError}`);
       ui.alert('警告', `プレゼンテーションを指定フォルダへ移動できませんでした。\nマイドライブのルートに保存されています。`, ui.ButtonSet.OK);
    }
  }
  return presentationId; // ★IDを返す
}

// ===================================================================
// 注: 以下の共通ヘルパー関数は commonHelpers.js に移動しました
// - _showSetupCompletionDialog()
// - stopTriggers_()
// - extractGoogleDriveId_() (一部)
// - _extractFolderIdFromUrl()
// ===================================================================

// --- 以下、元のコードから変更不要なヘルパー関数 ---
// ( _transferChunkToSlide_, extractGoogleDriveId_, _extractFolderIdFromUrl, assignGroupIdsToSheet )
// ... (元のコードをそのままコピーしてください) ...

/**
 * [新規] スプレッドシートの複数行データ(チャンク)を、1枚のGoogleスライドに転記する関数
 * @param {SlidesApp.Presentation} presentation - 書き込み先のプレゼンテーションオブジェクト
 * @param {SlidesApp.Slide} templateSlide - 複製元のテンプレートスライドオブジェクト
 * @param {Array[]} chunk - 転記するデータ行の配列 (最大5行)
 * @param {Object} detailMap - 事例一覧として個別詳細を入れるテキスト要素の代替テキストと列インデックスのマッピング
 * @param {number} startRowNumForLog - ログ表示用の開始行番号
 */
function _transferChunkToSlide_(presentation, templateSlide, chunk, startRowNumForLog, inputOnceMap, combinedMap, detailMap,) {
  if (!chunk || chunk.length === 0) return;

  // --- (日付ソート処理) ---
  try {
    // detailMap から "placeholder_date" の列インデックスを取得
    const dateColIndex = detailMap["placeholder_date"];
    
    // dateColIndexが 0 以上（有効）の場合のみソートを実行
    if (dateColIndex !== undefined && dateColIndex >= 0) {
      Logger.log(`ソートキー "placeholder_date" (列インデックス ${dateColIndex}) に基づいてチャンクをソートします。`);
      
      chunk.sort((a, b) => {
        const valA = a[dateColIndex];
        const valB = b[dateColIndex];

        // new Date() は Date オブジェクト、日付文字列の両方を扱える
        const dateA = new Date(valA);
        const dateB = new Date(valB);

        const timeA = dateA.getTime();
        const timeB = dateB.getTime();

        // 不正な日付 (Invalid Date) の getTime() は NaN を返す
        // 不正な日付は末尾に配置する
        if (isNaN(timeA) && isNaN(timeB)) {
          return 0; // 両方不正なら順序変更なし
        }
        if (isNaN(timeA)) {
          return 1; // A (a) が不正なら、a を b の後ろに
        }
        if (isNaN(timeB)) {
          return -1; // B (b) が不正なら、b を a の後ろに (a を b の前に)
        }

        // 古い順 (昇順)
        return timeB - timeA;
      });
      
      Logger.log("ソートが完了しました。");
    } else {
      Logger.log(`ソートキー "placeholder_date" が detailMap に見つからないか無効なため、ソートをスキップします。`);
    }
  } catch (e) {
    Logger.log(`警告: チャンクの日付ソート中にエラーが発生しました: ${e}。ソートせずに処理を続行します。`);
  }
  
  const newSlide = presentation.insertSlide(presentation.getSlides().length, templateSlide);
  const pageElements = newSlide.getPageElements();
  const chunkRowCount = chunk.length;
  const chunkFirstData = chunk[0];

  // --- ▼ここから追加▼ (日付挿入) ---
  try {
    const today = new Date();
    // 日付を 'yyyy/MM/dd' 形式にフォーマット
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    // "placeholder_created_date" という代替テキスト（タイトル）を持つ図形(Shape)を探す
    const datePlaceholder = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === "placeholder_created_date")?.asShape();
    
    if (datePlaceholder && datePlaceholder.getText) {
      datePlaceholder.getText().setText(formattedDate);
    } else {
      // プレースホルダーが見つからなくても処理は続行し、ログに警告を残す
      Logger.log(`情報(チャンク ${startRowNumForLog}行目〜): 代替テキスト "placeholder_created_date" がスライドテンプレートに見つかりません。`);
    }
  } catch (e) {
    Logger.log(`警告(チャンク ${startRowNumForLog}行目〜): 日付挿入処理でエラーが発生しました - ${e}`);
  }
  // --- ▲ここまで追加▲ ---

  // Group情報だけ先にスライドに入れる
  for (const baseAltText in inputOnceMap) {
    const colIndex = inputOnceMap[baseAltText];
    if (colIndex >= 0 && colIndex < chunkFirstData.length) {
      const targetAltText = baseAltText
      let replacementValue = chunkFirstData[colIndex];
      // console.log("replacementValue is "+replacementValue)
      const shape = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === targetAltText)?.asShape();
      if (shape && shape.getText) {
        shape.getText().setText(String(replacementValue || ''));
      } else {
        Logger.log(`警告： テキスト "${targetAltText}" がスライドに見つかりません。`);
      }
    }
  }

  // 1. 変数をリスト型（配列）で宣言
  let combinedListId = [];
  let combinedListPlace = [];
  let combinedListPointRough = [];
  let combinedListEquipNum = [];
  let combinedListOriginalNum = [];
  let combinedTextId;
  let combinedTextPlace;
  let combinedTextPointRough;
  let combinedTextEquipNum;
  let combinedTextOriginalNum;

  for (let i = 0; i < chunkRowCount; i++) {
    const rowData = chunk[i];
    const rowNumSuffix = `_${i + 1}`; // "_1", "_2", ...
    const currentRowNumForLog = startRowNumForLog + i;
    Logger.log(`  - 行 ${currentRowNumForLog} のデータをスライド ${newSlide.getObjectId()} に転記 (セット ${i + 1})`);

    // テキスト置換
    for (const baseAltText in detailMap) {
      const colIndex = detailMap[baseAltText];
      if (colIndex >= 0 && colIndex < rowData.length) {
        const targetAltText = baseAltText + rowNumSuffix; // 例: "placeholder_title_1"
        let replacementValue = rowData[colIndex];
        // console.log("replacementValue is "+replacementValue)
        if (replacementValue instanceof Date) {
          replacementValue = Utilities.formatDate(replacementValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        }
        const shape = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === targetAltText)?.asShape();
        if (shape && shape.getText) {
          shape.getText().setText(String(replacementValue || ''));
        } else {
          Logger.log(`警告(行 ${currentRowNumForLog}): テキスト "${targetAltText}" がスライドに見つかりません。`);
        }
      }
    }

    for (const baseAltText in combinedMap) {
      const colIndex = combinedMap[baseAltText];
      if (colIndex >= 0 && colIndex < rowData.length) {
        // 2. リストに加える形で情報を追加
        switch (baseAltText) {
          case "placeholder_id":
            combinedListId.push(rowData[colIndex]);
            break;
          case "placeholder_place":
            combinedListPlace.push(rowData[colIndex]);
            break;
          case "placeholder_point_rough":
            combinedListPointRough.push(rowData[colIndex]);
            break;
          case "placeholder_equip_num":
            combinedListEquipNum.push(rowData[colIndex]);
            break;
          case "placeholder_original_nums":
            combinedListOriginalNum.push(rowData[colIndex]);
            break;
        }
      }
    }

    // 3. 重複データを削除 (Setを使って一意な値のみを取得)
    // 4. リスト内のデータをカンマ区切りで繋いだテキストを生成
    // [...new Set(配列)] で重複を削除した新しい配列を作成し、.join() で連結します。
    combinedTextId = [...new Set(combinedListId)].join(', ');
    combinedTextPlace = [...new Set(combinedListPlace)].join(', ');
    combinedTextPointRough = [...new Set(combinedListPointRough)].join(', ');
    combinedTextEquipNum = [...new Set(combinedListEquipNum)].join(', ');
    combinedTextOriginalNum = [...new Set(combinedListOriginalNum)].join(', ');

  }
  // console.log("combinedTextId is "+combinedTextId)

  // // 結合したテキストデータを所定のテキストボックスに格納する
  for (const conbinedTargetAltText in combinedMap) {
    const colIndex = combinedMap[conbinedTargetAltText];
    let combinedText;
    switch (conbinedTargetAltText) {
      case "placeholder_id":
        combinedText = combinedTextId;
        break;
      case "placeholder_place":
        combinedText = combinedTextPlace;
        break;
      case "placeholder_point_rough":
        combinedText = combinedTextPointRough;
        break;
      case "placeholder_equip_num":
        combinedText = combinedTextEquipNum;
        break;
      case "placeholder_original_nums":
        combinedText = combinedTextOriginalNum;
        break;
    }
    // console.log("combinedText is "+combinedText)
    const shapeForCombinedText = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === conbinedTargetAltText)?.asShape();
    if (shapeForCombinedText && shapeForCombinedText.getText) {
      shapeForCombinedText.getText().setText(String(combinedText || ''));
    } else {
      // Logger.log(`警告(行 ${currentRowNumForLog}): テキスト "${conbinedTargetAltText}" がスライドに見つかりません。`);
      // ↑ currentRowNumForLog がこのスコープにないためコメントアウト
      Logger.log(`警告: 結合テキスト "${conbinedTargetAltText}" がスライドに見つかりません。`);
    }
  }
  // End loop for rows within chunk

}

// ===================================================================
// 注: extractGoogleDriveId_() と _extractFolderIdFromUrl() は
// commonHelpers.js に移動しました
// ===================================================================


// --- ★変更: ID採番用の関数を新設 ---
/**
 * [新規] スプレッドシートのA列にグループIDを採番して書き込む関数
 * groupingColumns は createSlidesMainFunc の定義に合わせる
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のシートオブジェクト
 */
function assignGroupIdsToSheet(sheet) {
  const allData = sheet.getDataRange().getValues();
  const header = allData[0];
  const dataRows = allData.slice(1);

  if (dataRows.length === 0) {
    Logger.log("ID採番: データ行がありません。");
    return; // データがなければ何もしない
  }

  
  const groupIndices = groupingColumns.map(colName => {
    const index = header.indexOf(colName);
    if (index === -1) {
      // ヘッダーに見つからない場合はエラーを投げる
      throw new Error(`ID採番エラー: データシートのヘッダーに列名「${colName}」が見つかりません。`);
    }
    return index;
  });

  // --- グルーピング実行 (元の行インデックス[0始まり]を保持) ---
  const groupedData = new Map(); // Map<グループキー, { originalIndices: number[] }>
  
  dataRows.forEach((row, index) => {
    // グループ化のキーとなる値を取得
    const keyValues = groupIndices.map(idx => row[idx]);
    
    // キーのいずれかが空欄の場合、その行はグループ化対象外とする
    if (keyValues.some(val => val === null || val === "")) {
      return; 
    }
    
    // グループキーを作成
    const groupKey = keyValues.join('|'); 
    
    if (!groupedData.has(groupKey)) {
      groupedData.set(groupKey, { originalIndices: [] });
    }
    // 0始まりの行インデックスをグループに追加
    groupedData.get(groupKey).originalIndices.push(index); 
  });

  // --- IDの生成と書き込み準備 ---
  let idCounter = 1;
  // dataRows.length 分の配列を [""] (空欄) で初期化
  const idsToWrite = Array.from({ length: dataRows.length }, () => [""]); 

  // グループ化されたデータにIDを割り当て
  for (const [groupKey, groupInfo] of groupedData.entries()) {
    
    // IDを "EC-TY001" 形式で生成
    const newId = "EC-TY" + String(idCounter++).padStart(3, '0');
    
    // このグループに属するすべての行の、書き込み用配列 (idsToWrite) の対応する位置にIDをセット
    groupInfo.originalIndices.forEach(index => {
      idsToWrite[index] = [newId];
    });
  }

  // --- シートへの書き込み (A列の2行目から) ---
  if (idsToWrite.length > 0) {
    // getRange(開始行, 開始列, 行数, 列数)
    sheet.getRange(2, 1, idsToWrite.length, 1).setValues(idsToWrite);
    Logger.log(`${groupedData.size} グループ (${idCounter - 1} 個) のIDをシートA列に書き込みました。`);
  }
}

// ===================================================================
// ★新設: 永続化対応 ID採番関数
// ===================================================================
/**
 * [新規] スプレッドシートのA列に「永続化された」グループIDを採番して書き込む関数
 * IDのマスターリスト（_GroupID_MasterList）を参照・更新する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のデータシートオブジェクト
 */
function assignPersistentGroupIds_(sheet, masterSheetName, id_col, ID_PREFIX, groupingColumns) {
  const MASTER_LIST_SHEET_NAME = masterSheetName;
  console.log(MASTER_LIST_SHEET_NAME);
  
  const ss = sheet.getParent(); // スプレッドシート本体を取得

  // --- 1. IDマスターリストを読み込む ---
  let masterSheet = ss.getSheetByName(MASTER_LIST_SHEET_NAME);
  const idMap = new Map();
  let maxIdNum = 0;

  if (!masterSheet) {
    // マスターシートが存在しない場合は作成する
    masterSheet = ss.insertSheet(MASTER_LIST_SHEET_NAME, 0);
    masterSheet.getRange("A1:B1").setValues([["GroupKey", "AssignedID"]]).setFontWeight("bold");
    Logger.log(`IDマスターリストシート "${MASTER_LIST_SHEET_NAME}" を作成しました。`);
  } else {
    // 既存のマスターリストを読み込む
    const lastRow = masterSheet.getLastRow();
    if (lastRow >= 2) {
      const masterData = masterSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      masterData.forEach(([key, id]) => {
        if (key && id) {
          idMap.set(key, id);
          // IDの最大値を取得 (例: "EC-TY005" -> 5)
          const num = parseInt(String(id).replace(ID_PREFIX, ""), 10);
          if (!isNaN(num) && num > maxIdNum) {
            maxIdNum = num;
          }
        }
      });
    }
  }
  
  // 次に採番するIDカウンターをセット (最大値 + 1)
  let nextIdCounter = maxIdNum + 1;
  Logger.log(`IDマスターリストを読み込みました。既存 ${idMap.size} 件。次のID: ${nextIdCounter}`);

  // --- 2. データシートを読み込み、グルーピング (元のロジック) ---
  const allData = sheet.getDataRange().getValues();
  const header = allData[0];
  const dataRows = allData.slice(1);

  if (dataRows.length === 0) {
    Logger.log("ID採番: データ行がありません。");
    return;
  }

  
  const groupIndices = groupingColumns.map(colName => {
    const index = header.indexOf(colName);
    if (index === -1) throw new Error(`ID採番エラー: ヘッダーに「${colName}」が見つかりません。`);
    return index;
  });

  const groupedData = new Map(); // Map<グループキー, { originalIndices: number[] }>
  dataRows.forEach((row, index) => {
    const keyValues = groupIndices.map(idx => row[idx]);
    if (keyValues.some(val => val === null || val === "")) {
      return; 
    }
    const groupKey = keyValues.join('|'); 
    if (!groupedData.has(groupKey)) {
      groupedData.set(groupKey, { originalIndices: [] });
    }
    groupedData.get(groupKey).originalIndices.push(index); 
  });

  // --- 3. IDの割り当て (★改善ロジック) ---
  const idsToWrite = Array.from({ length: dataRows.length }, () => [""]); 
  const newMasterListEntries = []; // マスターリストに追記する新しいペア

  for (const [groupKey, groupInfo] of groupedData.entries()) {
    let assignedId;
    
    if (idMap.has(groupKey)) {
      // 既存のグループ: マスターからIDを取得
      assignedId = idMap.get(groupKey);
    } else {
      // 新規のグループ: 新しいIDを採番
      assignedId = ID_PREFIX + String(nextIdCounter++).padStart(5, '0');
      // メモリ上のMapと、追記用リストに追加
      idMap.set(groupKey, assignedId);
      newMasterListEntries.push([groupKey, assignedId]);
    }
    
    // このIDを、該当するすべてのデータ行にセット
    groupInfo.originalIndices.forEach(index => {
      idsToWrite[index] = [assignedId];
    });
  }

  // --- 4. データシート (A列) への書き込み ---
  if (idsToWrite.length > 0) {
    sheet.getRange(2, id_col, idsToWrite.length, 1).setValues(idsToWrite);
    Logger.log(`データシートのA列にIDを書き込みました。`);
  }

  // --- 5. IDマスターリストへの追記 ---
  if (newMasterListEntries.length > 0) {
    masterSheet.getRange(masterSheet.getLastRow() + 1, 1, newMasterListEntries.length, 2)
      .setValues(newMasterListEntries);
    Logger.log(`${newMasterListEntries.length} 件の新規IDをマスターリストに追記しました。`);
  } else {
    Logger.log(`新規に採番されたIDはありませんでした。`);
  }
}

// ===================================================================
// スライドテンプレートマスタ関連
// ===================================================================

/**
 * スライドテンプレートマスタから設定を取得する
 * マスタシート構造:
 *   A列: GoogleスライドID (テンプレートID)
 *   B列: テンプレート名
 *   C列: スライドIndex
 *   D列: ALT_TEXT_TITLE_MAP (JSON)
 *   E列: IMAGE_ALT_TEXT
 *   F列: IMAGE_COL_INDEX
 *
 * @param {string} templateId - GoogleスライドID
 * @return {Object|null} テンプレート設定オブジェクト、見つからない場合はnull
 */
function _getSlideTemplateConfig(templateId) {
  const masterSheet = ss.getSheetByName(SLIDE_TEMPLATE_MASTER_SHEET_NAME);
  if (!masterSheet) {
    throw new Error(`マスタシート「${SLIDE_TEMPLATE_MASTER_SHEET_NAME}」が見つかりません。`);
  }

  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error(`マスタシート「${SLIDE_TEMPLATE_MASTER_SHEET_NAME}」にテンプレートが登録されていません。`);
  }

  const data = masterSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === templateId) {
      // E列・F列が空白の場合はfalseを設定（画像処理をスキップ）
      const imageAltText = data[i][4] !== "" ? data[i][4] : false;
      const imageColIndex = data[i][5] !== "" ? data[i][5] : false;

      return {
        templateId: data[i][0],           // A列: GoogleスライドID
        templateName: data[i][1],         // B列: テンプレート名
        slideIndex: data[i][2],           // C列: スライドIndex
        altTextTitleMap: JSON.parse(data[i][3]), // D列: ALT_TEXT_TITLE_MAP (JSON)
        imageAltText: imageAltText,       // E列: IMAGE_ALT_TEXT（空白ならfalse）
        imageColIndex: imageColIndex      // F列: IMAGE_COL_INDEX（空白ならfalse）
      };
    }
  }

  return null; // 見つからない場合
}

/**
 * [SETUP] テンプレートマスタを使用した汎用スライド生成セットアップ
 * promptシートのC16セルからテンプレートID（GoogleスライドID）を取得
 */
function createSlideFromTemplate_SETUP() {
  const ui = SpreadsheetApp.getUi();
  try {
    ss.toast('セットアップを開始します...', '開始', 10);

    // --- 1. テンプレートIDをpromptシートから取得 ---
    const templateId = promptSheet.getRange('C16').getValue();
    if (!templateId) {
      throw new Error('promptシートのC16セルにテンプレートID（GoogleスライドID）が入力されていません。');
    }

    // --- 2. マスタからテンプレート設定を取得 ---
    const config = _getSlideTemplateConfig(templateId);
    if (!config) {
      throw new Error(`テンプレートID「${templateId}」がマスタシートに登録されていません。`);
    }

    Logger.log(`テンプレート「${config.templateName}」を使用します。`);

    // --- 3. 対象シート取得 ---
    const targetSheetName = promptSheet.getRange(generateSlidesSheetName_pos).getValue();
    if (!targetSheetName) throw new Error(`promptシートのC13セルに対象シート名が入力されていません。`);
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`データシート "${targetSheetName}" が見つかりません。`);

    // --- 4. 新規プレゼンテーション作成 ---
    const newPresentationTitle = `詳細事例スライド_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
    const presentationId = _createAndMovePresentation(newPresentationTitle);

    // --- 5. データ行取得 ---
    const allData = sheet.getDataRange().getValues();
    const dataRows = allData.slice(1);
    if (dataRows.length === 0) throw new Error('シートにデータが見つかりません（ヘッダーを除く）。');

    // --- 6. 作業シート作成 & タスク書き込み ---
    const workSheet = _createWorkSheet(presentationId, targetSheetName);
    const workListData = [];

    const mode = 'Template'; // 汎用モード
    const combineRows = false;

    dataRows.forEach((_, index) => {
      const rowNum = index + 2; // 実際のシート行番号
      workListData.push([
        `Row_${rowNum}`,                    // TaskKey
        rowNum,                              // TaskData (行番号)
        STATUS_EMPTY,                        // Status
        mode,                                // Mode
        presentationId,                      // PresentationID
        config.templateId,                   // TemplateID
        config.slideIndex,                   // TemplateIndex
        combineRows,                         // CombineRows
        JSON.stringify(config.altTextTitleMap), // AltTextMap (JSON)
        config.imageAltText,                 // ImageAltText
        config.imageColIndex                 // ImageColIndex
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 11).setValues(workListData);
    }

    _showSetupCompletionDialog({
      workSheetName: WORK_LIST_SHEET_NAME,
      menuItemName: '📽️ スライド生成 > ⑦_2 スライド生成（実行）',
      processFunctionName: 'createSlides_PROCESS',
      useManualExecution: true
    });

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}