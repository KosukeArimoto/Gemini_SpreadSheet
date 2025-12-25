
// ===================================================================
// STEP 1: SETUP関数
// ===================================================================


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
    Logger.log(`conditionalBgColors設定: ${JSON.stringify(config.conditionalBgColors)}`);

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
        config.imageColIndex,                // ImageColIndex
        config.conditionalBgColors ? JSON.stringify(config.conditionalBgColors) : "" // ConditionalBgColors (JSON)
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 12).setValues(workListData);
    }

    // シートへの書き込みを即座に完了
    SpreadsheetApp.flush();

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

/**
 * [SETUP] 1行1スライド (DetailTR) のセットアップ - 統合モード
 * すべてのスライドを1つのプレゼンテーションに生成
 */
function createSlideDetailTR_Combined_SETUP() {
  _createSlideDetailTR_SETUP_Internal(false); // 統合モード
}

/**
 * [SETUP] 1行1スライド (DetailTR) のセットアップ - 分割モード
 * グループごとに別々のプレゼンテーションを生成
 */
function createSlideDetailTR_Split_SETUP() {
  _createSlideDetailTR_SETUP_Internal(true); // 分割モード
}

/**
 * [内部] DetailTR セットアップの共通ロジック
 * @param {boolean} isSplitMode - true: 分割モード, false: 統合モード
 */
function _createSlideDetailTR_SETUP_Internal(isSplitMode) {
  const ui = SpreadsheetApp.getUi();
  try {
    const modeLabel = isSplitMode ? '分割モード' : '統合モード';
    ss.toast(`セットアップ (DetailTR - ${modeLabel}) を開始します...`, '開始', 10);

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
    const IMAGE_ALT_TEXT_TITLE_TR = 'placeholder_image';
    const ILLUSTRATION_COLUMN_INDEX_TR = 13;
    const combineRows = false;
    const mode = 'DetailTR';
    // グルーピング用カテゴリをC16,C17,C18セルから取得
    const groupingColumns = [
      tokaiPromptSheet.getRange("C16").getValue(),
      tokaiPromptSheet.getRange("C17").getValue(),
      tokaiPromptSheet.getRange("C18").getValue()
    ].filter(col => col && col.trim() !== ""); // 空欄を除外
    if (groupingColumns.length === 0) throw new Error('グルーピング用カテゴリが設定されていません（C16〜C18セル）。');
    const baseTitle = "保全_(赤)_カルテ";

    // --- 1. 対象シート取得 ---
    const targetSheetName = tokaiPromptSheet.getRange("C12").getValue();
    if (!targetSheetName) throw new Error(`対象シート名が入力されていません。`);
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`データシート "${targetSheetName}" が見つかりません。`);

    // --- 2. ID採番 ---
    try {
      const masterSheetName = tokaiPromptSheet.getRange("C14").getValue();
      const id_col = 8;
      const ID_PREFIX = "DC-TY-";
      assignPersistentGroupIds_(sheet, masterSheetName, id_col, ID_PREFIX, groupingColumns);
      SpreadsheetApp.getActiveSpreadsheet().toast('グループIDをA列に採番・更新しました。', 'ID採番完了', 3);
    } catch (e) {
      throw new Error(`ID採番中にエラーが発生しました: ${e.message}`);
    }

    // --- 3. データをグループ化 ---
    const { groupedData, allData } = _groupDataByColumns(sheet, groupingColumns);
    if (groupedData.size === 0) throw new Error('グルーピング対象のデータが0件です。');

    const outputFolderUrl = promptSheet.getRange(slideSaveDir_pos).getValue();
    let workSheet;
    const workListData = [];

    if (isSplitMode) {
      // === 分割モード ===
      // サブフォルダを作成
      const { subFolderId, subFolderName } = _createSubfolderForSplitMode(baseTitle, outputFolderUrl);

      // 作業シートを作成（分割モード用）
      workSheet = _createWorkSheetForSplitMode(targetSheetName, subFolderId, true);

      // グループごとにプレゼンテーションを作成し、タスクを登録
      for (const [groupKey, rowNumbers] of groupedData.entries()) {
        const presentationId = _createPresentationForGroup(groupKey, baseTitle, subFolderId);

        rowNumbers.forEach(rowNum => {
          workListData.push([
            `Row_${rowNum}`,
            rowNum,
            STATUS_EMPTY,
            mode,
            presentationId, SLIDES_TEMPLATE_ID_TR, TEMPLATE_SLIDE_INDEX_TR, combineRows,
            JSON.stringify(ALT_TEXT_TITLE_MAP_TR),
            IMAGE_ALT_TEXT_TITLE_TR,
            ILLUSTRATION_COLUMN_INDEX_TR,
            "", // ConditionalBgColors
            groupKey // GroupKey
          ]);
        });
      }

      Logger.log(`分割モード: ${groupedData.size} 個のプレゼンテーションを作成しました。`);

    } else {
      // === 統合モード（従来の動作） ===
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const newPresentationTitle = `${baseTitle}_${timestamp}`;
      const presentationId = _createAndMovePresentation(newPresentationTitle);

      // 作業シートを作成（統合モード用）
      workSheet = _createWorkSheet(presentationId, targetSheetName);

      // 全行をタスクとして登録
      const dataRows = allData.slice(1);
      dataRows.forEach((_, index) => {
        const rowNum = index + 2;
        workListData.push([
          `Row_${rowNum}`,
          rowNum,
          STATUS_EMPTY,
          mode,
          presentationId, SLIDES_TEMPLATE_ID_TR, TEMPLATE_SLIDE_INDEX_TR, combineRows,
          JSON.stringify(ALT_TEXT_TITLE_MAP_TR),
          IMAGE_ALT_TEXT_TITLE_TR,
          ILLUSTRATION_COLUMN_INDEX_TR
        ]);
      });
    }

    if (workListData.length > 0) {
      const numCols = isSplitMode ? 13 : 11;
      workSheet.getRange(2, 1, workListData.length, numCols).setValues(workListData);
    }

    _showSetupCompletionDialog({
      workSheetName: WORK_LIST_SHEET_NAME,
      menuItemName: `🌡️ 東海理化用 > 1-6 スライド生成(詳細情報)（実行）`,
      processFunctionName: 'createSlides_PROCESS',
      useManualExecution: true
    });

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー (DetailTR):\n${e.message}`);
  }
}

/**
 * [SETUP] 複数行1スライド (SummaryTR) のセットアップ - 統合モード
 * すべてのスライドを1つのプレゼンテーションに生成
 */
function createSlideSummaryTR_Combined_SETUP() {
  _createSlideSummaryTR_SETUP_Internal(false); // 統合モード
}

/**
 * [SETUP] 複数行1スライド (SummaryTR) のセットアップ - 分割モード
 * グループごとに別々のプレゼンテーションを生成
 */
function createSlideSummaryTR_Split_SETUP() {
  _createSlideSummaryTR_SETUP_Internal(true); // 分割モード
}

/**
 * [内部] SummaryTR セットアップの共通ロジック
 * @param {boolean} isSplitMode - true: 分割モード, false: 統合モード
 */
function _createSlideSummaryTR_SETUP_Internal(isSplitMode) {
  const ui = SpreadsheetApp.getUi();
  try {
    const modeLabel = isSplitMode ? '分割モード' : '統合モード';
    ss.toast(`セットアップ (SummaryTR - ${modeLabel}) を開始します...`, '開始', 10);

    // --- 元の設定項目 ---
    const SLIDES_TEMPLATE_ID_TR = '1NYkmHwG4hHm8sadB_n15N6knXNGXtX3ZpLibePXfKS8';
    const TEMPLATE_SLIDE_INDEX_TR = 2;
    const ALT_TEXT_TITLE_MAP_TR = {
      "placeholder_equip": 3, "placeholder_line": 6, "placeholder_process": 8,
      "placeholder_trouble": 9, "placeholder_id": 0, "placeholder_place": 1,
      "placeholder_point_rough": 7, "placeholder_equip_num": 5, "placeholder_original_nums": 2,
      "placeholder_date": 4, "placeholder_title": 10, "placeholder_detail": 11,
      "placeholder_issue": 12, "placeholder_fix": 13, "placeholder_name": 14, "placeholder_original_num": 2
    };
    const IMAGE_ALT_TEXT_TITLE_TR = false;
    const ILLUSTRATION_COLUMN_INDEX_TR = false;
    const combineRows = true;
    const mode = 'SummaryTR';
    const chunkSize = 5;
    // グルーピング用カテゴリをC23,C24,C25セルから取得
    const groupingColumns = [
      tokaiPromptSheet.getRange("C23").getValue(),
      tokaiPromptSheet.getRange("C24").getValue(),
      tokaiPromptSheet.getRange("C25").getValue()
    ].filter(col => col && col.trim() !== ""); // 空欄を除外
    if (groupingColumns.length === 0) throw new Error('グルーピング用カテゴリが設定されていません（C23〜C25セル）。');
    const baseTitle = "保全_(青)_事例";

    // --- 1. 対象シート取得 ---
    const targetSheetName = tokaiPromptSheet.getRange("C19").getValue();
    if (!targetSheetName) throw new Error(`対象シート名が入力されていません（C19セル）。`);
    const sheet = ss.getSheetByName(targetSheetName);
    if (!sheet) throw new Error(`データシート "${targetSheetName}" が見つかりません。`);

    // --- 2. ID採番 ---
    try {
      const masterSheetName = tokaiPromptSheet.getRange("C21").getValue();
      const id_col = 1;
      const ID_PREFIX = "EC-TY-";
      assignPersistentGroupIds_(sheet, masterSheetName, id_col, ID_PREFIX, groupingColumns);
      SpreadsheetApp.getActiveSpreadsheet().toast('グループIDをA列に採番・更新しました。', 'ID採番完了', 3);
    } catch (e) {
      throw new Error(`ID採番中にエラーが発生しました: ${e.message}`);
    }

    // --- 3. データをグループ化 ---
    const { groupedData } = _groupDataByColumns(sheet, groupingColumns);
    if (groupedData.size === 0) throw new Error('グルーピング対象のデータが0件です。');

    // SummaryTR専用: C20セルからフォルダURLを取得
    const outputFolderUrl = tokaiPromptSheet.getRange("C20").getValue();
    let workSheet;
    const workListData = [];

    if (isSplitMode) {
      // === 分割モード ===
      // サブフォルダを作成
      const { subFolderId } = _createSubfolderForSplitMode(baseTitle, outputFolderUrl);

      // 作業シートを作成（分割モード用）
      workSheet = _createWorkSheetForSplitMode(targetSheetName, subFolderId, true);

      // グループごとにプレゼンテーションを作成し、タスクを登録
      for (const [groupKey, rowNumbers] of groupedData.entries()) {
        const presentationId = _createPresentationForGroup(groupKey, baseTitle, subFolderId);

        // チャンキング
        for (let i = 0; i < rowNumbers.length; i += chunkSize) {
          const chunkRowNumbers = rowNumbers.slice(i, i + chunkSize);

          workListData.push([
            `${groupKey}|Chunk${i}`,
            JSON.stringify(chunkRowNumbers),
            STATUS_EMPTY,
            mode,
            presentationId, SLIDES_TEMPLATE_ID_TR, TEMPLATE_SLIDE_INDEX_TR, combineRows,
            JSON.stringify(ALT_TEXT_TITLE_MAP_TR),
            IMAGE_ALT_TEXT_TITLE_TR,
            ILLUSTRATION_COLUMN_INDEX_TR,
            "", // ConditionalBgColors
            groupKey // GroupKey
          ]);
        }
      }

      Logger.log(`分割モード: ${groupedData.size} 個のプレゼンテーションを作成しました。`);

    } else {
      // === 統合モード（従来の動作） ===
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      const newPresentationTitle = `${baseTitle}_${timestamp}`;
      const presentationId = _createAndMovePresentation(newPresentationTitle);

      // 作業シートを作成（統合モード用）
      workSheet = _createWorkSheet(presentationId, targetSheetName);

      // グループごとにチャンキングしてタスクを登録
      for (const [groupKey, rowNumbers] of groupedData.entries()) {
        for (let i = 0; i < rowNumbers.length; i += chunkSize) {
          const chunkRowNumbers = rowNumbers.slice(i, i + chunkSize);

          workListData.push([
            `${groupKey}|Chunk${i}`,
            JSON.stringify(chunkRowNumbers),
            STATUS_EMPTY,
            mode,
            presentationId, SLIDES_TEMPLATE_ID_TR, TEMPLATE_SLIDE_INDEX_TR, combineRows,
            JSON.stringify(ALT_TEXT_TITLE_MAP_TR),
            IMAGE_ALT_TEXT_TITLE_TR,
            ILLUSTRATION_COLUMN_INDEX_TR
          ]);
        }
      }
    }

    if (workListData.length > 0) {
      const numCols = isSplitMode ? 13 : 11;
      workSheet.getRange(2, 1, workListData.length, numCols).setValues(workListData);
    }

    _showSetupCompletionDialog({
      workSheetName: WORK_LIST_SHEET_NAME,
      menuItemName: '🌡️ 東海理化用 > 2-2 スライド生成(まとめ一覧)（実行）',
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

  // --- 1. 共通設定を作業シートから取得 ---
  const targetSheetName = workSheet.getRange("E1").getValue();
  const modeFlag = workSheet.getRange("N1").getValue(); // "SPLIT" or "COMBINED"（N列に移動）
  const isSplitMode = (modeFlag === "SPLIT");
  const subFolderId = isSplitMode ? workSheet.getRange("O1").getValue() : null; // 分割モード時のサブフォルダID（O列）

  // 統合モードの場合のみD1からプレゼンテーションIDを取得
  const singlePresentationId = isSplitMode ? null : workSheet.getRange("D1").getValue();

  if (!targetSheetName) {
    Logger.log("作業シート E1 に対象シート名がありません。SETUPを先に実行してください。");
    return;
  }

  // 統合モードの場合、プレゼンテーションIDが必要
  if (!isSplitMode && !singlePresentationId) {
    Logger.log("統合モードですが、D1にプレゼンテーションIDがありません。");
    return;
  }

  let inputSheet;
  let allData;
  // 分割モード用: プレゼンテーションIDごとにキャッシュ
  const presentationCache = new Map();

  try {
    inputSheet = ss.getSheetByName(targetSheetName);
    if (!inputSheet) throw new Error(`入力シート ${targetSheetName} が見つかりません。`);
    allData = inputSheet.getDataRange().getValues(); // ★全データを一度だけ読み込む
  } catch (e) {
    Logger.log(`入力シートが開けません: ${e}`);
    return;
  }

  // --- 2. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 13); // 13列分取得（GroupKey含む）
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
      const taskPresentationId = workValues[i][4]; // E列: 各タスクのプレゼンテーションID
      const templateId = workValues[i][5];
      const templateIndex = workValues[i][6];
      const combineRows = workValues[i][7];
      const altTextMap = JSON.parse(workValues[i][8]);
      const imageAltText = workValues[i][9];
      const imageColIndex = workValues[i][10];
      const conditionalBgColors = workValues[i][11] ? JSON.parse(workValues[i][11]) : null;

      let templateSlide;
      let presentation;

      try {
        // 3a. ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        // プレゼンテーションを取得（キャッシュを使用）
        const presId = isSplitMode ? taskPresentationId : singlePresentationId;
        if (presentationCache.has(presId)) {
          presentation = presentationCache.get(presId);
        } else {
          presentation = SlidesApp.openById(presId);
          presentationCache.set(presId, presentation);
        }

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
            imageColIndex,
            conditionalBgColors
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
      if (isSplitMode) {
        // === 分割モード：全プレゼンテーションの空スライドを削除 ===
        const processedPresentationIds = new Set();
        for (const row of workValues) {
          const presId = row[4]; // E列: PresentationID
          if (presId && !processedPresentationIds.has(presId)) {
            processedPresentationIds.add(presId);
            try {
              const pres = SlidesApp.openById(presId);
              const slides = pres.getSlides();
              if (slides.length > 1) {
                slides[0].remove();
                Logger.log(`プレゼンテーション ${pres.getName()} の最初の空スライドを削除しました。`);
              }
            } catch (e) {
              Logger.log(`警告: プレゼンテーション ${presId} の空スライド削除中にエラー: ${e}`);
            }
          }
        }

        // 完了通知（サブフォルダへのリンク）
        const folderUrl = subFolderId ? `https://drive.google.com/drive/folders/${subFolderId}` : '';
        Logger.log(`処理完了。${processedPresentationIds.size} 個のプレゼンテーションを作成しました。`);
        _showProgress(`すべてのスライド生成が完了しました！(${processedPresentationIds.size}ファイル)`, '✅ 完了', 10);

        // 手動実行時のみアラート表示
        if (_isManualExecution()) {
          ui.alert('成功', `${processedPresentationIds.size} 個のプレゼンテーションを作成しました。\n\nフォルダURL: ${folderUrl}`, ui.ButtonSet.OK);
        }

      } else {
        // === 統合モード：従来の処理 ===
        const finalPresentation = SlidesApp.openById(singlePresentationId);
        const initialSlide = finalPresentation.getSlides()[0];
        if (initialSlide && finalPresentation.getSlides().length > 1) {
          initialSlide.remove();
          Logger.log("最初の空スライドを削除しました。");
        }

        // 完了通知
        const presentationUrl = finalPresentation.getUrl();
        Logger.log(`処理完了。プレゼンテーションURL: ${presentationUrl}`);
        _showProgress('すべてのスライド生成が完了しました！', '✅ 完了', 10);

        // 手動実行時のみアラート表示
        if (_isManualExecution()) {
          ui.alert('成功', `プレゼンテーションを作成しました: ${finalPresentation.getName()}\nURL: ${presentationUrl}`, ui.ButtonSet.OK);
        }
      }

      // 4c. トリガーを停止
      stopTriggers_('createSlides_PROCESS');

    } catch (e) {
      Logger.log(`完了処理（空スライド削除、トリガー停止）中にエラー: ${e}`);
    }
  } else if (remainingTasks > 0) {
    // まだタスクが残っている場合（タイムアウトで中断）
    _showProgress(`${processedCountInThisRun}件処理完了。残り${remainingTasks}件（次回継続）`, '⏸️ 中断', 5);
  } else {
    // 処理タスクがなかった場合（すでに全完了済み）
    _showProgress('処理対象のタスクがありません', '📋 確認', 3);
  }
}

// ===================================================================
// STEP 3: ヘルパー関数 (新規・変更・流用)
// ===================================================================

/**
 * [新規] 1行1スライドの転記処理 (createSlidesMainFunc の
 * * * ブロックから移植)
 * @param {SlidesApp.Presentation} presentation - 書き込み先のプレゼンテーション
 * @param {SlidesApp.Slide} templateSlide - テンプレートスライド
 * @param {Array} row - データ行
 * @param {number} rowNumForLog - ログ用行番号
 * @param {Object} altTextMap - 代替テキストと列インデックスのマッピング
 * @param {string|false} imageAltText - 画像プレースホルダーの代替テキスト（falseの場合は画像処理スキップ）
 * @param {number|false} imageColIndex - 画像データの列インデックス（falseの場合は画像処理スキップ）
 * @param {Object|null} conditionalBgColors - 条件付き背景色設定（nullの場合は背景色処理スキップ）
 *        例: {"placeholder_importance": {"設計技術": "#eb4164", "QCD向上": "#fff2cc"}}
 */
function _transferSingleRowToSlide(presentation, templateSlide, row, rowNumForLog, altTextMap, imageAltText, imageColIndex, conditionalBgColors) {

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

  // --- テキスト置換 & 条件付き背景色設定 ---
  for (const altTextTitle in altTextMap) {
    const colIndex = altTextMap[altTextTitle];
    if (colIndex >= 0 && colIndex < row.length) {
      let replacementValue = row[colIndex];
      if (replacementValue instanceof Date) {
        replacementValue = Utilities.formatDate(replacementValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      }
      const shape = pageElements.find(el => el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.getTitle() === altTextTitle)?.asShape();
      if (shape && shape.getText) {
        const textValue = String(replacementValue || '');
        shape.getText().setText(textValue);

        // --- 条件付き背景色設定 ---
        if (conditionalBgColors && conditionalBgColors[altTextTitle]) {
          const colorMap = conditionalBgColors[altTextTitle];
          if (colorMap[textValue]) {
            const hexColor = colorMap[textValue];
            try {
              shape.getFill().setSolidFill(hexColor);
              Logger.log(`行 ${rowNumForLog}: "${altTextTitle}" の背景色を ${hexColor} に設定しました（値: "${textValue}"）`);
            } catch (colorError) {
              Logger.log(`警告(行 ${rowNumForLog}): "${altTextTitle}" の背景色設定でエラー - ${colorError}`);
            }
          }
        }
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
    "AltTextMap (JSON)", "ImageAltText", "ImageColIndex", "ConditionalBgColors (JSON)"
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
// ===================================================================
// 永続化対応 ID採番関数
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
// 分割モード対応ヘルパー関数
// ===================================================================

/**
 * [新規] データを指定カラムでグループ化する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - データシート
 * @param {string[]} groupingColumns - グルーピングに使用する列名の配列
 * @return {Object} { header, groupedData: Map<groupKey, rowNumbers[]> }
 */
function _groupDataByColumns(sheet, groupingColumns) {
  const allData = sheet.getDataRange().getValues();
  const header = allData[0];
  const dataRows = allData.slice(1);

  const groupIndices = groupingColumns.map(colName => {
    const index = header.indexOf(colName);
    if (index === -1) throw new Error(`データシートのヘッダーに列名「${colName}」が見つかりません。`);
    return index;
  });

  const groupedData = new Map(); // Map<グループキー, rowNumbers[]>
  dataRows.forEach((row, index) => {
    const keyValues = groupIndices.map(idx => row[idx]);
    // グループ化のキーが空欄の場合はスキップ
    if (keyValues.some(val => val === null || val === "")) {
      return;
    }
    const groupKey = keyValues.join('|');
    if (!groupedData.has(groupKey)) {
      groupedData.set(groupKey, []);
    }
    groupedData.get(groupKey).push(index + 2); // 実際のシート行番号
  });

  return { header, groupedData, allData };
}

/**
 * [新規] 分割モード用：グループごとにプレゼンテーションを作成し、サブフォルダに保存
 * @param {string} baseTitle - 基本タイトル（例: "詳細事例スライド"）
 * @param {string} baseFolderUrl - 保存先フォルダのURL
 * @return {Object} { subFolderId, subFolderName }
 */
function _createSubfolderForSplitMode(baseTitle, baseFolderUrl) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  const subFolderName = `${baseTitle}_${timestamp}`;

  let parentFolder = null;
  if (baseFolderUrl) {
    const folderId = _extractFolderIdFromUrl(baseFolderUrl);
    if (folderId) {
      try {
        parentFolder = DriveApp.getFolderById(folderId);
      } catch (e) {
        Logger.log(`警告: 指定されたフォルダが見つかりません。ルートに作成します。`);
      }
    }
  }

  let subFolder;
  if (parentFolder) {
    subFolder = parentFolder.createFolder(subFolderName);
  } else {
    subFolder = DriveApp.getRootFolder().createFolder(subFolderName);
  }

  Logger.log(`サブフォルダを作成しました: ${subFolderName}`);
  return { subFolderId: subFolder.getId(), subFolderName: subFolderName };
}

/**
 * [新規] 分割モード用：グループごとにプレゼンテーションを作成
 * @param {string} groupKey - グループキー
 * @param {string} baseTitle - 基本タイトル
 * @param {string} subFolderId - サブフォルダID
 * @return {string} presentationId
 */
function _createPresentationForGroup(groupKey, baseTitle, subFolderId) {
  // グループキーからファイル名を作成（安全な文字に変換）
  const safeGroupKey = groupKey.replace(/\|/g, '_').replace(/[\/\\:*?"<>|]/g, '_');
  const presentationTitle = `${baseTitle}_${safeGroupKey}`;

  const tempPresentation = SlidesApp.create(presentationTitle);
  const presentationId = tempPresentation.getId();
  const presentationFile = DriveApp.getFileById(presentationId);

  if (subFolderId) {
    try {
      const subFolder = DriveApp.getFolderById(subFolderId);
      presentationFile.moveTo(subFolder);
      Logger.log(`プレゼンテーション「${presentationTitle}」をサブフォルダに移動しました。`);
    } catch (e) {
      Logger.log(`警告: プレゼンテーションの移動に失敗。ルートに残ります。`);
    }
  }

  return presentationId;
}

/**
 * [新規] 分割モード用：作業シートを作成（複数プレゼンテーション対応）
 * @param {string} targetSheetName - 読み込み元のシート名
 * @param {string} subFolderId - サブフォルダID（分割モード時）
 * @param {boolean} isSplitMode - 分割モードかどうか
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成またはクリアされた作業シート
 */
function _createWorkSheetForSplitMode(targetSheetName, subFolderId, isSplitMode) {
  let workSheet = ss.getSheetByName(WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(WORK_LIST_SHEET_NAME, 0);
  }

  const workHeader = [
    "TaskKey", "TaskData (JSON or RowNum)", "Status", "Mode",
    "PresentationID", "TemplateID", "TemplateIndex", "CombineRows",
    "AltTextMap (JSON)", "ImageAltText", "ImageColIndex", "ConditionalBgColors (JSON)",
    "GroupKey", // 分割モード用：どのグループに属するかを記録
    "SplitMode", // N列: 分割モードフラグ ("SPLIT" or "COMBINED")
    "SubFolderId" // O列: サブフォルダID（分割モード時のみ）
  ];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1 に対象シート名を保存（PROCESS時に参照）
  workSheet.getRange("E1").setValue(targetSheetName);
  // N1 に分割モードフラグを保存
  workSheet.getRange("N1").setValue(isSplitMode ? "SPLIT" : "COMBINED");
  // O1 にサブフォルダIDを保存（分割モード時）
  if (isSplitMode && subFolderId) {
    workSheet.getRange("O1").setValue(subFolderId);
  }

  workSheet.setTabColor('#999999');
  workSheet.autoResizeColumn(1);
  return workSheet;
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
 *   G列: CONDITIONAL_BG_COLORS (JSON) - 条件付き背景色設定
 *         例: {"placeholder_importance": {"設計技術": "#eb4164", "QCD向上": "#fff2cc"}}
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

  const data = masterSheet.getRange(2, 1, lastRow - 1, 7).getValues(); // 7列に拡張

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === templateId) {
      // E列・F列が空白の場合はfalseを設定（画像処理をスキップ）
      const imageAltText = data[i][4] !== "" ? data[i][4] : false;
      const imageColIndex = data[i][5] !== "" ? data[i][5] : false;
      // G列が空白の場合はnullを設定（条件付き背景色なし）
      Logger.log(`G列の値: "${data[i][6]}" (型: ${typeof data[i][6]})`);
      const conditionalBgColors = data[i][6] !== "" ? JSON.parse(data[i][6]) : null;
      Logger.log(`conditionalBgColors: ${JSON.stringify(conditionalBgColors)}`);

      return {
        templateId: data[i][0],           // A列: GoogleスライドID
        templateName: data[i][1],         // B列: テンプレート名
        slideIndex: data[i][2],           // C列: スライドIndex
        altTextTitleMap: JSON.parse(data[i][3]), // D列: ALT_TEXT_TITLE_MAP (JSON)
        imageAltText: imageAltText,       // E列: IMAGE_ALT_TEXT（空白ならfalse）
        imageColIndex: imageColIndex,     // F列: IMAGE_COL_INDEX（空白ならfalse）
        conditionalBgColors: conditionalBgColors // G列: CONDITIONAL_BG_COLORS（空白ならnull）
      };
    }
  }

  return null; // 見つからない場合
}


// ===================================================================
// スライド分割機能
// ===================================================================

/**
 * [メイン関数] 既存のスライドをカテゴリ別に分割して出力
 * 「スライド分割」シートから設定を読み込み、代替テキストタイトルに基づいて分割
 */
/**
 * [SETUP] スライド分割のセットアップを行う関数
 * 1. 設定を読み込み、スライドをカテゴリでグループ化
 * 2. 作業リスト（_スライド分割作業リスト）シートを作成
 * 3. 出力フォルダを作成
 */
function splitPresentationByCategory_SETUP() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ss.toast('スライド分割のセットアップを開始します...', '開始', 10);

    // --- 1. 設定読み込み ---
    const configSheet = ss.getSheetByName('スライド分割');
    if (!configSheet) {
      throw new Error('「スライド分割」シートが見つかりません。');
    }

    const sourceSlideId = configSheet.getRange('C7').getValue();
    const category1 = configSheet.getRange('C9').getValue();
    const category2 = configSheet.getRange('C10').getValue();
    const category3 = configSheet.getRange('C11').getValue();

    if (!sourceSlideId) {
      throw new Error('C7セルに分割対象のスライドIDを入力してください。');
    }
    if (!category1 || !category2 || !category3) {
      throw new Error('C9, C10, C11セルに全てのカテゴリ（代替テキストタイトル）を入力してください。');
    }

    const categoryTitles = [category1, category2, category3];
    Logger.log(`カテゴリ設定: ${categoryTitles.join(', ')}`);

    // --- 2. 元スライドを開く ---
    const sourcePresentation = SlidesApp.openById(sourceSlideId);
    const sourceSlides = sourcePresentation.getSlides();
    const sourcePresentationName = sourcePresentation.getName();

    Logger.log(`元スライド: ${sourcePresentationName} (${sourceSlides.length}枚)`);

    // --- 3. 出力フォルダ作成 ---
    const sourceFile = DriveApp.getFileById(sourceSlideId);
    const parentFolders = sourceFile.getParents();
    let parentFolder;
    if (parentFolders.hasNext()) {
      parentFolder = parentFolders.next();
    } else {
      parentFolder = DriveApp.getRootFolder();
    }

    const baseFolderName = `分割版_${sourcePresentationName}`;
    let outputFolderName = baseFolderName;
    let suffix = 1;

    // 既存フォルダがある場合はサフィックスを追加
    while (parentFolder.getFoldersByName(outputFolderName).hasNext()) {
      suffix++;
      outputFolderName = `${baseFolderName}_${suffix}`;
    }

    const outputFolder = parentFolder.createFolder(outputFolderName);
    const outputFolderId = outputFolder.getId();
    Logger.log(`新規フォルダ作成: ${outputFolderName}`);

    // --- 4. スライドをカテゴリでグループ化 ---
    const slideGroups = _groupSlidesByCategory(sourceSlides, categoryTitles);

    if (Object.keys(slideGroups).length === 0) {
      throw new Error('カテゴリに該当するスライドが見つかりませんでした。');
    }

    Logger.log(`グループ数: ${Object.keys(slideGroups).length}`);

    // --- 5. 作業シート作成 & タスク書き込み ---
    let workSheet = ss.getSheetByName(SLIDE_SPLIT_WORK_LIST_SHEET_NAME);
    if (workSheet) {
      workSheet.clear();
    } else {
      workSheet = ss.insertSheet(SLIDE_SPLIT_WORK_LIST_SHEET_NAME, 0);
    }

    const workHeader = ["CategoryKey", "SlideIndices (JSON)", "Status"];
    workSheet.getRange(1, 1, 1, 3).setValues([workHeader]).setFontWeight('bold');

    // E1, F1, G1 に実行時に必要な情報を保存
    workSheet.getRange("E1").setValue(sourceSlideId);           // 元スライドID
    workSheet.getRange("F1").setValue(outputFolderId);          // 出力フォルダID
    workSheet.getRange("G1").setValue('保全_(赤)_カルテ');       // ベースファイル名

    const workListData = [];
    for (const categoryKey in slideGroups) {
      workListData.push([
        categoryKey,
        JSON.stringify(slideGroups[categoryKey]), // スライドインデックスの配列をJSON文字列として保存
        STATUS_EMPTY
      ]);
    }

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 3).setValues(workListData);
      workSheet.autoResizeColumns(1, 3);
    }

    // タブの色をグレーに設定
    workSheet.setTabColor('#999999');

    ss.toast('セットアップが完了しました。', '完了', 5);
    _showSetupCompletionDialog({
      workSheetName: SLIDE_SPLIT_WORK_LIST_SHEET_NAME,
      menuItemName: '🌡️ 東海理化用 > 3-2 スライド分割（実行）',
      processFunctionName: 'splitPresentationByCategory_PROCESS',
      useManualExecution: true
    });

  } catch (e) {
    ss.toast('', '', 1);
    Logger.log(`エラー: ${e.message}\n${e.stack}`);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] スライド分割のバッチ処理を行うワーカー関数
 * 1. _スライド分割作業リスト シートから「未処理」のタスクを取得
 * 2. 時間の許す限りスライド分割を実行
 */
function splitPresentationByCategory_PROCESS() {
  const startTime = new Date().getTime();
  const taskExecutionTimes = [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workSheet = ss.getSheetByName(SLIDE_SPLIT_WORK_LIST_SHEET_NAME);

  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  _showProgress('スライド分割処理を開始します...', '📑 スライド分割', 3);

  // --- 1. 共通設定を作業シートから取得 ---
  const sourceSlideId = workSheet.getRange("E1").getValue();
  const outputFolderId = workSheet.getRange("F1").getValue();
  const baseFileName = workSheet.getRange("G1").getValue();

  if (!sourceSlideId || !outputFolderId) {
    Logger.log("作業シート E1 または F1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  let sourcePresentation;
  let outputFolder;

  try {
    sourcePresentation = SlidesApp.openById(sourceSlideId);
    outputFolder = DriveApp.getFolderById(outputFolderId);
  } catch (e) {
    Logger.log(`必須リソースが開けません: ${e}`);
    return;
  }

  // --- 2. 未処理のタスクを検索 ---
  const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 3);
  const workValues = workRange.getValues();

  let processedCountInThisRun = 0;

  // --- 3. バッチ処理ループ ---
  for (let i = 0; i < workValues.length; i++) {
    const currentStatus = workValues[i][2]; // C列: Status

    if (currentStatus === STATUS_EMPTY) {
      // 動的タイムアウトチェック
      if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
        Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
        break;
      }

      const taskStartTime = new Date().getTime();
      const sheetRow = i + 2; // 作業シートの行番号
      const categoryKey = workValues[i][0];
      const slideIndices = JSON.parse(workValues[i][1]);

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        Logger.log(`[${processedCountInThisRun + 1}] ${categoryKey} を作成中... (${slideIndices.length}枚)`);

        _createSplitPresentation(
          sourcePresentation,
          slideIndices,
          categoryKey,
          outputFolder,
          baseFileName
        );

        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;

        // このタスクの実行時間を記録
        const taskEndTime = new Date().getTime();
        const taskDuration = taskEndTime - taskStartTime;
        taskExecutionTimes.push(taskDuration);
        Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

        // 5件ごとに進捗を表示
        if (processedCountInThisRun % 5 === 0) {
          const totalTasks = workValues.length;
          _showProgress(
            `${processedCountInThisRun} / ${totalTasks} 件完了`,
            '📑 スライド分割中',
            2
          );
        }

        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク "${categoryKey}" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
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
    _showProgress(
      `すべてのスライド分割が完了しました！（合計 ${processedCountInThisRun} 件）`,
      '✅ 完了',
      10
    );
  } else {
    _showProgress(
      `今回 ${processedCountInThisRun} 件処理。残り ${remainingTasks} 件`,
      '⏸️ 一時停止',
      5
    );
  }
}

/**
 * [後方互換] 旧関数名のエイリアス（直接実行用）
 * ※ SETUP/PROCESS方式を推奨
 */
function splitPresentationByCategory() {
  splitPresentationByCategory_SETUP();
}

/**
 * [ヘルパー] スライドから指定された代替テキストタイトルを持つテキストボックスの内容を取得
 * @param {GoogleAppsScript.Slides.Slide} slide - 対象スライド
 * @param {string[]} altTextTitles - 検索する代替テキストタイトルの配列
 * @return {Object} - 代替テキストタイトルをキー、テキスト内容を値とするオブジェクト
 */
function _getSlideTextByAltTitle(slide, altTextTitles) {
  const result = {};

  // スライド内の全ページ要素を取得
  const pageElements = slide.getPageElements();

  for (const element of pageElements) {
    // 代替テキストのタイトルを取得
    const altTitle = element.getTitle();

    if (altTitle && altTextTitles.includes(altTitle)) {
      // Shape（図形/テキストボックス）の場合のみテキストを取得
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const shape = element.asShape();
        const textRange = shape.getText();
        if (textRange) {
          const text = textRange.asString().trim();
          result[altTitle] = text;
        }
      }
    }
  }

  return result;
}

/**
 * [ヘルパー] スライドをテキストボックスの内容でグループ化
 * @param {GoogleAppsScript.Slides.Slide[]} slides - スライドの配列
 * @param {string[]} altTextTitles - 検索する代替テキストタイトルの配列
 * @return {Object} - グループキーをキー、スライドインデックス配列を値とするオブジェクト
 */
function _groupSlidesByCategory(slides, altTextTitles) {
  const groups = {};

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    const textsByAltTitle = _getSlideTextByAltTitle(slide, altTextTitles);

    // 3つ全ての代替テキストタイトルに対応するテキストが見つからない場合は除外
    const foundCount = Object.keys(textsByAltTitle).length;
    if (foundCount === 0) {
      Logger.log(`スライド ${i + 1}: 対象テキストボックスなし（除外）`);
      continue;
    }

    // altTextTitlesの順序でテキスト内容を結合してグループキーを作成
    const keyParts = [];
    for (const altTitle of altTextTitles) {
      if (textsByAltTitle[altTitle]) {
        // ファイル名に使えない文字を置換
        const sanitizedText = textsByAltTitle[altTitle]
          .replace(/[\\/:\*\?"<>\|]/g, '_')  // ファイル名禁止文字
          .replace(/\n/g, ' ')                // 改行
          .trim();
        keyParts.push(sanitizedText);
      }
    }

    if (keyParts.length === 0) {
      Logger.log(`スライド ${i + 1}: テキスト内容が空（除外）`);
      continue;
    }

    const categoryKey = keyParts.join('_');

    Logger.log(`スライド ${i + 1}: ${categoryKey}`);

    if (!groups[categoryKey]) {
      groups[categoryKey] = [];
    }
    groups[categoryKey].push(i);
  }

  return groups;
}

/**
 * [ヘルパー] 分割されたプレゼンテーションを作成
 * @param {GoogleAppsScript.Slides.Presentation} sourcePresentation - 元のプレゼンテーション
 * @param {number[]} slideIndices - コピーするスライドのインデックス配列
 * @param {string} categoryKey - カテゴリキー（ファイル名に使用）
 * @param {GoogleAppsScript.Drive.Folder} outputFolder - 出力先フォルダ
 * @param {string} baseFileName - ベースファイル名
 */
function _createSplitPresentation(sourcePresentation, slideIndices, categoryKey, outputFolder, baseFileName) {
  // ファイル名作成
  const fileName = `${baseFileName}_${categoryKey}`;

  // 新規プレゼンテーション作成
  const newPresentation = SlidesApp.create(fileName);
  const newPresentationId = newPresentation.getId();

  // 元スライドから指定されたスライドをコピー
  const sourceSlides = sourcePresentation.getSlides();

  for (const index of slideIndices) {
    const sourceSlide = sourceSlides[index];
    newPresentation.appendSlide(sourceSlide);
  }

  // デフォルトの空スライドを削除（最初のスライド）
  const newSlides = newPresentation.getSlides();
  if (newSlides.length > slideIndices.length) {
    newSlides[0].remove();
  }

  // 保存して閉じる
  newPresentation.saveAndClose();

  // 出力フォルダに移動
  const newFile = DriveApp.getFileById(newPresentationId);
  newFile.moveTo(outputFolder);

  Logger.log(`作成完了: ${fileName} (${slideIndices.length}枚)`);
}
