// ===================================================================
// 画像生成処理: バッチ処理用の関数群
// ===================================================================

// 画像生成処理用の作業シート名
const IMAGE_WORK_LIST_SHEET_NAME = "_画像生成作業リスト";

/**
 * [SETUP] 行ごとの画像生成のセットアップ
 * 「画像生成」シートの設定に基づいて、画像生成タスクを作成します
 *
 * 設定:
 * - C6セル: 画像生成対象のシート名
 * - C7セル: 画像保存先フォルダURL（オプション）
 * - C8セル: 処理対象の通し番号（例: "1-5, 10, 15-20"）
 * - C31セル: 画像生成用のベースプロンプト
 */
function generateRowImages_SETUP() {
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('画像生成のセットアップを開始します...', '開始', 10);

    // --- 1. 設定情報を取得 ---
    const imageGenSheet = ss.getSheetByName('画像生成');
    if (!imageGenSheet) {
      throw new Error('シート「画像生成」が見つかりません。');
    }

    const targetSheetName = imageGenSheet.getRange('C6').getValue();
    if (!targetSheetName) {
      throw new Error('C6セルに画像生成対象のシート名が設定されていません。');
    }

    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      throw new Error(`画像生成対象シート「${targetSheetName}」が見つかりません。`);
    }

    const outputFolderUrl = imageGenSheet.getRange('C7').getValue();
    const targetNumbersString = imageGenSheet.getRange('C8').getValue();
    const basePrompt = imageGenSheet.getRange('C31').getValue();

    if (!basePrompt) {
      throw new Error('C31セルに画像生成用のプロンプトが設定されていません。');
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

    // 通し番号の列インデックスを特定（0列目と仮定）
    const serialNumberColIndex = 0;

    // --- 3. 処理対象の行を特定 ---
    let targetRows = [];
    if (targetNumbersString) {
      // C8セルに指定がある場合、その番号のみを対象とする
      const targetNumbers = new Set(_parseNumberRangeString(String(targetNumbersString)));
      dataRows.forEach((row, index) => {
        const serialNumber = parseInt(row[serialNumberColIndex], 10);
        if (targetNumbers.has(serialNumber)) {
          targetRows.push({
            rowIndex: index + 2, // シート上の行番号（1-indexed）
            serialNumber: serialNumber
          });
        }
      });
    } else {
      // C8セルが空の場合、全行を対象とする
      targetRows = dataRows.map((row, index) => ({
        rowIndex: index + 2,
        serialNumber: parseInt(row[serialNumberColIndex], 10)
      }));
    }

    if (targetRows.length === 0) {
      throw new Error('処理対象の行が見つかりませんでした。C8セルの指定を確認してください。');
    }

    // --- 4. 作業シート作成 & タスク書き込み ---
    const workSheet = _createImageWorkSheet(targetSheetName, outputFolderUrl, basePrompt);
    const workListData = [];

    targetRows.forEach(item => {
      workListData.push([
        `Row_${item.rowIndex}`, // TaskKey
        item.rowIndex, // TaskData (行番号)
        STATUS_EMPTY, // Status
        item.serialNumber // 通し番号（参照用）
      ]);
    });

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 4).setValues(workListData);
    }

    // --- 5. 画像列のヘッダー追加はPROCESS時に行う ---
    // （毎回新しい列に画像を追加する可能性があるため、ここでは追加しない）

    // 完了メッセージ
    ui.alert(
      '✅ セットアップ完了',
      `画像生成タスクを ${targetRows.length} 件作成しました。\n\n次に「🎨 行ごとの画像生成 (実行)」を実行してください。`,
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log(e);
    ui.alert(`セットアップエラー:\n${e.message}`);
  }
}

/**
 * [PROCESS] 画像生成バッチ処理ワーカー
 * この関数を繰り返し実行して、タスクを順次処理します
 */
function generateRowImages_PROCESS() {
  const startTime = new Date().getTime();

  const workSheet = ss.getSheetByName(IMAGE_WORK_LIST_SHEET_NAME);
  if (!workSheet || workSheet.getLastRow() < 2) {
    Logger.log("作業シートが見つからないか、タスクがありません。処理を終了します。");
    return;
  }

  // --- 1. 共通設定を作業シートから取得 ---
  const targetSheetName = workSheet.getRange("E1").getValue();
  const outputFolderUrl = workSheet.getRange("F1").getValue();
  const basePrompt = workSheet.getRange("G1").getValue();

  if (!targetSheetName || !basePrompt) {
    Logger.log("作業シート E1 または G1 に設定情報がありません。SETUPを先に実行してください。");
    return;
  }

  let targetSheet;
  let allData;
  let header;
  let outputFolder = null;

  try {
    targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) throw new Error(`入力シート ${targetSheetName} が見つかりません。`);
    allData = targetSheet.getDataRange().getValues();
    header = allData[0];

    // フォルダの取得（オプション）
    if (outputFolderUrl) {
      const folderId = _extractFolderIdFromUrl(outputFolderUrl);
      if (folderId) {
        try {
          outputFolder = DriveApp.getFolderById(folderId);
        } catch (e) {
          Logger.log(`警告: 指定されたフォルダにアクセスできません。`);
        }
      }
    }
  } catch (e) {
    Logger.log(`必須リソースが開けません: ${e}`);
    return;
  }

  // 画像を挿入する列（最終列の次）
  const imageColumnIndex = targetSheet.getLastColumn() + 1;

  // ヘッダー行に「生成画像」を追加（まだ空の場合のみ）
  const existingHeader = targetSheet.getRange(1, imageColumnIndex).getValue();
  if (!existingHeader) {
    targetSheet.getRange(1, imageColumnIndex).setValue('生成画像').setFontWeight('bold');
  }

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

      const sheetRow = i + 2; // 作業シートの行番号
      const taskKey = workValues[i][0];
      const rowIndex = workValues[i][1]; // 対象シートの行番号
      const serialNumber = workValues[i][3];

      try {
        // ステータスを「処理中」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

        // 対象行のデータを取得
        const row = allData[rowIndex - 1];

        // 行データをCSV形式に変換
        const rowCsvString = row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',');
        const rowWithHeaderCsv = header.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',') + '\n' + rowCsvString;

        // プロンプトを構築
        const finalPrompt = `${basePrompt}

# 入力データ（CSV形式）
以下のデータを基に画像を生成してください。
---
${rowWithHeaderCsv}
---`;

        Logger.log(`[${processedCountInThisRun + 1}] 行${rowIndex}（通し番号: ${serialNumber}）の画像を生成中...`);

        // 画像生成APIを呼び出し
        const base64Image = callGPTApi_(finalPrompt);

        // (1) Google Driveに保存（フォルダが指定されている場合）
        if (outputFolder) {
          try {
            const imageName = `${targetSheetName}_No${serialNumber}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMddHHmmss')}.png`;
            const decodedBytes = Utilities.base64Decode(base64Image);
            const imageBlob = Utilities.newBlob(decodedBytes, 'image/png', imageName);
            const savedFile = outputFolder.createFile(imageBlob);
            Logger.log(`画像を保存: ${savedFile.getName()}`);
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

        // 待機（API制限対策）
        Utilities.sleep(1000);

        // ステータスを「完了」に更新
        workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
        processedCountInThisRun++;
        SpreadsheetApp.flush();

      } catch (e) {
        Logger.log(`タスク "${taskKey}" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
        workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);
        // エラーが発生してもシートには「生成失敗」と表示
        try {
          targetSheet.getRange(rowIndex, imageColumnIndex).setValue('生成失敗');
        } catch (e2) {
          Logger.log(`エラー表示の書き込みに失敗: ${e2.message}`);
        }
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
 * [ヘルパー関数] 画像生成用の作業シートを作成
 */
function _createImageWorkSheet(targetSheetName, outputFolderUrl, basePrompt) {
  let workSheet = ss.getSheetByName(IMAGE_WORK_LIST_SHEET_NAME);
  if (workSheet) {
    workSheet.clear();
  } else {
    workSheet = ss.insertSheet(IMAGE_WORK_LIST_SHEET_NAME, 0);
  }

  const workHeader = ["TaskKey", "RowIndex", "Status", "SerialNumber"];
  workSheet.getRange(1, 1, 1, workHeader.length).setValues([workHeader]).setFontWeight('bold');

  // E1, F1, G1 に実行時に必要な情報を保存
  workSheet.getRange("E1").setValue(targetSheetName);
  workSheet.getRange("F1").setValue(outputFolderUrl || "");
  workSheet.getRange("G1").setValue(basePrompt);

  workSheet.autoResizeColumn(1);
  return workSheet;
}

// ===================================================================
// 注: 以下の共通ヘルパー関数は commonHelpers.js に移動しました
// - _parseNumberRangeString()
// - _extractFolderIdFromUrl()
// ===================================================================
