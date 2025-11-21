// ===================================================================
// Free Prompt バッチ処理
// free promptシートの設定に基づいてデータを処理する
// ===================================================================

/**
 * [STEP 1: 手動実行] Free Prompt実行の「セットアップ」を行う関数
 * 1. データを読み込み、分割単位でチャンク化する
 * 2. 作業リスト（_Free Prompt作業リスト）シートを作成する
 * 3. 結果出力シート（Free Prompt_結果）を作成する
 */
function freePrompt_SETUP() {
  const ui = SpreadsheetApp.getUi();
  try {
    ss.toast('Free Promptのセットアップを開始します...', '開始', 10);

    // --- 1. 設定情報を取得 ---
    if (!freePromptSheet) {
      throw new Error('設定シート「free prompt」が見つかりません。');
    }

    const inputSheetName = freePromptSheet.getRange('C6').getValue();
    const basePrompt = freePromptSheet.getRange('C25').getValue();

    if (!sep || isNaN(sep) || !inputSheetName || !basePrompt) {
      throw new Error('configシート(C4)またはfree promptシート(C6, C25)の設定が不足しています。');
    }

    // --- 2. 入力データを読み込む ---
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) {
      throw new Error(`データシート「${inputSheetName}」が見つかりません。`);
    }

    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];
    const dataRows = allData.slice(1);

    if (dataRows.length === 0) {
      throw new Error(`${inputSheetName}シートにデータがありません。`);
    }

    // --- 3. データをチャンク化してタスクリストを作成 ---
    const workListData = [];
    for (let i = 0; i < dataRows.length; i += sep) {
      const chunkEndIndex = Math.min(i + sep, dataRows.length);
      const taskData = {
        startIndex: i,
        endIndex: chunkEndIndex,
        chunkSize: chunkEndIndex - i
      };

      workListData.push([
        `Chunk_${i}-${chunkEndIndex}`, // TaskKey
        JSON.stringify(taskData), // TaskData (JSON文字列)
        STATUS_EMPTY // Status
      ]);
    }

    if (workListData.length === 0) {
      throw new Error('作成されたタスクが0件です。');
    }

    // --- 4. 作業リスト（_Free Prompt作業リスト）シートを作成 ---
    let workSheet = ss.getSheetByName(FREE_PROMPT_WORK_LIST_SHEET_NAME);
    if (workSheet) {
      workSheet.clear();
    } else {
      workSheet = ss.insertSheet(FREE_PROMPT_WORK_LIST_SHEET_NAME, 0);
    }

    const workHeader = ["TaskKey", "TaskData (JSON)", "Status"];
    workSheet.getRange(1, 1, 1, 3).setValues([workHeader]).setFontWeight('bold');

    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 3).setValues(workListData);
      workSheet.autoResizeColumns(1, 3);
    }

    // タブの色をグレーに設定
    workSheet.setTabColor('#999999');

    const resultSheetName = `分析結果_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;

    // D1セルに結果シート名をメモ
    workSheet.getRange("D1").setValue(resultSheetName);
    Logger.log(`作業シートのD1セルに結果シート名「${resultSheetName}」を書き込みました。`);

    // --- 5. 結果出力シート（Free Prompt_結果）を作成 ---
    let outputSheet = ss.getSheetByName(FREE_PROMPT_OUTPUT_SHEET_NAME);
    if (outputSheet) {
      outputSheet.clear();
    } else {
      outputSheet = ss.insertSheet(FREE_PROMPT_OUTPUT_SHEET_NAME, ss.getNumSheets() + 1);
    }
    outputSheet.getRange("A1").setValue("処理待機中...").setFontStyle('italic');

    ss.toast('セットアップが完了しました。', '完了', 5);
    _showSetupCompletionDialog();

  } catch (e) {
    Logger.log(e);
    ss.toast('セットアップ中にエラーが発生しました。', '失敗', 10);
    ui.alert('セットアップエラー:\n' + e.message, ui.ButtonSet.OK);
  }
}


/**
 * [STEP 2: トリガー実行] Free Promptの「バッチ処理」を行うワーカー関数
 * 1. _Free Prompt作業リスト シートから「未処理」のタスクを取得
 * 2. 時間の許す限りAPI処理を実行
 * 3. 処理結果を Free Prompt_結果 シートに追記
 */
function freePrompt_PROCESS() {
  const startTime = new Date().getTime();
  const taskExecutionTimes = [];

  try {
    // --- 1. 必要なシートと設定を取得 ---
    const workSheet = ss.getSheetByName(FREE_PROMPT_WORK_LIST_SHEET_NAME);
    const outputSheet = ss.getSheetByName(FREE_PROMPT_OUTPUT_SHEET_NAME);

    if (!workSheet || !outputSheet || !freePromptSheet) {
      Logger.log("必要なシート（_Free Prompt作業リスト, Free Prompt_結果, free prompt）がありません。処理を終了します。");
      return;
    }

    _showProgress('Free Prompt処理を開始します...', '📝 Free Prompt実行', 3);

    const inputSheetName = freePromptSheet.getRange('C6').getValue();
    const basePrompt = freePromptSheet.getRange('C25').getValue();
    const inputSheet = ss.getSheetByName(inputSheetName);

    if (!inputSheet) {
      Logger.log(`入力シート「${inputSheetName}」が見つかりません。`);
      return;
    }

    // 元データをすべて読み込む
    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];
    const dataRows = allData.slice(1);

    // --- 2. 未処理のタスクを検索 ---
    const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 3);
    const workValues = workRange.getValues();

    let processedCountInThisRun = 0;
    let isFirstOutput = (outputSheet.getLastRow() <= 1);
    let previousResultJsonForPrompt = "";

    // --- 3. バッチ処理ループ ---
    for (let i = 0; i < workValues.length; i++) {
      const currentStatus = workValues[i][2]; // ステータス列

      // 未処理のタスクか？
      if (currentStatus === STATUS_EMPTY) {

        // 動的タイムアウトチェック
        if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
          Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
          return;
        }

        const taskStartTime = new Date().getTime();
        const sheetRow = i + 2;
        const taskKey = workValues[i][0];
        const taskData = JSON.parse(workValues[i][1]);

        try {
          // 3a. ステータスを「処理中」に更新
          workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);

          // 3b. チャンクデータを取得
          const chunk = dataRows.slice(taskData.startIndex, taskData.endIndex);
          const chunkWithHeader = [header].concat(chunk);
          const csvChunk = chunkWithHeader.map(row =>
            row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
          ).join('\n');

          // 3c. プロンプトを構築
          let prompt = basePrompt;
          if (previousResultJsonForPrompt) {
            prompt += `\n\n# 前回までの出力結果の概要\n以下は前回までに出力した結果です。この内容や形式を参考に、一貫性を保ってください。\n${previousResultJsonForPrompt}`;
          }
          prompt += `\n\n# 今回分析するデータ (CSV形式)\n---\n${csvChunk}`;

          // 3d. APIを呼び出し
          const resultText = callGemini_(prompt);

          // 3e. 結果を解析
          let jsonToParse = "";

          try {
            // 戦略1: ```json ... ``` のマークダウンブロックを探す
            const jsonStringMatch = resultText.match(/```json\s*([\s\S]*?)\s*```/);

            if (jsonStringMatch && jsonStringMatch[1]) {
              jsonToParse = jsonStringMatch[1];
            } else {
              // 戦略2: { または [ で始まる最初のJSON部分を探す
              const startIndex = resultText.indexOf('{');
              const arrayStartIndex = resultText.indexOf('[');

              let jsonStartIndex = -1;

              if (startIndex !== -1 && arrayStartIndex !== -1) {
                jsonStartIndex = Math.min(startIndex, arrayStartIndex);
              } else if (startIndex !== -1) {
                jsonStartIndex = startIndex;
              } else if (arrayStartIndex !== -1) {
                jsonStartIndex = arrayStartIndex;
              }

              if (jsonStartIndex !== -1) {
                const startChar = resultText[jsonStartIndex];
                const endChar = (startChar === '{') ? '}' : ']';
                const jsonEndIndex = resultText.lastIndexOf(endChar);

                if (jsonEndIndex > jsonStartIndex) {
                  jsonToParse = resultText.substring(jsonStartIndex, jsonEndIndex + 1);
                } else {
                  jsonToParse = resultText.substring(jsonStartIndex);
                }
              } else {
                jsonToParse = resultText;
              }
            }

            // 抽出した文字列を解析
            let newResults = [];
            if (jsonToParse.trim() !== "") {
              const parsedJson = JSON.parse(jsonToParse);
              newResults = Array.isArray(parsedJson) ? parsedJson : [parsedJson];
            }

            if (newResults.length === 0) {
              throw new Error("APIから有効なJSONが返されませんでした。");
            }

            // 3f. 結果を出力シートに「追記」
            const outputHeader = Object.keys(newResults[0]);
            const outputData = newResults.map(item => {
              return outputHeader.map(key => item[key] || "");
            });

            if (isFirstOutput) {
              // 初回書き込み時のみヘッダーを書き込む
              outputSheet.clear();
              outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
              isFirstOutput = false;
            }

            // 最終行に追記
            const lastRow = outputSheet.getLastRow();
            outputSheet.getRange(lastRow + 1, 1, outputData.length, outputData[0].length)
              .setValues(outputData)
              .setWrap(true)
              .setVerticalAlignment('top');

            // 次回のプロンプトのため、最新の5件を概要として保存
            const currentLastRow = outputSheet.getLastRow();
            const recentCount = Math.min(5, currentLastRow - 1);
            if (recentCount > 0) {
              const recentRange = outputSheet.getRange(currentLastRow - recentCount + 1, 1, recentCount, outputHeader.length);
              const recentValues = recentRange.getValues();
              const recentObjects = recentValues.map(row => {
                const obj = {};
                outputHeader.forEach((key, idx) => {
                  obj[key] = row[idx];
                });
                return obj;
              });
              previousResultJsonForPrompt = JSON.stringify(recentObjects, null, 2);
            }

            // 3g. ステータスを「完了」に更新
            workSheet.getRange(sheetRow, 3).setValue(STATUS_DONE);
            processedCountInThisRun++;

            // このタスクの実行時間を記録
            const taskEndTime = new Date().getTime();
            const taskDuration = taskEndTime - taskStartTime;
            taskExecutionTimes.push(taskDuration);
            Logger.log(`  タスク実行時間: ${(taskDuration / 1000).toFixed(2)}秒`);

            // 進捗を表示
            if (processedCountInThisRun % 3 === 0) {
              const totalTasks = workValues.length;
              _showProgress(
                `${processedCountInThisRun} / ${totalTasks} 件完了`,
                '📝 Free Prompt実行中',
                2
              );
            }

          } catch (parseError) {
            // JSON解析エラー
            throw new Error(`JSON解析エラー: ${parseError.message}`);
          }

        } catch (e) {
          // 3h. エラー処理
          Logger.log(`タスク \"${taskKey}\" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
          workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

          const taskEndTime = new Date().getTime();
          const taskDuration = taskEndTime - taskStartTime;
          taskExecutionTimes.push(taskDuration);
        }
      }
    }

    Logger.log(`今回の実行で ${processedCountInThisRun} 件のタスクを処理しました。`);

    // シートへの書き込みを強制的に反映
    SpreadsheetApp.flush();

    // 完了チェック：最新のステータスを再取得
    const lastRow = workSheet.getLastRow();
    let remainingTasks = 0;

    if (lastRow >= 2) {
      const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
      remainingTasks = newStatusValues.filter(
        row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
      ).length;
    }

    // すべて完了した場合
    if (remainingTasks === 0 && processedCountInThisRun > 0) {
      // 結果シート名を変更
      const newSheetName = workSheet.getRange("D1").getValue();
      if (!newSheetName) {
        Logger.log("エラー: _Free Prompt作業リスト シートのD1セルに結果シート名がありません。");
        return;
      }

      try {
        outputSheet.setName(newSheetName);
        Logger.log(`シート名を「${newSheetName}」に変更しました。`);
      } catch (e) {
        Logger.log(`シート名変更中にエラー: ${e}`);
      }

      Logger.log("すべてのタスクの処理が完了しました。");
      _showProgress('すべてのFree Prompt処理が完了しました！', '✅ 完了', 10);

      // トリガーを自動停止
      stopFreePromptTriggers_();
    }

  } catch (e) {
    Logger.log(`バッチ処理ワーカーで致命的なエラーが発生しました: ${e}`);
  }
}


/**
 * 'freePrompt_PROCESS' を実行するトリガーを自動停止する関数
 */
function stopFreePromptTriggers_() {
  stopTriggers_('freePrompt_PROCESS');
}
