
/**
 * [STEP 1: 手動実行] 保全ナレッジ生成の「セットアップ」を行う関数
 * 1. データを読み込み、グループ化する
 * 2. 作業リスト（_詳細スライド生成作業リスト）シートを作成する
 * 3. 結果出力シート（保全ナレッジ_結果）を作成する
 */
function generateKnowledge_SETUP() {
  const ui = SpreadsheetApp.getUi();
  try {
    ss.toast('ナレッジ生成のセットアップを開始します...', '開始', 10);

    // --- 1. 設定情報を取得 (元のコードと同じ) ---
    const knowledgeConfigSheet = ss.getSheetByName('カテゴリごとに知見作成');
    if (!knowledgeConfigSheet) {
      throw new Error('設定シート「カテゴリごとに知見作成」が見つかりません。');
    }
    const inputSheetName = knowledgeConfigSheet.getRange('C6').getValue();
    const targetColumns = knowledgeConfigSheet.getRange('C7:C11').getValues()
                            .flat().filter(String);
    if (targetColumns.length === 0) throw new Error('C7:C11に分析対象列がありません。');
    
    // --- 2. 入力データを読み込む (元のコードと同じ) ---
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) throw new Error(`入力シート「${inputSheetName}」が見つかりません。`);

    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];
    const dataRows = allData.slice(1);
    if (dataRows.length === 0) throw new Error('入力シートにデータがありません。');

    // --- 3. 指定された列のインデックスを特定 (元のコードと同じ) ---
    const targetIndices = targetColumns.map(colName => {
      const index = header.indexOf(colName);
      if (index === -1) throw new Error(`列名「${colName}」が見つかりません。`);
      return index;
    });

    // --- 4. データをグループ化し、"行番号" を記録する ---
    const groupedData = new Map(); // Map<グループキー, [行番号の配列]>
    
    dataRows.forEach((row, rowIndex) => { // rowIndex (0から始まる) に注意
      const groupKey = targetIndices.map(index => row[index]).join('|');
      
      if (!groupedData.has(groupKey)) {
        groupedData.set(groupKey, []);
      }
      // allData[0] がヘッダーなので、データ行の実際のシート行番号は (rowIndex + 2)
      // dataRows のインデックスは rowIndex
      groupedData.get(groupKey).push(rowIndex + 2); // 実際のシート行番号を格納
    });

    if (groupedData.size === 0) {
      throw new Error('作成されたグループが0件です。');
    }

    // --- 5. 作業リスト（_詳細スライド生成作業リスト）シートを作成 ---
    let workSheet = ss.getSheetByName(WORK_LIST_SHEET_NAME);
    if (workSheet) {
      workSheet.clear(); // 既存のシートをクリア
    } else {
      workSheet = ss.insertSheet(WORK_LIST_SHEET_NAME, 0);
    }
    
    const workHeader = ["GroupKey", "TargetRowNumbers (JSON)", "Status"];
    workSheet.getRange(1, 1, 1, 3).setValues([workHeader]).setFontWeight('bold');
    
    const workListData = [];
    for (const [groupKey, rowNumbers] of groupedData.entries()) {
      workListData.push([
        groupKey,
        JSON.stringify(rowNumbers), // 行番号の配列をJSON文字列として保存
        STATUS_EMPTY // 初期ステータスは空
      ]);
    }
    
    if (workListData.length > 0) {
      workSheet.getRange(2, 1, workListData.length, 3).setValues(workListData);
      workSheet.autoResizeColumns(1, 3);
    }

    // タブの色をグレーに設定
    workSheet.setTabColor('#999999');

    const resultSheetName = `保全ナレッジ_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;

    // 「_詳細スライド生成作業リスト」シートのD1セルに、今回使うシート名をメモとして書き込む
    workSheet.getRange("D1").setValue(resultSheetName);
    Logger.log(`作業シートのD1セルに結果シート名「${resultSheetName}」を書き込みました。`);

    // --- 6. 結果出力シート（保全ナレッジ_結果）を作成 ---
    let outputSheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
    if (outputSheet) {
      outputSheet.clear(); // 既存のシートをクリア
    } else {
      outputSheet = ss.insertSheet(OUTPUT_SHEET_NAME, ss.getNumSheets() + 1);
    }
    // ヘッダーは PROCESS 側で初回書き込み時に動的に設定する
    outputSheet.getRange("A1").setValue("処理待機中...").setFontStyle('italic');

    ss.toast('セットアップが完了しました。', '完了', 5);
    ui.alert('セットアップ完了', `作業リスト（${WORK_LIST_SHEET_NAME}）を作成しました。\n\n次に、このスクリプトの「generateKnowledge_PROCESS」関数に対して「30分ごと」の時間ベーストリガーを設定してください。`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e);
    ss.toast('セットアップ中にエラーが発生しました。', '失敗', 10);
    ui.alert('セットアップエラー:\n' + e.message, ui.ButtonSet.OK);
  }
}


/**
 * [STEP 2: トリガー実行] ナレッジ生成の「バッチ処理」を行うワーカー関数
 * 1. _詳細スライド生成作業リスト シートから「未処理」のタスクを取得
 * 2. 時間の許す限りAPI処理を実行
 * 3. 処理結果を 保全ナレッジ_結果 シートに追記
 */
function generateKnowledge_PROCESS() {
  const startTime = new Date().getTime();
  const taskExecutionTimes = []; // タスクごとの実行時間を記録

  try {
    // --- 1. 必要なシートと設定を取得 ---
    const workSheet = ss.getSheetByName(WORK_LIST_SHEET_NAME);
    const outputSheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
    const knowledgeConfigSheet = ss.getSheetByName('カテゴリごとに知見作成');

    if (!workSheet || !outputSheet || !knowledgeConfigSheet) {
      Logger.log("必要なシート（_詳細スライド生成作業リスト, 保全ナレッジ_結果, カテゴリごとに知見作成）がありません。処理を終了します。");
      return; // トリガーなのでエラーは出さずに終了
    }

    _showProgress('保全ナレッジ生成処理を開始します...', '📝 ナレッジ生成', 3);

    const basePrompt = knowledgeConfigSheet.getRange('C31').getValue();
    const inputSheetName = knowledgeConfigSheet.getRange('C6').getValue();
    const inputSheet = ss.getSheetByName(inputSheetName);
    if (!inputSheet) {
      Logger.log(`入力シート「${inputSheetName}」が見つかりません。`);
      return;
    }

    // 元データをすべて読み込む（グループ復元用）
    const allData = inputSheet.getDataRange().getValues();
    const header = allData[0];

    // --- 2. 未処理のタスクを検索 ---
    const workRange = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 3);
    const workValues = workRange.getValues();

    let processedCountInThisRun = 0;
    let isFirstOutput = (outputSheet.getLastRow() <= 1);

    // --- 3. バッチ処理ループ ---
    for (let i = 0; i < workValues.length; i++) {
      const currentStatus = workValues[i][2]; // ステータス列

      // 未処理のタスクか？
      if (currentStatus === STATUS_EMPTY) {

        // 動的タイムアウトチェック：次のタスクを実行可能かを判定
        if (!_shouldContinueProcessing(startTime, taskExecutionTimes)) {
          Logger.log(`次のタスクで30分を超える可能性があるため、処理を中断します。`);
          return; // 次のトリガー実行に任せる
        }

        const taskStartTime = new Date().getTime();
        const sheetRow = i + 2; // スプレッドシートの実際の行番号
        const groupKey = workValues[i][0];
        const targetRowNumbers = JSON.parse(workValues[i][1]); // ["2", "5", "10"] など

        try {
          // 3a. ステータスを「処理中」に更新
          workSheet.getRange(sheetRow, 3).setValue(STATUS_PROCESSING);
          
          // 3b. グループデータを復元
          const groupRows = targetRowNumbers.map(rowNum => {
            // allData は 0-indexed, ヘッダーが0行目。
            // 2行目のデータは allData[1]
            return allData[rowNum - 1]; 
          });

          // 3c. CSVチャンクを作成 (元のコードと同じ)
          const csvChunk = [header] 
                            .concat(groupRows) 
                            .map(row =>
                               row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
                             ).join('\n');
          
          // 3d. プロンプトを構築 (元のコードと同じ)
          let prompt = basePrompt;
          prompt += `\n\n# 今回分析するデータセット (CSV形式)\n以下のデータは「${groupKey.replace(/\|/g, ', ')}」の値がすべて同じグループです。\n---\n${csvChunk}`;

          // 3e. APIを呼び出し (robustFetch_ を使う callGemini_ を想定)
          const resultText = callGemini_(prompt); 

          // 3f. 結果を解析
          const jsonStringMatch = resultText.match(/```json\s*([\s\S]*?)\s*```/);
          const cleanedJsonString = jsonStringMatch ? jsonStringMatch[1] : resultText;
          
          let newResults = [];
          if (cleanedJsonString.trim() !== "") {
            const parsedJson = JSON.parse(cleanedJsonString);
            newResults = Array.isArray(parsedJson) ? parsedJson : [parsedJson];
          }

          if (newResults.length === 0) {
            throw new Error("APIから有効なJSONが返されませんでした。");
          }

          // 3g. 結果を出力シートに「追記」
          const outputHeader = Object.keys(newResults[0]);
          const outputData = newResults.map(item => {
            return outputHeader.map(key => item[key] || "");
          });

          if (isFirstOutput) {
            // 初回書き込み時のみヘッダーを書き込む
            outputSheet.clear(); // "処理待機中..." を消す
            outputSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
            isFirstOutput = false; // フラグを下ろす
          }

          // 最終行に追記
          const lastRow = outputSheet.getLastRow();
          outputSheet.getRange(lastRow + 1, 1, outputData.length, outputData[0].length)
            .setValues(outputData)
            .setWrap(true)
            .setVerticalAlignment('top');

          // 3h. ステータスを「完了」に更新
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
              '📝 ナレッジ生成中',
              2
            );
          }

        } catch (e) {
          // 3i. エラー処理
          Logger.log(`グループ "${groupKey}" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
          workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);

          // エラーの場合も実行時間を記録
          const taskEndTime = new Date().getTime();
          const taskDuration = taskEndTime - taskStartTime;
          taskExecutionTimes.push(taskDuration);
        }

        // Utilities.sleep(SLEEP_MS_PER_GROUP); // API負荷軽減 (robustFetch_ で制御しているなら不要かも)
      }
    }


    Logger.log(`今回の実行で ${processedCountInThisRun} 件のグループを処理しました。`);

    // 1. シートへの書き込みを強制的に反映させる
    SpreadsheetApp.flush(); 

    // 2. 完了チェックのために、作業シートから「最新の」ステータスを再取得する
    const lastRow = workSheet.getLastRow();
    let remainingTasks = 0; // デフォルト値

    if (lastRow >= 2) { // データ行が1行以上ある場合
      // 3列目（ステータス列）の値だけを再取得
      const newStatusValues = workSheet.getRange(2, 3, lastRow - 1, 1).getValues();
      
      // 最新のステータス配列を元に残タスクを計算
      remainingTasks = newStatusValues.filter(
        row => row[0] === STATUS_EMPTY || row[0] === STATUS_PROCESSING
      ).length;
    }
    // データ行がない (lastRow < 2) 場合、remainingTasks は 0 のまま（正しい）


    // 「今回の実行で処理したタスクがあり」かつ「（最新のステータスで）残タスクが0になった」場合
    if (remainingTasks === 0 && processedCountInThisRun > 0) {

    // (A) 「_詳細スライド生成作業リスト」シートのD1セルから、使用する結果シート名を取得
    const newSheetName = workSheet.getRange("D1").getValue();
    if (!newSheetName) {
       Logger.log("エラー: _詳細スライド生成作業リスト シートのD1セルに結果シート名がありません。SETUPを先に実行してください。");
      return;
    }

    // (B) 完了したシート名をタイムスタンプ付きに「リネーム（名前変更）」する
      try {
        outputSheet.setName(newSheetName);
        Logger.log(`シート名を「${newSheetName}」に変更しました。`);
      } catch (e) {
        Logger.log(`シート名変更中にエラー: ${e}`);
        // （もし同名シートが既にあっても）処理は続行する
      }

      Logger.log("すべてのグループの処理が完了しました。");
      _showProgress('すべてのナレッジ生成が完了しました！', '✅ 完了', 10);
      
      
      // (オプション) ここでトリガーを自動停止する処理も追加可能
      stopTriggers_(); // ※別途 stopTriggers_() 関数を作成する必要があります
    }

  } catch (e) {
    Logger.log(`バッチ処理ワーカーで致命的なエラーが発生しました: ${e}`);
  }
}

/**
 * [新規] 'generateKnowledge_PROCESS' を実行するトリガーを自動停止する関数
 * 注: commonHelpers.js の stopTriggers_() を使用することもできます
 */
function stopTriggers_() {
  // commonHelpers.js の汎用版を利用
  stopTriggers_('generateKnowledge_PROCESS');
}