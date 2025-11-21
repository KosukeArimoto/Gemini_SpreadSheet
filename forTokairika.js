// /**
//  * [新規] 保全記録データを基に、AIを使って保全ナレッジを生成し、新しいシートに出力する関数
//  */
// function generateMaintenanceKnowledge() {

//   try {
//     ss.toast('保全ナレッジの生成を開始します...', '開始', 5);

//     // --- 1. 設定情報を取得 ---
//     const knowledgeConfigSheet = ss.getSheetByName('カテゴリごとに知見作成'); // 新しい設定シート
//     if (!configSheet || !promptSheet || !knowledgeConfigSheet) {
//       throw new Error('必要な設定シート（config, prompt, カテゴリごとに知見作成）が見つかりません。');
//     }

//     const inputSheetName = knowledgeConfigSheet.getRange('C6').getValue(); // 入力シート名をC6から取得
//     const basePrompt = knowledgeConfigSheet.getRange('C31').getValue(); // 基本プロンプトをC31から取得

//     // 分析対象の列名リストを取得 (C7からC11まで、空白は除外)
//     const targetColumns = knowledgeConfigSheet.getRange('C7:C11').getValues()
//                             .flat() // 2次元配列を1次元に変換
//                             .filter(String); // 空白を除外
//     if (targetColumns.length === 0) {
//       throw new Error('「カテゴリごとに知見作成」シートのC7:C11に分析対象の列名が指定されていません。');
//     }

//     if (!inputSheetName || !basePrompt) {
//       throw new Error('promptシートのC6(入力シート名)またはC31(プロンプト)が空です。');
//     }

//     // --- 2. 入力データを読み込む ---
//     const inputSheet = ss.getSheetByName(inputSheetName);
//     if (!inputSheet) throw new Error(`入力シート「${inputSheetName}」が見つかりません。`);

//     const allData = inputSheet.getDataRange().getValues();
//     const header = allData[0];
//     const dataRows = allData.slice(1);

//     if (dataRows.length === 0) {
//       throw new Error(`入力シート「${inputSheetName}」にデータがありません。`);
//     }

//     // --- 3. 指定された列のインデックスを特定 ---
//     const targetIndices = targetColumns.map(colName => {
//       const index = header.indexOf(colName);
//       if (index === -1) throw new Error(`入力シートのヘッダーに列名「${colName}」が見つかりません。`);
//       return index;
//     });
//     // 指定された列名だけのヘッダーを作成 (CSV生成用)
//     const targetHeader = targetColumns;

//     const groupedData = new Map(); // Map<グループキー, 行データの配列>

//     dataRows.forEach(row => {
//       // グループ化キーを作成 (指定列の値を結合)
//       const groupKey = targetIndices.map(index => row[index]).join('|'); // 区切り文字で結合

//       if (!groupedData.has(groupKey)) {
//         groupedData.set(groupKey, []);
//       }
//       groupedData.get(groupKey).push(row); // 同じキーを持つグループに行を追加
//     });
//     // console.log("groupedDataは"+groupedData.entries())

//     // --- 4. データを分割し、ループ処理 ---
//     let allResults = []; // 全てのグループからの結果を格納する配列
//     let processedGroups = 0;
//     const totalGroups = groupedData.size;

//     for (const [groupKey, groupRows] of groupedData.entries()) {
//       processedGroups++;
//       // グループキーから代表的な情報を取得して表示（例として最初の3要素）
//       const groupInfo = groupKey.split('|').slice(0, 3).join(', ');
//       ss.toast(`グループ ${processedGroups}/${totalGroups} を処理中 (${groupInfo})...`, 'API連携中', -1);

//       // --- 4a. このグループのデータをCSV形式に変換 ---
//       // ★注意：プロンプト指示に合わせて、渡す列をtargetColumnsだけに限定するか、全列渡すか要検討★
//       // ここでは *全列* をCSVにして渡す例 (プロンプト内で必要な列を参照させる想定)
//       const csvChunk = [header] // 全ヘッダー
//                         .concat(groupRows) // このグループの行データ (全列)
//                         .map(row =>
//                            row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
//                          ).join('\n');
//       // console.log(`Group Key: ${groupKey}, Row Count: ${groupRows.length}`);
//       // console.log("CSV Chunk:", csvChunk); // デバッグ用

//       // --- 4b. プロンプトを構築 ---
//       let prompt = basePrompt;
//       prompt += `\n\n# 今回分析するデータセット (CSV形式)\n以下のデータは「${targetHeader.join(', ')}」の値がすべて同じグループです。\n---\n${csvChunk}`;
//       // console.log("prompt is "+prompt)

//       // --- 4c. APIを呼び出し、結果を解析・結合 ---
//       const resultText = callGemini_(prompt); // API呼び出し
//       // console.log("resultTextは"+resultText);

//       try {
//         const jsonStringMatch = resultText.match(/```json\s*([\s\S]*?)\s*```/);
//         const cleanedJsonString = jsonStringMatch ? jsonStringMatch[1] : resultText;
//         if (cleanedJsonString.trim() !== "") {
//           const newResults = JSON.parse(cleanedJsonString);
//           // 結果が配列でない場合も考慮してconcat
//           allResults = allResults.concat(Array.isArray(newResults) ? newResults : [newResults]);
//         } else {
//           Logger.log(`警告: グループ "${groupKey}" のAPI応答が空でした。スキップします。`);
//         }
//       } catch (e) {
//         Logger.log(`警告: グループ "${groupKey}" のJSON解析エラー。スキップします。エラー: ${e}, API応答: ${resultText}`);
//         continue; // エラーがあっても次のグループへ
//       }

//       Utilities.sleep(1000); // API負荷軽減
//     }

//     // --- 7. 最終結果を動的に解釈してシートに出力 ---
//     if (allResults.length === 0) {
//       throw new Error("AIからの有効な応答（ナレッジ）がありませんでした。");
//     }

//     ss.toast('結果を出力しています...', '最終処理中', -1);

//     // 最初の結果オブジェクトからキーを取得してヘッダーにする
//     const outputHeader = Object.keys(allResults[0]);
//     // 全てのオブジェクトをヘッダーの順に値を取り出して2次元配列に変換
//     const outputData = allResults.map(item => {
//       return outputHeader.map(key => item[key] || ""); // 存在しないキーの場合は空文字
//     });

//     const resultSheetName = `保全ナレッジ_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
//     const resultSheet = ss.insertSheet(resultSheetName, ss.getNumSheets() + 1);

//     // ヘッダーとデータを書き込み
//     resultSheet.getRange(1, 1, 1, outputHeader.length).setValues([outputHeader]).setFontWeight('bold');
//     if (outputData.length > 0) {
//       resultSheet.getRange(2, 1, outputData.length, outputData[0].length)
//         .setValues(outputData)
//         .setWrap(true) // セル内折り返しを有効に
//         .setVerticalAlignment('top'); // 上揃え
//     }
//     resultSheet.autoResizeColumns(1, outputHeader.length); // 列幅を自動調整

//     ss.toast('処理が完了しました！', '完了', 5);
//     ui.alert('成功', `シート「${resultSheetName}」に保全ナレッジを出力しました。`, ui.ButtonSet.OK);

//   } catch (e) {
//     Logger.log(e);
//     ss.toast('エラーが発生しました。', '失敗', 10);
//     ui.alert('処理中にエラーが発生しました。\n\n詳細:\n' + e.message, ui.ButtonSet.OK);
//   }
// }





/**
 * [STEP 1: 手動実行] 保全ナレッジ生成の「セットアップ」を行う関数
 * 1. データを読み込み、グループ化する
 * 2. 作業リスト（_作業グループ）シートを作成する
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

    // --- 5. 作業リスト（_作業グループ）シートを作成 ---
    let workSheet = ss.getSheetByName(WORK_LIST_SHEET_NAME);
    if (workSheet) {
      workSheet.clear(); // 既存のシートをクリア
    } else {
      workSheet = ss.insertSheet(WORK_LIST_SHEET_NAME, ss.getNumSheets() + 1);
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

    const resultSheetName = `保全ナレッジ_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
    
    // 「_作業グループ」シートのD1セルに、今回使うシート名をメモとして書き込む
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
 * 1. _作業グループ シートから「未処理」のタスクを取得
 * 2. 時間の許す限りAPI処理を実行
 * 3. 処理結果を 保全ナレッジ_結果 シートに追記
 */
function generateKnowledge_PROCESS() {
  const startTime = new Date().getTime();
  
  try {
    // --- 1. 必要なシートと設定を取得 ---
    const workSheet = ss.getSheetByName(WORK_LIST_SHEET_NAME);
    const outputSheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
    const knowledgeConfigSheet = ss.getSheetByName('カテゴリごとに知見作成');
    
    if (!workSheet || !outputSheet || !knowledgeConfigSheet) {
      Logger.log("必要なシート（_作業グループ, 保全ナレッジ_結果, カテゴリごとに知見作成）がありません。処理を終了します。");
      return; // トリガーなのでエラーは出さずに終了
    }

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
        
        // 実行時間が5分を超えそうなら、自主的に終了
        const currentTime = new Date().getTime();
        if (currentTime - startTime > MAX_EXECUTION_TIME_MS) {
          Logger.log(`時間上限 (${MAX_EXECUTION_TIME_MS / 60000}分) に近づいたため、処理を中断します。`);
          return; // 次のトリガー実行に任せる
        }
        
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

        } catch (e) {
          // 3i. エラー処理
          Logger.log(`グループ "${groupKey}" (行 ${sheetRow}) の処理中にエラー: ${e.message}`);
          workSheet.getRange(sheetRow, 3).setValue(`${STATUS_ERROR}: ${e.message.substring(0, 200)}`);
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

    // (A) 「_作業グループ」シートのD1セルから、使用する結果シート名を取得
    const newSheetName = workSheet.getRange("D1").getValue();
    if (!newSheetName) {
       Logger.log("エラー: _作業グループ シートのD1セルに結果シート名がありません。SETUPを先に実行してください。");
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
      ss.toast('すべてのナレッジ生成が完了しました！', '完了', 10);
      
      
      // (オプション) ここでトリガーを自動停止する処理も追加可能
      stopTriggers_(); // ※別途 stopTriggers_() 関数を作成する必要があります
    }

  } catch (e) {
    Logger.log(`バッチ処理ワーカーで致命的なエラーが発生しました: ${e}`);
  }
}

/**
 * [新規] 'generateKnowledge_PROCESS' を実行するトリガーを自動停止する関数
 */
function stopTriggers_() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    triggers.forEach(trigger => {
      // 停止させたい関数名（'generateKnowledge_PROCESS'）と一致するかチェック
      if (trigger.getHandlerFunction() === 'generateKnowledge_PROCESS') {
        ScriptApp.deleteTrigger(trigger); // トリガーを削除
        deletedCount++;
      }
    });

    if (deletedCount > 0) {
      Logger.log(`${deletedCount}件の 'generateKnowledge_PROCESS' トリガーを削除しました。`);
      // (オプション) 完了をメールで通知する場合
      // MailApp.sendEmail(Session.getActiveUser().getEmail(), "バッチ処理完了とトリガー停止", "処理が完了したため、トリガーを自動停止しました。");
    } else {
      Logger.log("'generateKnowledge_PROCESS' を実行するトリガーは見つかりませんでした。");
    }
  } catch (e) {
    Logger.log(`トリガーの停止中にエラーが発生しました: ${e}`);
    // メール通知
    // MailApp.sendEmail(Session.getActiveUser().getEmail(), "トリガー停止エラー", `処理は完了しましたが、トリガーの自動停止に失敗しました。\nエラー: ${e.message}`);
  }
}