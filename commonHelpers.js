// ===================================================================
// 共通ヘルパー関数
// 複数のファイルで使用される汎用的な関数をまとめたファイル
// ===================================================================

/**
 * Google DriveのフォルダURLからフォルダIDを抽出する
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
 * カンマ区切りとハイフンつなぎの数字の文字列（例: "1, 3, 5-9"）を
 * 数値の配列（例: [1, 3, 5, 6, 7, 8, 9]）に変換する
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
 * 列指定文字列（例: "A, C, E-G"）を0ベースのインデックス配列（例: [0, 2, 4, 5, 6]）に変換する
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

/**
 * 列文字（A, B, AA等）を0ベースのインデックスに変換する
 * @param {string} columnLetter - 列文字
 * @return {number} - 0ベースのインデックス、無効な場合は-1
 */
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
 * AIが生成したMarkdownテーブル形式のテキストを解析し、
 * スプレッドシート用の2次元配列に変換する
 * @param {string} markdownText - Markdownテーブル形式のテキスト
 * @return {Array<Array<string>>} - 2次元配列
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
 * プロンプト内のプレースホルダーを置換する
 * promptシートのB20:C28から置換リストを取得して処理
 * @param {string} originalPrompt - 元のプロンプト
 * @return {string} - 置換後のプロンプト
 */
function _replacePrompts(originalPrompt) {
  // B20からC28までの置換リストを一度に取得
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
 * Google DriveのURLからIDを抽出する（汎用版）
 * @param {string} url - Google DriveのURL
 * @return {string | null} - ID、見つからない場合はnull
 */
function extractGoogleDriveId_(url) {
  if (!url || typeof url !== 'string') return null;
  let id = null;
  let match = url.match(/[-\w]{25,}/);
  if (match && match[0]) { id = match[0]; }
  else { match = url.match(/[?&]id=([-\w]{25,})/); if (match && match[1]) { id = match[1]; } }
  return (id && id.length >= 25) ? id : null;
}

/**
 * セットアップ完了時のダイアログを表示する
 */
function _showSetupCompletionDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '✅ セットアップ完了',
    `タスクの準備が完了しました。\n\n次に対応する「実行」メニューを選択してください。\n\n実行は30分手前で自動停止し、繰り返し実行することで全タスクを完了します。`,
    ui.ButtonSet.OK
  );
}

/**
 * 指定された関数名のトリガーをすべて削除する
 * @param {string} functionName - トリガーを削除する関数名
 */
function stopTriggers_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`トリガーを削除: ${functionName}`);
    }
  }
}
