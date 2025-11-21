const ss = SpreadsheetApp.getActiveSpreadsheet();

const configSheet = ss.getSheetByName('config');
const promptSheet = ss.getSheetByName('prompt');
const variablesSetSheet = ss.getSheetByName('variables');
const freePromptSheet = ss.getSheetByName('free prompt');

const prompt1_pos = variablesSetSheet.getRange('D3').getValue();
const prompt2_pos = variablesSetSheet.getRange('D4').getValue();
const prompt3_pos = variablesSetSheet.getRange('D5').getValue();
const prompt4_pos = variablesSetSheet.getRange('D6').getValue();
const prompt5_pos = variablesSetSheet.getRange('D7').getValue();
const inputSheetName_pos = variablesSetSheet.getRange('D9').getValue();
const outputSheetName_pos = variablesSetSheet.getRange('D10').getValue();
const categorySheetName_pos = variablesSetSheet.getRange('D11').getValue();
const feedbackSheetName_pos = variablesSetSheet.getRange('D12').getValue();
const imagePromptSheetName_pos = variablesSetSheet.getRange('D14').getValue();
const imageTargetNum_pos = variablesSetSheet.getRange('D15').getValue();
const generateSlidesSheetName_pos = variablesSetSheet.getRange('D16').getValue();
const slideSaveDir_pos = variablesSetSheet.getRange('D17').getValue();
const imageSaveDir_pos = variablesSetSheet.getRange('D18').getValue();

const sep = parseInt(configSheet.getRange('C4').getValue(), 10); // 分割数を取得


// for東海理科
const tokaiPromptSheet = ss.getSheetByName('カテゴリごとに知見作成');


// ===================================================================
// バッチ処理用 定数
// ===================================================================

// 2つの関数で共有するシート名
const knowledgeConfigSheet = ss.getSheetByName('カテゴリごとに知見作成'); // 新しい設定シート
const inputSheetName = knowledgeConfigSheet.getRange('C6').getValue(); // 入力シート名をC6から取得
const WORK_LIST_SHEET_NAME = "_作業グループ"; // 作業リストを管理するシート
const OUTPUT_SHEET_NAME = `保全ナレッジ_結果`  // 最終的な結果を追記していくシート

// ステータスの定義
const STATUS_EMPTY = "";
const STATUS_PROCESSING = "処理中";
const STATUS_DONE = "完了";
const STATUS_ERROR = "エラー";

// GASの実行時間制限 (30分) より前に自主的に停止する (30分 = 1,800,000 ミリ秒)
const MAX_EXECUTION_TIME_MS = 1680000; // 28分用
// const MAX_EXECUTION_TIME_MS = 240000; // 4分用
// const MAX_EXECUTION_TIME_MS = 60000; // 1分用

// API間の最小待機時間 (堅牢なFetch側で制御しているなら不要な場合もある)
const SLEEP_MS_PER_GROUP = 1000;
const SLEEP_MS_PER_SLIDE = 1000;