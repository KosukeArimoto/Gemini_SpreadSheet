// --- 完成版コード (OAuth2ライブラリ使用) ---
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 AI 連携ツール') // メニュー名を変更

    // --- 認証設定 ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔑 認証設定')
      .addItem('Google Cloud (Gemini) 認証', 'setUserCredentials')
      .addItem('OpenAI 認証', 'setOpenAiCredentials'))
    .addSeparator()

    // --- データ整理・分類フェーズ ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📊 分類・整理')
      .addItem('①-1 分類リストを生成 (セットアップ)', 'generateCategories_SETUP')
      .addItem('①-2 分類リストを生成 (実行)', 'generateCategories_PROCESS')
      .addItem('②-1 データに分類を付与 (セットアップ)', 'mergeCategories_SETUP')
      .addItem('②-2 データに分類を付与 (実行)', 'mergeCategories_PROCESS'))
    .addSeparator()

    // --- 設計フィードバックフェーズ ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📝 設計FB')
      .addItem('③-1 設計FBを生成 (セットアップ)', 'generateFeedback_SETUP')
      .addItem('③-2 設計FBを生成 (実行)', 'generateFeedback_PROCESS')
      .addItem('④-1 FBを個別に修正 (セットアップ)', 'reviseFeedback_SETUP')
      .addItem('④-2 FBを個別に修正 (実行)', 'reviseFeedback_PROCESS'))
    .addSeparator()

    // --- イラスト生成フェーズ ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🎨 イラスト生成')
      .addItem('⑤-1 イラスト用プロンプト案を生成 (セットアップ)', 'createIllustrationPrompts_SETUP')
      .addItem('⑤-2 イラスト用プロンプト案を生成 (実行)', 'createIllustrationPrompts_PROCESS')
      .addSeparator()
      .addItem('⑥-1 イラストを一括生成 (セットアップ)', 'createImages_SETUP')
      .addItem('⑥-2 イラストを一括生成 (実行)', 'createImages_PROCESS'))
    .addSeparator()

    // --- スライド生成フェーズ ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📽️ スライド生成')
      .addItem('⑦_1 スライド生成(TOMY)のセットアップ', 'createSlideTomy_SETUP')
      .addItem('⑦_2 スライド生成(TOMY)の実行', 'createSlides_PROCESS'))
    .addSeparator()

    // --- 自由分析 ---
    .addItem('⑧ 自由プロンプトを実行 (free promptシート)', 'freePrompt')
    .addSeparator()

    // --- 東海理科用ツール ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🌡️ 東海理科用')
      .addItem('1-1 詳細情報生成のセットアップ', 'generateKnowledge_SETUP')
      .addItem('1-2 詳細情報生成の実行', 'generateKnowledge_PROCESS')
      .addItem('1-3 スライド生成(詳細情報)のセットアップ', 'createSlideDetailTR_SETUP')
      .addItem('1-4 スライド生成(詳細情報)の実行', 'createSlides_PROCESS')
      .addItem('2-1 スライド生成(まとめ一覧)のセットアップ', 'createSlideSummaryTR_SETUP')
      .addItem('2-2 スライド生成(まとめ一覧)の実行', 'createSlides_PROCESS')
      .addSeparator()
      .addItem('🎨-1 行ごとの画像生成(セットアップ)', 'generateRowImages_SETUP')
      .addItem('🎨-2 行ごとの画像生成(実行)', 'generateRowImages_PROCESS'))

    .addToUi();
}

function dummyFunctionForPausingTrigger() {
  Logger.log('トリガーは現在、一時停止中です。');
}