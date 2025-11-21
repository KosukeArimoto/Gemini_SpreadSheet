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
      .addItem('① 分類リストを生成 (prompt1)', 'generateCategories')
      .addItem('② データに分類を付与 (prompt2)', 'mergeCategories'))
    .addSeparator()

    // --- 設計フィードバックフェーズ ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📝 設計FB')
      .addItem('③ 設計FBを生成 (prompt3)', 'generateFeedback')
      .addItem('④ FBを個別に修正', 'reviseFeedback'))
    .addSeparator()

    // --- イラスト生成フェーズ ---
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🎨 イラスト生成')
      .addItem('⑤ イラスト用プロンプト案を生成 (prompt4)', 'createIllustrationPrompts')
      .addItem('⑥ イラストを一括生成 (prompt5)', 'createImages'))
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
      .addItem('🎨 行ごとの画像生成', 'generateRowImages'))

    .addToUi();
}

function dummyFunctionForPausingTrigger() {
  Logger.log('トリガーは現在、一時停止中です。');
}