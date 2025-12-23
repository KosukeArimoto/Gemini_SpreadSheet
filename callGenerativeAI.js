/**
 * ===================================================================
 * API 制御用の定数
 * ===================================================================
 */
// 最大リトライ回数
const MAX_RETRIES = 5; 
// 429/50x系エラーの初回待機時間（ミリ秒）
const INITIAL_BACKOFF_MS = 1500; 
// API間の最小呼び出し間隔 (ミリ秒)。APIの仕様に応じて調整してください。
const MIN_INTERVAL_MS = 1000; 
// 1分あたりのリクエスト上限 (Vertex AIのデフォルト60より控えめに設定)
const RPM_LIMIT = 45; 
// キャッシュキー（レート制限管理用）
const CACHE_KEY_REQUESTS = 'API_REQUEST_TIMESTAMPS';
// キャッシュの有効期限（秒）
const CACHE_EXPIRATION_SECONDS = 70; // 60s + バッファ

/**
 * ===================================================================
 * [新規] API 制御のためのヘルパー関数
 * ===================================================================
 */

/**
 * API呼び出しの前に待機を挿入し、レート制限を遵守する (llm_processor.py の RateLimiter に相当)
 * LockServiceとCacheServiceを使い、複数の同時実行でもプロジェクト全体の制限を守ります。
 */
function rateLimiterWait_() {

  // スクリプト全体でロックを取得し、キャッシュの同時書き込みを防ぐ
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { // 10秒待ってもロックが取れなければエラー
    Logger.log("レートリミッターのロック取得に失敗しました。");
    // ロックが取れない場合でも、安全のために最小間隔は待機する
    Utilities.sleep(MIN_INTERVAL_MS + Math.floor(Math.random() * 1000));
    return;
  }

  try {
    const cache = CacheService.getScriptCache();
    const now = new Date().getTime();
    const oneMinuteAgo = now - 60000;

    // Logger.log("cache is "+cache);

    // 1. 最小間隔 (min_interval) のチェック
    const lastRequestTime = cache.get('LAST_API_REQUEST_TIME');
    if (lastRequestTime) {
      const elapsed = now - parseInt(lastRequestTime, 10);
      if (elapsed < MIN_INTERVAL_MS) {
        Utilities.sleep(MIN_INTERVAL_MS - elapsed);
      }
    }

    // 2. RPM (Requests Per Minute) のチェック
    let timestamps = [];
    const timestampsJson = cache.get(CACHE_KEY_REQUESTS);
    if (timestampsJson) {
      timestamps = JSON.parse(timestampsJson);
    }

    // 1分以上前のタイムスタンプを削除
    timestamps = timestamps.filter(ts => ts > oneMinuteAgo);

    // RPM制限に達しているかチェック
      // Logger.log("RPM制限は "+RPM_LIMIT)
      Logger.log("現在かかっている時間は"+timestamps.length)

    if (timestamps.length >= RPM_LIMIT) {
      const oldestTimestamp = timestamps[0];
      // 待つべき時間 = (1分前の最も古いリクエスト時刻 + 60秒) - 現在時刻
      const waitTime = (oldestTimestamp + 60000) - now; 
      if (waitTime > 0) {
        Logger.log(`RPM制限 (${RPM_LIMIT}) に達しました。${waitTime}ms 待機します。`);
        Utilities.sleep(waitTime);
      }
    }

    // 待機後の現在時刻を記録
    const newNow = new Date().getTime(); 
    timestamps.push(newNow);
    
    // キャッシュに保存
    cache.put(CACHE_KEY_REQUESTS, JSON.stringify(timestamps), CACHE_EXPIRATION_SECONDS);
    cache.put('LAST_API_REQUEST_TIME', newNow.toString(), CACHE_EXPIRATION_SECONDS);

  } catch(e) {
      Logger.log(`レートリミッターでエラーが発生しました: ${e}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * [新規] レート制限とリトライ処理を組み込んだ堅牢なUrlFetchAppラッパー
 * @param {string} url - 呼び出すURL
 * @param {object} options - UrlFetchAppのオプション
 * @return {GoogleAppsScript.URL_Fetch.HTTPResponse} - 成功したレスポンス
 */
function robustFetch_(url, options) {
  // 1. まずレート制限のチェックと待機を行う (Proactive)
  rateLimiterWait_();

  // 2. API呼び出しとリトライ処理 (Reactive)
  let lastError = null;
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    try {
      // muteHttpExceptions は options に設定されている前提
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        return response; // 成功
      }

      // 失敗時のログ
      lastError = response.getContentText();
      Logger.log(`API呼び出し失敗 (試行 ${attempt + 1}/${MAX_RETRIES}) - Code: ${responseCode}, Response: ${lastError.substring(0, 500)}`);

      // リトライすべきエラーか、即時停止すべきエラーか判断
      if (responseCode === 429 || responseCode >= 500) {
        // 429 (レート制限) または 50x (サーバーエラー) -> リトライ
        // 指数バックオフ + Jitter (ランダムな揺らぎ)
        const waitTime = (INITIAL_BACKOFF_MS * Math.pow(2, attempt)) + Math.floor(Math.random() * 1000);
        Logger.log(`... ${waitTime}ms 待機してリトライします ...`);
        Utilities.sleep(waitTime);
      } else {
        // 400, 401, 403, 404 などのクライアントエラー -> リトライ不要
        throw new Error(`リトライ不可能なエラー (${responseCode}): ${lastError}`);
      }

    } catch (e) {
      // ネットワークエラー (例: タイムアウト) や、上記でスローされたエラー
      lastError = e.message;
      Logger.log(`ネットワークエラーまたは例外 (試行 ${attempt + 1}/${MAX_RETRIES}): ${e.message}`);
      
      // ネットワークエラーの場合もリトライ
      if (attempt < MAX_RETRIES - 1) {
         const waitTime = (INITIAL_BACKOFF_MS * Math.pow(2, attempt)) + Math.floor(Math.random() * 1000);
         Logger.log(`... ${waitTime}ms 待機してリトライします ...`);
         Utilities.sleep(waitTime);
      }
    }
  }
  
  // すべてのリトライが失敗した場合
  throw new Error(`API呼び出しが ${MAX_RETRIES} 回の試行後に失敗しました。 最後のエラー: ${lastError}`);
}


/**
 * OpenAIのAPIキーをユーザプロパティに保存する
 */
function setOpenAiCredentials() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'OpenAI APIキーの設定',
    'OpenAIから発行されたAPIキー (sk-...) を貼り付けてください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText() != '') {
    PropertiesService.getUserProperties().setProperty('OPENAI_API_KEY', response.getResponseText());
    ui.alert('成功', 'OpenAI APIキーを安全に保存しました。', ui.ButtonSet.OK);
  }
}


/**
 * ===================================================================
 * 画像生成 API 関連
 * ===================================================================
 */

/**
 * 画像生成のメイン関数（モデル名に応じてAPIを切り替え）
 * @param {string} prompt - 画像生成用のプロンプト
 * @return {string} - Base64エンコードされた画像データ
 */
function generateImage_(prompt) {
  // configシートからモデル名を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
  if (!configSheet) {
    throw new Error('設定シート「config」が見つかりません。');
  }
  const model = configSheet.getRange('C7').getValue() || "gpt-image-1";
  console.log("Image model is " + model);

  // モデル名に応じてAPIを切り替え
  if (model.includes('imagen')) {
    // Gemini Imagen API
    return generateImageWithGemini_(prompt, model);
  } else {
    // OpenAI API (gpt-image-1, dall-e-3 など)
    return generateImageWithOpenAI_(prompt, model);
  }
}

/**
 * 後方互換性のためのエイリアス（既存コードからの呼び出しに対応）
 * @param {string} prompt - 画像生成用のプロンプト
 * @return {string} - Base64エンコードされた画像データ
 */
function callGPTApi_(prompt) {
  return generateImage_(prompt);
}

/**
 * OpenAI API と通信し、画像（Base64）を生成する関数
 * @param {string} prompt - 画像生成用のプロンプト
 * @param {string} model - 使用するモデル名 (gpt-image-1, dall-e-3 など)
 * @return {string} - Base64エンコードされた画像データ
 */
function generateImageWithOpenAI_(prompt, model) {
  const apiKey = PropertiesService.getUserProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OpenAI APIキーが設定されていません。「AI連携ツール」メニューから設定してください。');
  }

  const url = "https://api.openai.com/v1/images/generations";

  const payload = {
    "model": model,
    "prompt": prompt,
    "n": 1,
    "size": "1536x1024", // 横長のイラスト (16:9に近い)
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  const response = robustFetch_(url, options);
  const responseJson = JSON.parse(response.getContentText());

  if (responseJson.data && responseJson.data.length > 0) {
    return responseJson.data[0].b64_json;
  } else {
    throw new Error("OpenAI APIから画像が返されませんでした。");
  }
}

/**
 * Gemini Imagen API と通信し、画像（Base64）を生成する関数
 * @param {string} prompt - 画像生成用のプロンプト
 * @param {string} model - 使用するモデル名 (imagen-3.0-generate-001 など)
 * @return {string} - Base64エンコードされた画像データ
 */
function generateImageWithGemini_(prompt, model) {
  const userProperties = PropertiesService.getUserProperties();
  const projectId = userProperties.getProperty('GEMINI_PROJECT_ID');
  if (!projectId) {
    throw new Error('Gemini認証情報が設定されていません。「AI連携ツール」メニューから設定してください。');
  }

  const geminiService = getGeminiService();
  const accessToken = geminiService.getAccessToken();

  // configシートからリージョンを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
  const region = configSheet.getRange('C1').getValue() || "us-central1";

  // Vertex AI Imagen API エンドポイント
  const endpoint = region === "global"
    ? "https://aiplatform.googleapis.com"
    : `https://${region}-aiplatform.googleapis.com`;
  const url = `${endpoint}/v1/projects/${projectId}/locations/${region}/publishers/google/models/${model}:predict`;

  const payload = {
    "instances": [
      {
        "prompt": prompt
      }
    ],
    "parameters": {
      "sampleCount": 1,
      "aspectRatio": "16:9",
      "outputOptions": {
        "mimeType": "image/png"
      }
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + accessToken
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  const response = robustFetch_(url, options);
  const responseJson = JSON.parse(response.getContentText());

  if (responseJson.predictions && responseJson.predictions.length > 0) {
    // Imagen APIはbytesBase64Encodedフィールドで返す
    return responseJson.predictions[0].bytesBase64Encoded;
  } else {
    const errorMsg = responseJson.error ? responseJson.error.message : response.getContentText();
    throw new Error("Gemini Imagen APIから画像が返されませんでした: " + errorMsg);
  }
}

/**
 * 実際にGemini APIと通信する関数
 */
function callGeminiApi(prompt, projectId, accessToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
    if (!configSheet) {
      throw new Error('設定シート「config」が見つかりません。');
    }
  const region = configSheet.getRange('C1').getValue() || "us-central1";
  const model = configSheet.getRange('C2').getValue();

  // globalリージョンの場合はプレフィックスなし、それ以外は${region}-プレフィックス
  const endpoint = region === "global"
    ? "https://aiplatform.googleapis.com"
    : `https://${region}-aiplatform.googleapis.com`;
  const url = `${endpoint}/v1/projects/${projectId}/locations/${region}/publishers/google/models/${model}:generateContent`;

  const payload = {
    "contents": [{
      "role": "user",
      "parts": [{ "text": prompt }]
    }]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + accessToken
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  // const response = UrlFetchApp.fetch(url, options);
  // API通信を堅牢にするため修正
  const response = robustFetch_(url, options);
  const responseJson = JSON.parse(response.getContentText());

  if (responseJson.candidates && responseJson.candidates.length > 0) {
    return responseJson.candidates[0].content.parts[0].text;
  } else {
    // 安全性フィルターなどにより応答がない場合
    return "応答がありませんでした。入力内容が不適切と判断された可能性があります。 詳細: " + response.getContentText();
  }
}

/**
 * [OAuth2ライブラリ] 認証サービスを構築して返す関数
 */
function getGeminiService() {
  const userProperties = PropertiesService.getUserProperties();
  return OAuth2.createService('GeminiAPI')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(userProperties.getProperty('GEMINI_PRIVATE_KEY'))
    .setIssuer(userProperties.getProperty('GEMINI_CLIENT_EMAIL'))
    .setSubject(userProperties.getProperty('GEMINI_CLIENT_EMAIL'))
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/cloud-platform');
}


/**
 * Gemini API と通信し、テキストを生成する関数
 * @param {string} prompt - 画像生成用のプロンプト
 * @return {string} text - 回答
 */
function callGemini_(prompt) {
  const userProperties = PropertiesService.getUserProperties();
  const projectId = userProperties.getProperty('GEMINI_PROJECT_ID');
  if (!projectId) throw new Error('認証情報が設定されていません。');
  
  const geminiService = getGeminiService();
  const accessToken = geminiService.getAccessToken();
  const resultText = callGeminiApi(prompt, projectId, accessToken);
  return resultText;
}