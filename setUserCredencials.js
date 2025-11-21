/**
 * ユーザーに認証情報(JSON)の入力を求め、安全なユーザプロパティに保存する関数
 */
function setUserCredentials() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '認証情報の設定',
    'Google CloudからダウンロードしたサービスアカウントのJSONファイルの中身を全てここに貼り付けてください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK && response.getResponseText() != '') {
    const jsonString = response.getResponseText();
    try {
      const credentials = JSON.parse(jsonString);
      if (!credentials.private_key || !credentials.client_email || !credentials.project_id) {
        throw new Error('JSONに必要な情報(private_key, client_email, project_id)が含まれていません。');
      }
      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('GEMINI_PRIVATE_KEY', credentials.private_key);
      userProperties.setProperty('GEMINI_CLIENT_EMAIL', credentials.client_email);
      userProperties.setProperty('GEMINI_PROJECT_ID', credentials.project_id);
      ui.alert('成功', '認証情報を安全に保存しました。', ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('エラー', '入力されたテキストが正しいJSON形式ではありません。\n' + e.message, ui.ButtonSet.OK);
    }
  }
}