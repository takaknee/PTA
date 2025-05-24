/**
 * @OnlyCurrentDoc
 * 認証関連のトラブルシューティングを行うヘルパースクリプト
 */

/**
 * 認証情報を確認するためのメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('認証管理')
    .addItem('権限の状態を確認', 'checkAuthStatus')
    .addItem('認証情報をリセット', 'resetAuth')
    .addToUi();
}

/**
 * 現在の権限状態を確認して表示する
 */
function checkAuthStatus() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Gmailの権限テスト
    const gmailLabels = GmailApp.getUserLabels();
    let gmailStatus = `Gmail API: アクセス可能 (${gmailLabels.length}個のラベルを読み取り)`;
    
    // スプレッドシートの権限テスト
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    let sheetStatus = `Spreadsheet API: アクセス可能 (${sheets.length}個のシートを読み取り)`;
    
    // ユーザー情報の取得
    const userEmail = Session.getActiveUser().getEmail();
    const effectiveUserEmail = Session.getEffectiveUser().getEmail();
    
    const message = 
      "権限ステータス確認結果:\n\n" +
      `${gmailStatus}\n` +
      `${sheetStatus}\n\n` +
      `アクティブユーザー: ${userEmail}\n` +
      `有効なユーザー: ${effectiveUserEmail}\n\n` +
      "すべての必要な権限が正常に設定されています。";
    
    ui.alert('権限確認', message, ui.ButtonSet.OK);
    
  } catch (error) {
    const errorMessage = 
      "権限エラーが検出されました:\n\n" +
      `${error.toString()}\n\n` +
      "「認証情報をリセット」を実行して権限を再度取得してください。";
    
    ui.alert('権限エラー', errorMessage, ui.ButtonSet.OK);
  }
}

/**
 * 認証情報をリセットするためのガイダンスを表示
 */
function resetAuth() {
  const ui = SpreadsheetApp.getUi();
  
  const message = 
    "認証情報のリセット手順:\n\n" +
    "1. Google アカウントのセキュリティ設定を開きます:\n" +
    "   https://myaccount.google.com/security\n\n" +
    "2. 「アプリのパスワードとアクセス権」→「サードパーティのアクセス権」を開きます\n\n" +
    "3. このスクリプトを見つけて削除します\n\n" +
    "4. ブラウザのキャッシュをクリアします\n\n" +
    "5. スプレッドシートを更新してスクリプトを再度実行します\n\n" +
    "これにより新しい権限リクエストが表示されます。";
  
  ui.alert('認証情報リセット', message, ui.ButtonSet.OK);
  
  // 実際にスクリプトから認証情報をクリアすることはできないため、
  // ユーザーに手動での操作を案内する
}

/**
 * スクリプトに必要な権限スコープを表示する
 */
function showRequiredScopes() {
  const ui = SpreadsheetApp.getUi();
  
  const scopes = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.container.ui"
  ];
  
  let scopesText = "";
  scopes.forEach(scope => {
    scopesText += `- ${scope}\n`;
  });
  
  const message = 
    "このスクリプトに必要な権限スコープ:\n\n" +
    scopesText + "\n" +
    "これらの権限が許可されていることを確認してください。";
  
  ui.alert('必要な権限', message, ui.ButtonSet.OK);
}
