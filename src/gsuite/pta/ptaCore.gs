/**
 * @OnlyCurrentDoc
 * PTA情報配信システムのメインスクリプト
 * PTA参加者の管理、情報配信、アンケート管理を統合的に行います
 * 
 * 必要な権限スコープ:
 * - https://www.googleapis.com/auth/gmail.send (メール送信)
 * - https://www.googleapis.com/auth/spreadsheets (スプレッドシート操作)
 * - https://www.googleapis.com/auth/forms (Googleフォーム操作)
 * - https://www.googleapis.com/auth/drive (ドライブファイル操作)
 * - https://www.googleapis.com/auth/calendar (カレンダー操作)
 */

// システム設定
const PTA_CONFIG = {
  // メンバー管理用シート名
  MEMBER_SHEET_NAME: 'PTAメンバー',
  // 配信履歴用シート名
  DISTRIBUTION_HISTORY_SHEET_NAME: '配信履歴',
  // アンケート管理用シート名
  SURVEY_SHEET_NAME: 'アンケート管理',
  // ログ用シート名
  LOG_SHEET_NAME: 'システムログ',
  // 最大一度に送信するメール数
  MAX_EMAIL_BATCH_SIZE: 50,
  // ログ有効フラグ
  LOG_ENABLED: true
};

/**
 * スプレッドシート初期化時に実行される関数
 * PTAシステムのメニューを追加します
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PTA情報配信システム')
    .addSubMenu(ui.createMenu('メンバー管理')
      .addItem('メンバー一覧を表示', 'showMemberList')
      .addItem('新しいメンバーを追加', 'addNewMember')
      .addItem('メンバーステータスを更新', 'updateMemberStatus')
      .addItem('メンバーをエクスポート', 'exportMemberList'))
    .addSubMenu(ui.createMenu('情報配信')
      .addItem('情報をメール配信', 'distributeInformation')
      .addItem('配信履歴を確認', 'showDistributionHistory')
      .addItem('配信テンプレートを管理', 'manageEmailTemplates'))
    .addSubMenu(ui.createMenu('アンケート管理')
      .addItem('新しいアンケートを作成', 'createSurvey')
      .addItem('アンケートを配信', 'distributeSurvey')
      .addItem('アンケート結果を確認', 'viewSurveyResults')
      .addItem('結果概要を配信', 'distributeSurveyResults'))
    .addSubMenu(ui.createMenu('スケジュール')
      .addItem('イベント通知を送信', 'sendEventNotification')
      .addItem('定期配信を設定', 'setupScheduledDistribution'))
    .addSubMenu(ui.createMenu('システム')
      .addItem('初期設定を実行', 'initializePTASystem')
      .addItem('システムログを確認', 'showSystemLogs')
      .addItem('設定を確認', 'showSystemConfig'))
    .addToUi();
}

/**
 * PTAシステムの初期設定を行う
 * 必要なシートを作成し、基本的な構造を準備します
 */
function initializePTASystem() {
  try {
    logInfo('PTAシステムの初期設定を開始します');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // メンバー管理シートの作成
    createMemberSheet(ss);
    
    // 配信履歴シートの作成
    createDistributionHistorySheet(ss);
    
    // アンケート管理シートの作成
    createSurveySheet(ss);
    
    // システムログシートの作成
    createLogSheet(ss);
    
    logInfo('PTAシステムの初期設定が完了しました');
    
    SpreadsheetApp.getUi().alert(
      '初期設定完了',
      'PTAシステムの初期設定が完了しました。\n各シートが作成され、システムが利用可能になりました。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    logError('PTAシステムの初期設定中にエラーが発生しました', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `初期設定中にエラーが発生しました: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * メンバー管理用シートを作成する
 * @param {SpreadsheetApp.Spreadsheet} ss - スプレッドシートオブジェクト
 */
function createMemberSheet(ss) {
  let sheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(PTA_CONFIG.MEMBER_SHEET_NAME);
    logInfo(`${PTA_CONFIG.MEMBER_SHEET_NAME}シートを作成しました`);
  } else {
    logInfo(`${PTA_CONFIG.MEMBER_SHEET_NAME}シートは既に存在します`);
    return;
  }
  
  // ヘッダー行の設定
  const headers = [
    'ID', '氏名', 'メールアドレス', 'ステータス', '登録日', '退会日', '備考'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 60);  // ID
  sheet.setColumnWidth(2, 120); // 氏名
  sheet.setColumnWidth(3, 200); // メールアドレス
  sheet.setColumnWidth(4, 80);  // ステータス
  sheet.setColumnWidth(5, 100); // 登録日
  sheet.setColumnWidth(6, 100); // 退会日
  sheet.setColumnWidth(7, 150); // 備考
  
  // データ検証の設定（ステータス列）
  const statusRange = sheet.getRange('D2:D1000');
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['アクティブ', '非アクティブ', '退会'])
    .build();
  statusRange.setDataValidation(statusValidation);
}

/**
 * 配信履歴管理用シートを作成する
 * @param {SpreadsheetApp.Spreadsheet} ss - スプレッドシートオブジェクト
 */
function createDistributionHistorySheet(ss) {
  let sheet = ss.getSheetByName(PTA_CONFIG.DISTRIBUTION_HISTORY_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(PTA_CONFIG.DISTRIBUTION_HISTORY_SHEET_NAME);
    logInfo(`${PTA_CONFIG.DISTRIBUTION_HISTORY_SHEET_NAME}シートを作成しました`);
  } else {
    logInfo(`${PTA_CONFIG.DISTRIBUTION_HISTORY_SHEET_NAME}シートは既に存在します`);
    return;
  }
  
  // ヘッダー行の設定
  const headers = [
    '配信ID', '配信日時', '件名', '配信タイプ', '配信先数', '成功数', '失敗数', '配信者', '備考'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 80);  // 配信ID
  sheet.setColumnWidth(2, 120); // 配信日時
  sheet.setColumnWidth(3, 200); // 件名
  sheet.setColumnWidth(4, 100); // 配信タイプ
  sheet.setColumnWidth(5, 80);  // 配信先数
  sheet.setColumnWidth(6, 80);  // 成功数
  sheet.setColumnWidth(7, 80);  // 失敗数
  sheet.setColumnWidth(8, 120); // 配信者
  sheet.setColumnWidth(9, 150); // 備考
}

/**
 * アンケート管理用シートを作成する
 * @param {SpreadsheetApp.Spreadsheet} ss - スプレッドシートオブジェクト
 */
function createSurveySheet(ss) {
  let sheet = ss.getSheetByName(PTA_CONFIG.SURVEY_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(PTA_CONFIG.SURVEY_SHEET_NAME);
    logInfo(`${PTA_CONFIG.SURVEY_SHEET_NAME}シートを作成しました`);
  } else {
    logInfo(`${PTA_CONFIG.SURVEY_SHEET_NAME}シートは既に存在します`);
    return;
  }
  
  // ヘッダー行の設定
  const headers = [
    'アンケートID', '作成日', 'タイトル', 'フォームURL', 'ステータス', '回答数', '締切日', '作成者', '備考'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 100); // アンケートID
  sheet.setColumnWidth(2, 100); // 作成日
  sheet.setColumnWidth(3, 200); // タイトル
  sheet.setColumnWidth(4, 300); // フォームURL
  sheet.setColumnWidth(5, 80);  // ステータス
  sheet.setColumnWidth(6, 80);  // 回答数
  sheet.setColumnWidth(7, 100); // 締切日
  sheet.setColumnWidth(8, 120); // 作成者
  sheet.setColumnWidth(9, 150); // 備考
}

/**
 * システムログ用シートを作成する
 * @param {SpreadsheetApp.Spreadsheet} ss - スプレッドシートオブジェクト
 */
function createLogSheet(ss) {
  let sheet = ss.getSheetByName(PTA_CONFIG.LOG_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(PTA_CONFIG.LOG_SHEET_NAME);
    logInfo(`${PTA_CONFIG.LOG_SHEET_NAME}シートを作成しました`);
  } else {
    logInfo(`${PTA_CONFIG.LOG_SHEET_NAME}シートは既に存在します`);
    return;
  }
  
  // ヘッダー行の設定
  const headers = [
    '日時', 'レベル', 'メッセージ', 'ユーザー', '詳細'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // 列幅の調整
  sheet.setColumnWidth(1, 120); // 日時
  sheet.setColumnWidth(2, 80);  // レベル
  sheet.setColumnWidth(3, 300); // メッセージ
  sheet.setColumnWidth(4, 120); // ユーザー
  sheet.setColumnWidth(5, 200); // 詳細
}

/**
 * システム設定を表示する
 */
function showSystemConfig() {
  const ui = SpreadsheetApp.getUi();
  
  const configText = Object.keys(PTA_CONFIG)
    .map(key => `${key}: ${PTA_CONFIG[key]}`)
    .join('\n');
  
  ui.alert(
    'システム設定',
    `現在のPTAシステム設定:\n\n${configText}`,
    ui.ButtonSet.OK
  );
}

/**
 * システムログを表示する
 */
function showSystemLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(PTA_CONFIG.LOG_SHEET_NAME);
  
  if (!logSheet) {
    SpreadsheetApp.getUi().alert(
      'エラー',
      'ログシートが見つかりません。先に初期設定を実行してください。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // ログシートをアクティブにする
  ss.setActiveSheet(logSheet);
  
  SpreadsheetApp.getUi().alert(
    'システムログ',
    'システムログシートを表示しました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 情報ログを出力する
 * @param {string} message - ログメッセージ
 */
function logInfo(message) {
  if (PTA_CONFIG.LOG_ENABLED) {
    Logger.log(`[INFO] ${message}`);
    writeLogToSheet('INFO', message);
  }
}

/**
 * エラーログを出力する
 * @param {string} message - エラーメッセージ
 * @param {Error} error - エラーオブジェクト
 */
function logError(message, error) {
  if (PTA_CONFIG.LOG_ENABLED) {
    const errorMessage = `${message}: ${error.toString()}`;
    Logger.log(`[ERROR] ${errorMessage}`);
    if (error.stack) {
      Logger.log(`[STACK] ${error.stack}`);
    }
    writeLogToSheet('ERROR', errorMessage, error.stack);
  }
}

/**
 * ログをシートに書き込む
 * @param {string} level - ログレベル
 * @param {string} message - メッセージ
 * @param {string} details - 詳細情報（オプション）
 */
function writeLogToSheet(level, message, details = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(PTA_CONFIG.LOG_SHEET_NAME);
    
    if (!logSheet) {
      return; // ログシートがない場合は何もしない
    }
    
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();
    
    logSheet.appendRow([timestamp, level, message, user, details]);
    
    // 古いログを削除（1000行以上の場合）
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1000) {
      logSheet.deleteRows(2, lastRow - 1000);
    }
    
  } catch (error) {
    // ログ書き込み中のエラーは無視
    Logger.log(`[LOG_ERROR] ${error.toString()}`);
  }
}