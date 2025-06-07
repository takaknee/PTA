/**
 * [機能名] - [機能の簡単な説明]
 * 
 * @description [詳細な説明]
 * @param {type} paramName - パラメータの説明
 * @returns {type} 戻り値の説明
 * 
 * @example
 * const result = functionName(param);
 * console.log(result);
 * 
 * @author PTA Development Team
 * @since 2024-01-01
 */
function templateFunction(paramName) {
  try {
    // === メイン処理開始 ===
    logInfo(`[機能名] 処理開始: ${paramName}`);
    
    // 入力値検証
    if (!paramName) {
      throw new Error('必須パラメータが指定されていません');
    }
    
    // メイン処理ロジック
    let result;
    // TODO: 具体的な処理を実装
    
    // 成功ログ
    logInfo(`[機能名] 処理完了: ${result}`);
    return result;
    
  } catch (error) {
    // エラーログ（日本語）
    logError(`[機能名] 処理エラー`, error.message);
    
    // ユーザー向けエラーメッセージ
    throw new Error(`処理に失敗しました: ${error.message}`);
  }
}

/**
 * 共通エラーハンドリング関数
 */
function logInfo(message) {
  console.info(`[INFO] ${new Date().toISOString()}: ${message}`);
  Logger.log(`[INFO] ${message}`);
}

function logError(context, message) {
  const errorMsg = `[ERROR] ${new Date().toISOString()}: ${context} - ${message}`;
  console.error(errorMsg);
  Logger.log(errorMsg);
}

/**
 * 設定値の安全な取得
 */
function getConfigValue(key, defaultValue = null) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value || defaultValue;
  } catch (error) {
    logError(`設定取得`, `キー '${key}' の取得に失敗: ${error.message}`);
    return defaultValue;
  }
}

/**
 * スプレッドシートの安全な取得
 */
function getSafeSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート '${sheetName}' が見つかりません`);
  }
  
  return sheet;
}