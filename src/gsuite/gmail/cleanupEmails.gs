/**
 * @OnlyCurrentDoc
 * Gmailの整理を行うためのユーティリティスクリプト
 * プロモーションメールの削除や特定ラベルのメール削除を行います
 * 
 * 必要な権限スコープ:
 * - https://www.googleapis.com/auth/gmail.readonly (メールの読み取り)
 * - https://www.googleapis.com/auth/gmail.modify (メールの削除)
 * - https://www.googleapis.com/auth/spreadsheets.currentonly (現在のスプレッドシートのみ操作)
 */

// バッチ処理の最大件数（Gmail APIの制限を考慮）
const BATCH_SIZE = 100;
// 一度の実行での最大処理数（実行時間制限を考慮）
const MAX_EMAILS_TO_PROCESS = 500;
// ログ設定
const LOG_ENABLED = true;

/**
 * メインメニューを設定する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gmail整理')
    .addItem('プロモーションメールを削除', 'deletePromotionEmails')
    .addItem('一年以上前のメールを削除', 'deleteOldEmails')
    .addItem('送信元別メール分析', 'analyzeSenders')
    .addItem('ラベル一覧を表示', 'listAllLabels')
    .addItem('特定ラベルのメールを削除', 'deleteEmailsByLabelPrompt')
    .addToUi();
}

/**
 * プロモーションカテゴリに分類されているメールを削除する
 * @returns {Object} 処理結果の情報
 */
function deletePromotionEmails() {
  try {
    const query = 'category:promotions';
    return batchDeleteEmails(query, 'プロモーション');
  } catch (error) {
    logError('プロモーションメール削除中にエラーが発生しました', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * 一年以上前のメールを削除する
 * @returns {Object} 処理結果の情報
 */
function deleteOldEmails() {
  try {
    // 一年前の日付を計算
    const oneYearAgo = new Date();
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    
    // Gmail検索クエリ用の日付フォーマット（YYYY/MM/DD）
    const dateString = `${oneYearAgo.getFullYear()}/${(oneYearAgo.getMonth() + 1).toString().padStart(2, '0')}/${oneYearAgo.getDate().toString().padStart(2, '0')}`;
    
    // 確認ダイアログを表示
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      `${dateString} より前のメールをすべて削除します。この操作は元に戻せません。続行しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      logInfo('一年以上前のメール削除をキャンセルしました');
      return { success: false, message: 'キャンセルされました' };
    }
    
    const query = `before:${dateString}`;
    const result = batchDeleteEmails(query, '一年以上前の');
    
    if (result.success) {
      ui.alert('完了', `${result.count}件の一年以上前のメールを削除しました。`, ui.ButtonSet.OK);
    }
    
    return result;
  } catch (error) {
    logError('一年以上前のメール削除中にエラーが発生しました', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `削除中にエラーが発生しました: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { success: false, error: error.toString() };
  }
}

/**
 * 指定されたラベルのメールを削除する
 * @param {string} labelName - 削除対象のラベル名
 * @returns {Object} 処理結果の情報
 */
function deleteEmailsByLabel(labelName) {
  try {
    if (!labelName) {
      throw new Error('ラベル名が指定されていません');
    }
    
    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      throw new Error(`ラベル "${labelName}" が見つかりません`);
    }
    
    const query = `label:${labelName}`;
    return batchDeleteEmails(query, labelName);
  } catch (error) {
    logError(`ラベル "${labelName}" のメール削除中にエラーが発生しました`, error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ユーザープロンプトを表示して特定ラベルのメールを削除する
 */
function deleteEmailsByLabelPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'ラベルメール削除',
    '削除するメールのラベル名を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // ユーザーの応答を確認
  if (response.getSelectedButton() == ui.Button.OK) {
    const labelName = response.getResponseText().trim();
    if (labelName) {
      // 削除前に確認
      const confirmResponse = ui.alert(
        '確認',
        `ラベル "${labelName}" のメールをすべて削除します。この操作は元に戻せません。続行しますか？`,
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResponse == ui.Button.YES) {
        const result = deleteEmailsByLabel(labelName);
        if (result.success) {
          ui.alert('完了', `${result.count}件のメールを削除しました。`, ui.ButtonSet.OK);
        } else {
          ui.alert('エラー', `削除中にエラーが発生しました: ${result.error}`, ui.ButtonSet.OK);
        }
      }
    } else {
      ui.alert('エラー', 'ラベル名を入力してください。', ui.ButtonSet.OK);
    }
  }
}

/**
 * バッチ処理でメールを削除する
 * @param {string} query - 検索クエリ
 * @param {string} type - メールの種類（ログ用）
 * @returns {Object} 処理結果の情報
 */
function batchDeleteEmails(query, type) {
  let processedCount = 0;
  let totalDeleted = 0;
  let continueProcessing = true;
  
  logInfo(`${type}メールの削除を開始します...`);
  
  while (continueProcessing) {
    // バッチサイズでスレッドを取得
    const threads = GmailApp.search(query, 0, BATCH_SIZE);
    
    if (threads.length === 0) {
      logInfo(`削除対象の${type}メールがこれ以上見つかりません`);
      break;
    }
    
    // スレッドを削除
    GmailApp.moveThreadsToTrash(threads);
    
    totalDeleted += threads.length;
    processedCount += threads.length;
    
    logInfo(`${processedCount}件の${type}メールを削除しました`);
    
    // 最大処理数に達したら一旦処理を終了
    if (processedCount >= MAX_EMAILS_TO_PROCESS) {
      continueProcessing = false;
      logInfo(`最大処理数(${MAX_EMAILS_TO_PROCESS}件)に達しました。残りは次回の実行で処理します。`);
    }
    
    // API制限を考慮して少し待機
    Utilities.sleep(100);
  }
  
  return {
    success: true,
    count: totalDeleted,
    message: `${totalDeleted}件の${type}メールを削除しました`
  };
}

/**
 * すべてのラベルを一覧表示する
 * @returns {Array} ラベルの配列
 */
function listAllLabels() {
  try {
    const labels = GmailApp.getUserLabels();
    const labelNames = labels.map(label => {
      const name = label.getName();
      const count = GmailApp.search(`label:${name}`).length;
      return { name, count };
    });
    
    // スプレッドシートに出力
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Gmail Labels');
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('Gmail Labels');
    } else {
      sheet.clear();
    }
    
    // ヘッダー行
    sheet.appendRow(['ラベル名', 'メール数']);
    
    // ラベルデータ
    labelNames.forEach(label => {
      sheet.appendRow([label.name, label.count]);
    });
    
    // 書式設定
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.autoResizeColumns(1, 2);
    
    SpreadsheetApp.getUi().alert(
      'ラベル一覧',
      `${labelNames.length}件のラベルを「Gmail Labels」シートに表示しました。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return labelNames;
  } catch (error) {
    logError('ラベル一覧の取得中にエラーが発生しました', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `ラベル一覧の取得中にエラーが発生しました: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return [];
  }
}

/**
 * 送信元別にメールを分析する
 * @returns {Array} 送信元別の統計情報
 */
function analyzeSenders() {
  try {
    logInfo('送信元別メール分析を開始します...');
    
    // 受信トレイのメールを取得（最大500件から分析）
    const threads = GmailApp.search('in:inbox', 0, 500);
    const senderStats = {};
    let processedCount = 0;
    
    logInfo(`${threads.length}件のスレッドを分析します`);
    
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const sender = message.getFrom();
        // メールアドレスを抽出（表示名を除去）
        const email = extractEmailAddress(sender);
        
        if (senderStats[email]) {
          senderStats[email].count++;
        } else {
          senderStats[email] = {
            email: email,
            displayName: sender,
            count: 1
          };
        }
        processedCount++;
      });
    });
    
    // 送信元を数の多い順にソート
    const sortedSenders = Object.values(senderStats)
      .sort((a, b) => b.count - a.count)
      .slice(0, 100); // 上位100件のみ
    
    // スプレッドシートに出力
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('送信元分析');
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('送信元分析');
    } else {
      sheet.clear();
    }
    
    // ヘッダー行
    sheet.appendRow(['送信元メールアドレス', '表示名', 'メール数']);
    
    // 送信元データ
    sortedSenders.forEach(sender => {
      sheet.appendRow([sender.email, sender.displayName, sender.count]);
    });
    
    // 書式設定
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.autoResizeColumns(1, 3);
    
    // 統計情報を表示
    const totalSenders = sortedSenders.length;
    const totalEmails = processedCount;
    const topSender = sortedSenders[0];
    
    SpreadsheetApp.getUi().alert(
      '送信元分析完了',
      `分析結果を「送信元分析」シートに表示しました。\n\n` +
      `・総メール数: ${totalEmails}件\n` +
      `・送信元数: ${totalSenders}件\n` +
      `・最多送信元: ${topSender ? topSender.email : 'なし'} (${topSender ? topSender.count : 0}件)`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    logInfo(`送信元分析が完了しました。${totalSenders}件の送信元、${totalEmails}件のメールを分析`);
    
    return sortedSenders;
  } catch (error) {
    logError('送信元分析中にエラーが発生しました', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `送信元分析中にエラーが発生しました: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return [];
  }
}

/**
 * メールアドレスを抽出する（表示名を除去）
 * @param {string} fromField - Fromフィールドの値
 * @returns {string} メールアドレス
 */
function extractEmailAddress(fromField) {
  // 角括弧内のメールアドレスを抽出
  const match = fromField.match(/<(.+?)>/);
  if (match) {
    return match[1];
  }
  
  // 角括弧がない場合は、そのまま返す（既にメールアドレスのみの場合）
  if (fromField.includes('@')) {
    return fromField.trim();
  }
  
  return fromField;
}

/**
 * 情報ログを出力する
 * @param {string} message - ログメッセージ
 */
function logInfo(message) {
  if (LOG_ENABLED) {
    Logger.log(`[INFO] ${message}`);
  }
}

/**
 * エラーログを出力する
 * @param {string} message - エラーメッセージ
 * @param {Error} error - エラーオブジェクト
 */
function logError(message, error) {
  if (LOG_ENABLED) {
    Logger.log(`[ERROR] ${message}: ${error.toString()}`);
    if (error.stack) {
      Logger.log(`[STACK] ${error.stack}`);
    }
  }
}
