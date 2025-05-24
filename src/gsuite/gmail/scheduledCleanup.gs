/**
 * Gmail自動整理用の定期実行スクリプト
 * トリガーを設定して定期的に実行することを想定しています
 */

// 削除対象のラベル名のリスト
const TARGET_LABELS = [
  // 例: 'Ads', 'Newsletter', 'SocialMedia'
  // 実際に削除したいラベルをここに追加してください
];

/**
 * トリガーを設定するための関数
 * この関数を一度実行することで、dailyCleanup関数が毎日実行されるようになります
 */
function setupDailyTrigger() {
  // 既存のトリガーがあれば削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'dailyCleanup') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 毎日午前2時に実行するトリガーを設定
  ScriptApp.newTrigger('dailyCleanup')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
  
  Logger.log('毎日午前2時に自動実行されるトリガーを設定しました');
}

/**
 * トリガーで毎日実行される関数
 * プロモーションメールと設定された特定ラベルのメールを削除します
 */
function dailyCleanup() {
  try {
    Logger.log('日次メール整理を開始します');
    
    // 実行ロックを取得して重複実行を防止
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      Logger.log('別の実行が進行中のため、処理をスキップします');
      return;
    }
    
    // プロモーションメールを削除
    const promotionResult = deletePromotionEmails();
    Logger.log(`プロモーションメール削除: ${JSON.stringify(promotionResult)}`);
    
    // 各ターゲットラベルのメールを削除
    for (const labelName of TARGET_LABELS) {
      if (labelName && labelName.trim() !== '') {
        const labelResult = deleteEmailsByLabel(labelName);
        Logger.log(`ラベル "${labelName}" のメール削除: ${JSON.stringify(labelResult)}`);
      }
    }
    
    // 処理結果をメールで通知
    sendNotificationEmail();
    
    // ロックを解放
    lock.releaseLock();
    
    Logger.log('日次メール整理が完了しました');
  } catch (error) {
    Logger.log(`日次メール整理中にエラーが発生しました: ${error.toString()}`);
    if (error.stack) {
      Logger.log(`スタックトレース: ${error.stack}`);
    }
    
    // エラーがあってもロックを解放
    try {
      LockService.getScriptLock().releaseLock();
    } catch (e) {
      // ロック解放中のエラーは無視
    }
  }
}

/**
 * 処理結果を自分自身にメールで通知する
 */
function sendNotificationEmail() {
  const userEmail = Session.getActiveUser().getEmail();
  const date = new Date().toLocaleString();
  const subject = `Gmail自動整理レポート - ${date}`;
  
  // ラベルの統計情報を取得
  const labels = GmailApp.getUserLabels();
  let labelStats = '';
  for (const label of labels) {
    const name = label.getName();
    const count = GmailApp.search(`label:${name}`).length;
    labelStats += `- ${name}: ${count}件\n`;
  }
  
  const body = 
    `Gmailの自動整理処理が完了しました。\n\n` +
    `実行日時: ${date}\n\n` +
    `現在のラベル統計:\n${labelStats}\n\n` +
    `※このメールは自動送信されています。`;
  
  GmailApp.sendEmail(userEmail, subject, body);
}
