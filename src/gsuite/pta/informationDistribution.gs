/**
 * PTA情報配信機能
 * メンバーへのメール配信、配信履歴管理、テンプレート管理を提供します
 */

/**
 * 情報をメール配信する
 */
function distributeInformation() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // アクティブなメンバーを取得
    const activeMembers = getActiveMembers();
    
    if (activeMembers.length === 0) {
      ui.alert(
        '情報',
        'アクティブなメンバーがいません。先にメンバーを追加してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 件名の入力
    const subjectResponse = ui.prompt(
      '情報配信',
      '件名を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (subjectResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const subject = subjectResponse.getResponseText().trim();
    if (!subject) {
      ui.alert('エラー', '件名を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 本文の入力
    const bodyResponse = ui.prompt(
      '情報配信',
      '本文を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (bodyResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const body = bodyResponse.getResponseText().trim();
    if (!body) {
      ui.alert('エラー', '本文を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 送信確認
    const confirmResponse = ui.alert(
      '送信確認',
      `以下の内容で${activeMembers.length}名のメンバーにメールを送信します。\n\n件名: ${subject}\n\n本文: ${body.length > 100 ? body.substring(0, 100) + '...' : body}\n\n送信しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // メール配信の実行
    const result = sendEmailToMembers(activeMembers, subject, body, '情報配信');
    
    ui.alert(
      '配信完了',
      `メール配信が完了しました。\n\n配信先: ${result.totalCount}名\n成功: ${result.successCount}名\n失敗: ${result.failureCount}名`,
      ui.ButtonSet.OK
    );
    
    logInfo(`情報配信を実行しました: ${subject} (成功: ${result.successCount}, 失敗: ${result.failureCount})`);
    
  } catch (error) {
    logError('情報配信中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `情報配信中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 配信履歴を確認する
 */
function showDistributionHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(PTA_CONFIG.DISTRIBUTION_HISTORY_SHEET_NAME);
    
    if (!historySheet) {
      SpreadsheetApp.getUi().alert(
        'エラー',
        '配信履歴シートが見つかりません。先に初期設定を実行してください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 配信履歴シートをアクティブにする
    ss.setActiveSheet(historySheet);
    
    // 統計情報を取得
    const stats = getDistributionStatistics();
    
    SpreadsheetApp.getUi().alert(
      '配信履歴',
      `配信履歴を表示しました。\n\n統計情報:\n- 総配信回数: ${stats.totalDistributions}回\n- 今月の配信: ${stats.thisMonthDistributions}回\n- 総配信メール数: ${stats.totalEmails}通`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    logInfo('配信履歴を表示しました');
    
  } catch (error) {
    logError('配信履歴の表示中にエラーが発生しました', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `配信履歴の表示中にエラーが発生しました: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 配信テンプレートを管理する
 */
function manageEmailTemplates() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const response = ui.alert(
      'テンプレート管理',
      'テンプレート管理機能:\n\n1. 新しいテンプレートを作成\n2. 既存のテンプレートを使用\n\n新しいテンプレートを作成しますか？',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      createEmailTemplate();
    } else if (response === ui.Button.NO) {
      useEmailTemplate();
    }
    
  } catch (error) {
    logError('テンプレート管理中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `テンプレート管理中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * メンバーにメールを送信する
 * @param {Array} members - 送信先メンバーのリスト
 * @param {string} subject - 件名
 * @param {string} body - 本文
 * @param {string} distributionType - 配信タイプ
 * @returns {Object} 送信結果
 */
function sendEmailToMembers(members, subject, body, distributionType) {
  let successCount = 0;
  let failureCount = 0;
  const failures = [];
  
  // 送信者情報を取得
  const senderEmail = Session.getActiveUser().getEmail();
  
  // メール本文にフッターを追加
  const fullBody = `${body}\n\n---\nこのメールはPTA情報配信システムから自動送信されています。\n配信元: ${senderEmail}`;
  
  // バッチ処理でメール送信
  const batchSize = Math.min(PTA_CONFIG.MAX_EMAIL_BATCH_SIZE, members.length);
  
  for (let i = 0; i < members.length; i += batchSize) {
    const batch = members.slice(i, i + batchSize);
    
    for (const member of batch) {
      try {
        // 個別メールを送信
        GmailApp.sendEmail(
          member.email,
          subject,
          fullBody,
          {
            name: 'PTA情報配信システム'
          }
        );
        
        successCount++;
        logInfo(`メール送信成功: ${member.email}`);
        
      } catch (error) {
        failureCount++;
        failures.push({
          email: member.email,
          error: error.toString()
        });
        logError(`メール送信失敗: ${member.email}`, error);
      }
    }
    
    // API制限を考慮して少し待機
    if (i + batchSize < members.length) {
      Utilities.sleep(1000);
    }
  }
  
  // 配信履歴を記録
  recordDistributionHistory(subject, distributionType, members.length, successCount, failureCount);
  
  return {
    totalCount: members.length,
    successCount: successCount,
    failureCount: failureCount,
    failures: failures
  };
}

/**
 * 配信履歴をシートに記録する
 * @param {string} subject - 件名
 * @param {string} distributionType - 配信タイプ
 * @param {number} totalCount - 配信先総数
 * @param {number} successCount - 成功数
 * @param {number} failureCount - 失敗数
 */
function recordDistributionHistory(subject, distributionType, totalCount, successCount, failureCount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(PTA_CONFIG.DISTRIBUTION_HISTORY_SHEET_NAME);
    
    if (!historySheet) {
      logError('配信履歴シートが見つかりません', new Error('Sheet not found'));
      return;
    }
    
    // 配信IDを生成
    const lastRow = historySheet.getLastRow();
    const distributionId = `DIST${String(lastRow).padStart(4, '0')}`;
    
    // 現在の日時
    const timestamp = new Date();
    
    // 配信者
    const distributor = Session.getActiveUser().getEmail();
    
    // 履歴を記録
    historySheet.appendRow([
      distributionId,    // 配信ID
      timestamp,         // 配信日時
      subject,           // 件名
      distributionType,  // 配信タイプ
      totalCount,        // 配信先数
      successCount,      // 成功数
      failureCount,      // 失敗数
      distributor,       // 配信者
      ''                 // 備考
    ]);
    
    logInfo(`配信履歴を記録しました: ${distributionId}`);
    
  } catch (error) {
    logError('配信履歴の記録中にエラーが発生しました', error);
  }
}

/**
 * 配信統計情報を取得する
 * @returns {Object} 統計情報
 */
function getDistributionStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(PTA_CONFIG.DISTRIBUTION_HISTORY_SHEET_NAME);
  
  if (!historySheet) {
    return { totalDistributions: 0, thisMonthDistributions: 0, totalEmails: 0 };
  }
  
  const lastRow = historySheet.getLastRow();
  if (lastRow <= 1) {
    return { totalDistributions: 0, thisMonthDistributions: 0, totalEmails: 0 };
  }
  
  // データを取得
  const data = historySheet.getRange(2, 1, lastRow - 1, 9).getValues();
  
  const thisMonth = new Date();
  thisMonth.setDate(1);
  thisMonth.setHours(0, 0, 0, 0);
  
  let totalDistributions = data.length;
  let thisMonthDistributions = 0;
  let totalEmails = 0;
  
  data.forEach(row => {
    const distributionDate = new Date(row[1]); // 配信日時
    const recipientCount = row[4]; // 配信先数
    
    // 今月の配信をカウント
    if (distributionDate >= thisMonth) {
      thisMonthDistributions++;
    }
    
    // 総メール数を加算
    totalEmails += recipientCount;
  });
  
  return {
    totalDistributions: totalDistributions,
    thisMonthDistributions: thisMonthDistributions,
    totalEmails: totalEmails
  };
}

/**
 * メールテンプレートを作成する
 */
function createEmailTemplate() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // テンプレート名の入力
    const nameResponse = ui.prompt(
      'テンプレート作成',
      'テンプレート名を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (nameResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const templateName = nameResponse.getResponseText().trim();
    if (!templateName) {
      ui.alert('エラー', 'テンプレート名を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 件名テンプレートの入力
    const subjectResponse = ui.prompt(
      'テンプレート作成',
      '件名テンプレートを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (subjectResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const subjectTemplate = subjectResponse.getResponseText().trim();
    if (!subjectTemplate) {
      ui.alert('エラー', '件名テンプレートを入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 本文テンプレートの入力
    const bodyResponse = ui.prompt(
      'テンプレート作成',
      '本文テンプレートを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (bodyResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const bodyTemplate = bodyResponse.getResponseText().trim();
    if (!bodyTemplate) {
      ui.alert('エラー', '本文テンプレートを入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // テンプレートを保存
    saveEmailTemplate(templateName, subjectTemplate, bodyTemplate);
    
    ui.alert(
      '作成完了',
      `メールテンプレート「${templateName}」を作成しました。`,
      ui.ButtonSet.OK
    );
    
    logInfo(`メールテンプレートを作成しました: ${templateName}`);
    
  } catch (error) {
    logError('テンプレート作成中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `テンプレート作成中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 既存のメールテンプレートを使用する
 */
function useEmailTemplate() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 利用可能なテンプレートを取得
    const templates = getEmailTemplates();
    
    if (templates.length === 0) {
      ui.alert(
        '情報',
        '利用可能なテンプレートがありません。先にテンプレートを作成してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // テンプレート選択用のメッセージを作成
    const templateList = templates
      .map((template, index) => `${index + 1}: ${template.name}`)
      .join('\n');
    
    const response = ui.prompt(
      'テンプレート選択',
      `利用可能なテンプレート:\n\n${templateList}\n\n使用するテンプレートの番号を入力してください:`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const templateIndex = parseInt(response.getResponseText().trim()) - 1;
    if (isNaN(templateIndex) || templateIndex < 0 || templateIndex >= templates.length) {
      ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const selectedTemplate = templates[templateIndex];
    
    // テンプレートを使用してメール配信
    distributeInformationWithTemplate(selectedTemplate);
    
  } catch (error) {
    logError('テンプレート使用中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `テンプレート使用中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * メールテンプレートを保存する（プロパティサービスを使用）
 * @param {string} name - テンプレート名
 * @param {string} subject - 件名テンプレート
 * @param {string} body - 本文テンプレート
 */
function saveEmailTemplate(name, subject, body) {
  const properties = PropertiesService.getScriptProperties();
  const templateKey = `EMAIL_TEMPLATE_${name}`;
  const templateData = {
    name: name,
    subject: subject,
    body: body,
    createdAt: new Date().toISOString()
  };
  
  properties.setProperty(templateKey, JSON.stringify(templateData));
}

/**
 * 保存されているメールテンプレートを取得する
 * @returns {Array} テンプレートの配列
 */
function getEmailTemplates() {
  const properties = PropertiesService.getScriptProperties();
  const allProperties = properties.getProperties();
  const templates = [];
  
  for (const key in allProperties) {
    if (key.startsWith('EMAIL_TEMPLATE_')) {
      try {
        const templateData = JSON.parse(allProperties[key]);
        templates.push(templateData);
      } catch (error) {
        logError(`テンプレート読み込みエラー: ${key}`, error);
      }
    }
  }
  
  return templates.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

/**
 * テンプレートを使用して情報配信を行う
 * @param {Object} template - テンプレートデータ
 */
function distributeInformationWithTemplate(template) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // アクティブなメンバーを取得
    const activeMembers = getActiveMembers();
    
    if (activeMembers.length === 0) {
      ui.alert(
        '情報',
        'アクティブなメンバーがいません。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 送信確認
    const confirmResponse = ui.alert(
      '送信確認',
      `テンプレート「${template.name}」を使用して${activeMembers.length}名のメンバーにメールを送信します。\n\n件名: ${template.subject}\n\n本文: ${template.body.length > 100 ? template.body.substring(0, 100) + '...' : template.body}\n\n送信しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // メール配信の実行
    const result = sendEmailToMembers(activeMembers, template.subject, template.body, 'テンプレート配信');
    
    ui.alert(
      '配信完了',
      `テンプレートを使用したメール配信が完了しました。\n\n配信先: ${result.totalCount}名\n成功: ${result.successCount}名\n失敗: ${result.failureCount}名`,
      ui.ButtonSet.OK
    );
    
    logInfo(`テンプレート配信を実行しました: ${template.name} (成功: ${result.successCount}, 失敗: ${result.failureCount})`);
    
  } catch (error) {
    logError('テンプレート配信中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `テンプレート配信中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}