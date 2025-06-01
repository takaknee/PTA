/**
 * PTAスケジュール通知機能
 * イベント通知の送信と定期配信の設定を提供します
 */

/**
 * イベント通知を送信する
 */
function sendEventNotification() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // イベント名の入力
    const eventNameResponse = ui.prompt(
      'イベント通知',
      'イベント名を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (eventNameResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const eventName = eventNameResponse.getResponseText().trim();
    if (!eventName) {
      ui.alert('エラー', 'イベント名を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // イベント日時の入力
    const eventDateResponse = ui.prompt(
      'イベント通知',
      'イベント日時を入力してください (YYYY-MM-DD HH:MM形式):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (eventDateResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const eventDateStr = eventDateResponse.getResponseText().trim();
    const eventDate = new Date(eventDateStr);
    
    if (isNaN(eventDate.getTime())) {
      ui.alert('エラー', '有効な日時を入力してください (例: 2024-12-31 10:00)。', ui.ButtonSet.OK);
      return;
    }
    
    // イベント場所の入力
    const eventLocationResponse = ui.prompt(
      'イベント通知',
      'イベント場所を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (eventLocationResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const eventLocation = eventLocationResponse.getResponseText().trim();
    if (!eventLocation) {
      ui.alert('エラー', 'イベント場所を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // イベント詳細の入力
    const eventDetailsResponse = ui.prompt(
      'イベント通知',
      'イベント詳細を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (eventDetailsResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const eventDetails = eventDetailsResponse.getResponseText().trim();
    if (!eventDetails) {
      ui.alert('エラー', 'イベント詳細を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 参加確認が必要かどうか
    const needsRsvpResponse = ui.alert(
      'イベント通知',
      '参加確認（RSVP）が必要ですか？',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (needsRsvpResponse === ui.Button.CANCEL) {
      return;
    }
    
    const needsRsvp = (needsRsvpResponse === ui.Button.YES);
    
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
    
    // イベント情報をカレンダーに追加するかどうか
    const addToCalendarResponse = ui.alert(
      'カレンダー追加',
      'このイベントをGoogleカレンダーに追加しますか？',
      ui.ButtonSet.YES_NO
    );
    
    let calendarEvent = null;
    if (addToCalendarResponse === ui.Button.YES) {
      calendarEvent = createCalendarEvent(eventName, eventDate, eventLocation, eventDetails);
    }
    
    // 通知メールの作成と送信
    const emailSubject = `【PTA】イベントのお知らせ: ${eventName}`;
    const emailBody = createEventNotificationEmailBody(eventName, eventDate, eventLocation, eventDetails, needsRsvp, calendarEvent);
    
    // 送信確認
    const confirmResponse = ui.alert(
      '送信確認',
      `イベント通知を${activeMembers.length}名のメンバーに送信します。\n\nイベント: ${eventName}\n日時: ${Utilities.formatDate(eventDate, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm')}\n場所: ${eventLocation}\n\n送信しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    const result = sendEmailToMembers(activeMembers, emailSubject, emailBody, 'イベント通知');
    
    ui.alert(
      '送信完了',
      `イベント通知の送信が完了しました。\n\n配信先: ${result.totalCount}名\n成功: ${result.successCount}名\n失敗: ${result.failureCount}名`,
      ui.ButtonSet.OK
    );
    
    logInfo(`イベント通知を送信しました: ${eventName} (成功: ${result.successCount}, 失敗: ${result.failureCount})`);
    
  } catch (error) {
    logError('イベント通知送信中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `イベント通知送信中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 定期配信を設定する
 */
function setupScheduledDistribution() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const response = ui.alert(
      '定期配信設定',
      '定期配信の設定を行います。\n\n1. 新しい定期配信を設定\n2. 既存の定期配信を確認\n3. 定期配信を停止\n\n新しい定期配信を設定しますか？',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      createScheduledDistribution();
    } else if (response === ui.Button.NO) {
      manageExistingScheduledDistributions();
    }
    
  } catch (error) {
    logError('定期配信設定中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `定期配信設定中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Googleカレンダーにイベントを作成する
 * @param {string} eventName - イベント名
 * @param {Date} eventDate - イベント日時
 * @param {string} eventLocation - イベント場所
 * @param {string} eventDetails - イベント詳細
 * @returns {GoogleAppsScript.Calendar.CalendarEvent|null} 作成されたイベント
 */
function createCalendarEvent(eventName, eventDate, eventLocation, eventDetails) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    
    // イベント終了時刻（開始から2時間後と仮定）
    const endDate = new Date(eventDate.getTime() + (2 * 60 * 60 * 1000));
    
    const event = calendar.createEvent(
      eventName,
      eventDate,
      endDate,
      {
        description: eventDetails,
        location: eventLocation,
        guests: getActiveMembers().map(member => member.email).join(','),
        sendInvites: false // 手動で通知するため、自動招待は無効にする
      }
    );
    
    logInfo(`カレンダーイベントを作成しました: ${eventName}`);
    return event;
    
  } catch (error) {
    logError('カレンダーイベント作成中にエラーが発生しました', error);
    return null;
  }
}

/**
 * イベント通知メールの本文を作成する
 * @param {string} eventName - イベント名
 * @param {Date} eventDate - イベント日時
 * @param {string} eventLocation - イベント場所
 * @param {string} eventDetails - イベント詳細
 * @param {boolean} needsRsvp - 参加確認が必要かどうか
 * @param {GoogleAppsScript.Calendar.CalendarEvent} calendarEvent - カレンダーイベント
 * @returns {string} メール本文
 */
function createEventNotificationEmailBody(eventName, eventDate, eventLocation, eventDetails, needsRsvp, calendarEvent) {
  const eventDateStr = Utilities.formatDate(eventDate, 'Asia/Tokyo', 'yyyy年MM月dd日（E） HH:mm');
  
  let rsvpText = '';
  if (needsRsvp) {
    rsvpText = `
【参加確認】
参加の可否について、お手数ですが以下までご連絡ください。
連絡先: ${Session.getActiveUser().getEmail()}
※参加確認の締切: ${Utilities.formatDate(new Date(eventDate.getTime() - (3 * 24 * 60 * 60 * 1000)), 'Asia/Tokyo', 'yyyy年MM月dd日')}
`;
  }
  
  let calendarText = '';
  if (calendarEvent) {
    calendarText = `
【カレンダー】
このイベントはGoogleカレンダーにも追加されています。
カレンダーでも確認できます。
`;
  }
  
  return `PTA関係者の皆様

いつもPTA活動にご協力いただき、ありがとうございます。

以下のとおりイベントを開催いたしますので、お知らせいたします。

【イベント名】
${eventName}

【日時】
${eventDateStr}

【場所】
${eventLocation}

【詳細】
${eventDetails}${rsvpText}${calendarText}

ご不明な点がございましたら、お気軽にお問い合わせください。
皆様のご参加をお待ちしております。`;
}

/**
 * 新しい定期配信を作成する
 */
function createScheduledDistribution() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 配信名の入力
    const nameResponse = ui.prompt(
      '定期配信設定',
      '定期配信名を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (nameResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const distributionName = nameResponse.getResponseText().trim();
    if (!distributionName) {
      ui.alert('エラー', '配信名を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 配信頻度の選択
    const frequencyResponse = ui.prompt(
      '定期配信設定',
      '配信頻度を選択してください:\n\n1: 毎日\n2: 毎週\n3: 毎月\n\n番号を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (frequencyResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const frequencyChoice = frequencyResponse.getResponseText().trim();
    let frequency;
    
    switch (frequencyChoice) {
      case '1':
        frequency = 'daily';
        break;
      case '2':
        frequency = 'weekly';
        break;
      case '3':
        frequency = 'monthly';
        break;
      default:
        ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
        return;
    }
    
    // 配信時刻の入力
    const timeResponse = ui.prompt(
      '定期配信設定',
      '配信時刻を入力してください (HH:MM形式、例: 09:00):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (timeResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const timeStr = timeResponse.getResponseText().trim();
    const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})$/);
    
    if (!timeMatch) {
      ui.alert('エラー', '有効な時刻を入力してください (例: 09:00)。', ui.ButtonSet.OK);
      return;
    }
    
    const hour = parseInt(timeMatch[1]);
    const minute = parseInt(timeMatch[2]);
    
    if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
      ui.alert('エラー', '有効な時刻を入力してください (00:00-23:59)。', ui.ButtonSet.OK);
      return;
    }
    
    // メールテンプレートの選択
    const templates = getEmailTemplates();
    
    if (templates.length === 0) {
      ui.alert(
        '情報',
        '利用可能なメールテンプレートがありません。先にテンプレートを作成してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // テンプレート選択
    const templateList = templates
      .map((template, index) => `${index + 1}: ${template.name}`)
      .join('\n');
    
    const templateResponse = ui.prompt(
      '定期配信設定',
      `使用するメールテンプレート:\n\n${templateList}\n\n番号を入力してください:`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (templateResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const templateIndex = parseInt(templateResponse.getResponseText().trim()) - 1;
    if (isNaN(templateIndex) || templateIndex < 0 || templateIndex >= templates.length) {
      ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const selectedTemplate = templates[templateIndex];
    
    // トリガーを作成
    const trigger = createScheduledTrigger(distributionName, frequency, hour, minute, selectedTemplate);
    
    if (trigger) {
      ui.alert(
        '設定完了',
        `定期配信を設定しました。\n\n配信名: ${distributionName}\n頻度: ${getFrequencyDescription(frequency)}\n時刻: ${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}\nテンプレート: ${selectedTemplate.name}`,
        ui.ButtonSet.OK
      );
      
      logInfo(`定期配信を設定しました: ${distributionName} (${frequency})`);
    }
    
  } catch (error) {
    logError('定期配信作成中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `定期配信作成中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 定期配信用のトリガーを作成する
 * @param {string} distributionName - 配信名
 * @param {string} frequency - 配信頻度
 * @param {number} hour - 時
 * @param {number} minute - 分
 * @param {Object} template - メールテンプレート
 * @returns {GoogleAppsScript.Script.Trigger|null} 作成されたトリガー
 */
function createScheduledTrigger(distributionName, frequency, hour, minute, template) {
  try {
    // トリガー設定を保存
    const triggerConfig = {
      distributionName: distributionName,
      frequency: frequency,
      hour: hour,
      minute: minute,
      template: template,
      createdAt: new Date().toISOString()
    };
    
    const properties = PropertiesService.getScriptProperties();
    const triggerKey = `SCHEDULED_DISTRIBUTION_${distributionName}`;
    properties.setProperty(triggerKey, JSON.stringify(triggerConfig));
    
    // 実際のトリガーを作成
    let triggerBuilder = ScriptApp.newTrigger('executeScheduledDistribution')
      .timeBased()
      .atHour(hour);
    
    switch (frequency) {
      case 'daily':
        triggerBuilder = triggerBuilder.everyDays(1);
        break;
      case 'weekly':
        triggerBuilder = triggerBuilder.onWeekDay(ScriptApp.WeekDay.MONDAY);
        break;
      case 'monthly':
        triggerBuilder = triggerBuilder.onMonthDay(1);
        break;
    }
    
    const trigger = triggerBuilder.create();
    
    // トリガーIDを保存
    triggerConfig.triggerId = trigger.getUniqueId();
    properties.setProperty(triggerKey, JSON.stringify(triggerConfig));
    
    return trigger;
    
  } catch (error) {
    logError('トリガー作成中にエラーが発生しました', error);
    return null;
  }
}

/**
 * 定期配信を実行する（トリガーから呼び出される）
 */
function executeScheduledDistribution() {
  try {
    logInfo('定期配信を実行します');
    
    // 今日実行すべき定期配信を取得
    const scheduledDistributions = getScheduledDistributions();
    const currentHour = new Date().getHours();
    
    for (const distribution of scheduledDistributions) {
      if (distribution.hour === currentHour && shouldExecuteToday(distribution)) {
        // アクティブなメンバーを取得
        const activeMembers = getActiveMembers();
        
        if (activeMembers.length > 0) {
          // メール配信を実行
          const result = sendEmailToMembers(
            activeMembers,
            distribution.template.subject,
            distribution.template.body,
            '定期配信'
          );
          
          logInfo(`定期配信を実行しました: ${distribution.distributionName} (成功: ${result.successCount}, 失敗: ${result.failureCount})`);
        }
      }
    }
    
  } catch (error) {
    logError('定期配信実行中にエラーが発生しました', error);
  }
}

/**
 * 既存の定期配信を管理する
 */
function manageExistingScheduledDistributions() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const distributions = getScheduledDistributions();
    
    if (distributions.length === 0) {
      ui.alert(
        '情報',
        '設定されている定期配信がありません。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 定期配信一覧を表示
    const distributionList = distributions
      .map((dist, index) => `${index + 1}: ${dist.distributionName} (${getFrequencyDescription(dist.frequency)} ${String(dist.hour).padStart(2, '0')}:${String(dist.minute).padStart(2, '0')})`)
      .join('\n');
    
    const response = ui.prompt(
      '定期配信管理',
      `設定されている定期配信:\n\n${distributionList}\n\n停止する定期配信の番号を入力してください（キャンセルで終了）:`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const distributionIndex = parseInt(response.getResponseText().trim()) - 1;
    if (isNaN(distributionIndex) || distributionIndex < 0 || distributionIndex >= distributions.length) {
      ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    const selectedDistribution = distributions[distributionIndex];
    
    // 停止確認
    const confirmResponse = ui.alert(
      '停止確認',
      `定期配信「${selectedDistribution.distributionName}」を停止しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse === ui.Button.YES) {
      stopScheduledDistribution(selectedDistribution);
      
      ui.alert(
        '停止完了',
        `定期配信「${selectedDistribution.distributionName}」を停止しました。`,
        ui.ButtonSet.OK
      );
      
      logInfo(`定期配信を停止しました: ${selectedDistribution.distributionName}`);
    }
    
  } catch (error) {
    logError('定期配信管理中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `定期配信管理中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 定期配信を停止する
 * @param {Object} distribution - 停止する定期配信
 */
function stopScheduledDistribution(distribution) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const triggerKey = `SCHEDULED_DISTRIBUTION_${distribution.distributionName}`;
    
    // トリガーを削除
    if (distribution.triggerId) {
      const triggers = ScriptApp.getProjectTriggers();
      for (const trigger of triggers) {
        if (trigger.getUniqueId() === distribution.triggerId) {
          ScriptApp.deleteTrigger(trigger);
          break;
        }
      }
    }
    
    // 設定を削除
    properties.deleteProperty(triggerKey);
    
  } catch (error) {
    logError('定期配信停止中にエラーが発生しました', error);
  }
}

/**
 * 設定されている定期配信を取得する
 * @returns {Array} 定期配信の配列
 */
function getScheduledDistributions() {
  const properties = PropertiesService.getScriptProperties();
  const allProperties = properties.getProperties();
  const distributions = [];
  
  for (const key in allProperties) {
    if (key.startsWith('SCHEDULED_DISTRIBUTION_')) {
      try {
        const distributionData = JSON.parse(allProperties[key]);
        distributions.push(distributionData);
      } catch (error) {
        logError(`定期配信設定読み込みエラー: ${key}`, error);
      }
    }
  }
  
  return distributions;
}

/**
 * 今日実行すべき定期配信かどうかを判定する
 * @param {Object} distribution - 定期配信設定
 * @returns {boolean} 今日実行すべき場合はtrue
 */
function shouldExecuteToday(distribution) {
  const today = new Date();
  
  switch (distribution.frequency) {
    case 'daily':
      return true;
    case 'weekly':
      return today.getDay() === 1; // 月曜日
    case 'monthly':
      return today.getDate() === 1; // 月初
    default:
      return false;
  }
}

/**
 * 頻度の説明を取得する
 * @param {string} frequency - 頻度
 * @returns {string} 頻度の説明
 */
function getFrequencyDescription(frequency) {
  switch (frequency) {
    case 'daily':
      return '毎日';
    case 'weekly':
      return '毎週';
    case 'monthly':
      return '毎月';
    default:
      return frequency;
  }
}