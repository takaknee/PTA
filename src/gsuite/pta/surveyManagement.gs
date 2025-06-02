/**
 * PTAアンケート管理機能
 * Googleフォームを使用したアンケートの作成、配信、結果集計を提供します
 */

/**
 * 新しいアンケートを作成する
 */
function createSurvey() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // アンケートタイトルの入力
    const titleResponse = ui.prompt(
      'アンケート作成',
      'アンケートのタイトルを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (titleResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const title = titleResponse.getResponseText().trim();
    if (!title) {
      ui.alert('エラー', 'タイトルを入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // アンケート説明の入力
    const descriptionResponse = ui.prompt(
      'アンケート作成',
      'アンケートの説明を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (descriptionResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const description = descriptionResponse.getResponseText().trim();
    if (!description) {
      ui.alert('エラー', '説明を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 締切日の入力
    const dueDateResponse = ui.prompt(
      'アンケート作成',
      '締切日を入力してください (YYYY-MM-DD形式):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (dueDateResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const dueDateStr = dueDateResponse.getResponseText().trim();
    const dueDate = new Date(dueDateStr);
    
    if (isNaN(dueDate.getTime())) {
      ui.alert('エラー', '有効な日付を入力してください (例: 2024-12-31)。', ui.ButtonSet.OK);
      return;
    }
    
    // Googleフォームを作成
    const form = createGoogleForm(title, description);
    
    // アンケート情報をシートに記録
    const surveyId = recordSurveyInfo(title, form.getPublishedUrl(), dueDate);
    
    ui.alert(
      '作成完了',
      `アンケートを作成しました。\n\nタイトル: ${title}\nアンケートID: ${surveyId}\nフォームURL: ${form.getPublishedUrl()}\n\n次に質問を追加してください。`,
      ui.ButtonSet.OK
    );
    
    // 質問を追加するかどうか確認
    const addQuestionsResponse = ui.alert(
      '質問追加',
      '質問を追加しますか？',
      ui.ButtonSet.YES_NO
    );
    
    if (addQuestionsResponse === ui.Button.YES) {
      addQuestionsToSurvey(form);
    }
    
    logInfo(`新しいアンケートを作成しました: ${title} (ID: ${surveyId})`);
    
  } catch (error) {
    logError('アンケート作成中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `アンケート作成中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * アンケートを配信する
 */
function distributeSurvey() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 利用可能なアンケートを取得
    const surveys = getActiveSurveys();
    
    if (surveys.length === 0) {
      ui.alert(
        '情報',
        '配信可能なアンケートがありません。先にアンケートを作成してください。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // アンケート選択
    const survey = selectSurveyFromList(surveys);
    if (!survey) {
      return;
    }
    
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
    
    // 配信確認
    const confirmResponse = ui.alert(
      '配信確認',
      `アンケート「${survey.title}」を${activeMembers.length}名のメンバーに配信します。\n\nフォームURL: ${survey.formUrl}\n\n配信しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // アンケート配信メールの作成と送信
    const emailSubject = `【PTA】アンケートのお願い: ${survey.title}`;
    const emailBody = createSurveyEmailBody(survey);
    
    const result = sendEmailToMembers(activeMembers, emailSubject, emailBody, 'アンケート配信');
    
    ui.alert(
      '配信完了',
      `アンケート配信が完了しました。\n\n配信先: ${result.totalCount}名\n成功: ${result.successCount}名\n失敗: ${result.failureCount}名`,
      ui.ButtonSet.OK
    );
    
    logInfo(`アンケート配信を実行しました: ${survey.title} (成功: ${result.successCount}, 失敗: ${result.failureCount})`);
    
  } catch (error) {
    logError('アンケート配信中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `アンケート配信中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * アンケート結果を確認する
 */
function viewSurveyResults() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 利用可能なアンケートを取得
    const surveys = getAllSurveys();
    
    if (surveys.length === 0) {
      ui.alert(
        '情報',
        'アンケートがありません。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // アンケート選択
    const survey = selectSurveyFromList(surveys);
    if (!survey) {
      return;
    }
    
    // アンケート結果を取得
    const results = getSurveyResults(survey);
    
    // 結果を表示
    displaySurveyResults(survey, results);
    
    logInfo(`アンケート結果を表示しました: ${survey.title}`);
    
  } catch (error) {
    logError('アンケート結果確認中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `結果確認中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * アンケート結果概要を配信する
 */
function distributeSurveyResults() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 利用可能なアンケートを取得
    const surveys = getAllSurveys();
    
    if (surveys.length === 0) {
      ui.alert(
        '情報',
        'アンケートがありません。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // アンケート選択
    const survey = selectSurveyFromList(surveys);
    if (!survey) {
      return;
    }
    
    // アンケート結果を取得
    const results = getSurveyResults(survey);
    
    if (results.responseCount === 0) {
      ui.alert(
        '情報',
        'このアンケートにはまだ回答がありません。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 結果概要メールの作成
    const emailSubject = `【PTA】アンケート結果報告: ${survey.title}`;
    const emailBody = createSurveyResultsEmailBody(survey, results);
    
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
    
    // 配信確認
    const confirmResponse = ui.alert(
      '配信確認',
      `アンケート結果概要を${activeMembers.length}名のメンバーに配信します。\n\nアンケート: ${survey.title}\n回答数: ${results.responseCount}件\n\n配信しますか？`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // 結果概要を配信
    const result = sendEmailToMembers(activeMembers, emailSubject, emailBody, 'アンケート結果配信');
    
    ui.alert(
      '配信完了',
      `アンケート結果の配信が完了しました。\n\n配信先: ${result.totalCount}名\n成功: ${result.successCount}名\n失敗: ${result.failureCount}名`,
      ui.ButtonSet.OK
    );
    
    logInfo(`アンケート結果を配信しました: ${survey.title} (成功: ${result.successCount}, 失敗: ${result.failureCount})`);
    
  } catch (error) {
    logError('アンケート結果配信中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `結果配信中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Googleフォームを作成する
 * @param {string} title - フォームタイトル
 * @param {string} description - フォーム説明
 * @returns {GoogleAppsScript.Forms.Form} 作成されたフォーム
 */
function createGoogleForm(title, description) {
  const form = FormApp.create(title);
  form.setDescription(description);
  form.setCollectEmail(true);
  form.setLimitOneResponsePerUser(true);
  
  // デフォルトの感謝メッセージを設定
  form.setConfirmationMessage('ご回答ありがとうございました。回答内容は記録されました。');
  
  return form;
}

/**
 * アンケートに質問を追加する
 * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
 */
function addQuestionsToSurvey(form) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    let continueAdding = true;
    
    while (continueAdding) {
      // 質問タイプの選択
      const typeResponse = ui.prompt(
        '質問追加',
        '質問タイプを選択してください:\n\n1: テキスト入力\n2: 選択肢（単一選択）\n3: 選択肢（複数選択）\n4: 評価（1-5）\n\n番号を入力してください:',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (typeResponse.getSelectedButton() !== ui.Button.OK) {
        break;
      }
      
      const questionType = typeResponse.getResponseText().trim();
      
      // 質問文の入力
      const questionResponse = ui.prompt(
        '質問追加',
        '質問文を入力してください:',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (questionResponse.getSelectedButton() !== ui.Button.OK) {
        break;
      }
      
      const questionText = questionResponse.getResponseText().trim();
      if (!questionText) {
        ui.alert('エラー', '質問文を入力してください。', ui.ButtonSet.OK);
        continue;
      }
      
      // 質問タイプに応じて質問を追加
      switch (questionType) {
        case '1':
          form.addTextItem().setTitle(questionText).setRequired(true);
          break;
        case '2':
          addMultipleChoiceQuestion(form, questionText, false);
          break;
        case '3':
          addMultipleChoiceQuestion(form, questionText, true);
          break;
        case '4':
          form.addScaleItem()
            .setTitle(questionText)
            .setBounds(1, 5)
            .setLabels('全く思わない', 'とても思う')
            .setRequired(true);
          break;
        default:
          ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
          continue;
      }
      
      // 続けて質問を追加するか確認
      const continueResponse = ui.alert(
        '質問追加',
        'さらに質問を追加しますか？',
        ui.ButtonSet.YES_NO
      );
      
      continueAdding = (continueResponse === ui.Button.YES);
    }
    
  } catch (error) {
    logError('質問追加中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `質問追加中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 選択肢問題を追加する
 * @param {GoogleAppsScript.Forms.Form} form - フォーム
 * @param {string} questionText - 質問文
 * @param {boolean} allowMultiple - 複数選択を許可するか
 */
function addMultipleChoiceQuestion(form, questionText, allowMultiple) {
  const ui = SpreadsheetApp.getUi();
  
  // 選択肢の入力
  const choicesResponse = ui.prompt(
    '選択肢入力',
    '選択肢をカンマ区切りで入力してください（例: はい,いいえ,分からない）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (choicesResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const choicesText = choicesResponse.getResponseText().trim();
  if (!choicesText) {
    ui.alert('エラー', '選択肢を入力してください。', ui.ButtonSet.OK);
    return;
  }
  
  const choices = choicesText.split(',').map(choice => choice.trim()).filter(choice => choice);
  
  if (choices.length === 0) {
    ui.alert('エラー', '有効な選択肢を入力してください。', ui.ButtonSet.OK);
    return;
  }
  
  if (allowMultiple) {
    form.addCheckboxItem()
      .setTitle(questionText)
      .setChoiceValues(choices)
      .setRequired(true);
  } else {
    form.addMultipleChoiceItem()
      .setTitle(questionText)
      .setChoiceValues(choices)
      .setRequired(true);
  }
}

/**
 * アンケート情報をシートに記録する
 * @param {string} title - アンケートタイトル
 * @param {string} formUrl - フォームURL
 * @param {Date} dueDate - 締切日
 * @returns {string} 生成されたアンケートID
 */
function recordSurveyInfo(title, formUrl, dueDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const surveySheet = ss.getSheetByName(PTA_CONFIG.SURVEY_SHEET_NAME);
  
  if (!surveySheet) {
    throw new Error('アンケートシートが見つかりません');
  }
  
  // アンケートIDを生成
  const lastRow = surveySheet.getLastRow();
  const surveyId = `SURVEY${String(lastRow).padStart(4, '0')}`;
  
  // 現在の日時
  const createdAt = new Date();
  
  // 作成者
  const creator = Session.getActiveUser().getEmail();
  
  // アンケート情報を記録
  surveySheet.appendRow([
    surveyId,      // アンケートID
    createdAt,     // 作成日
    title,         // タイトル
    formUrl,       // フォームURL
    'アクティブ',   // ステータス
    0,             // 回答数
    dueDate,       // 締切日
    creator,       // 作成者
    ''             // 備考
  ]);
  
  return surveyId;
}

/**
 * アクティブなアンケートを取得する
 * @returns {Array} アクティブなアンケートの配列
 */
function getActiveSurveys() {
  return getAllSurveys().filter(survey => survey.status === 'アクティブ');
}

/**
 * 全てのアンケートを取得する
 * @returns {Array} アンケートの配列
 */
function getAllSurveys() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const surveySheet = ss.getSheetByName(PTA_CONFIG.SURVEY_SHEET_NAME);
  
  if (!surveySheet) {
    return [];
  }
  
  const lastRow = surveySheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  
  const data = surveySheet.getRange(2, 1, lastRow - 1, 9).getValues();
  
  return data.map(row => ({
    id: row[0],
    createdDate: row[1],
    title: row[2],
    formUrl: row[3],
    status: row[4],
    responseCount: row[5],
    dueDate: row[6],
    creator: row[7],
    notes: row[8]
  }));
}

/**
 * アンケートリストから選択する
 * @param {Array} surveys - アンケートの配列
 * @returns {Object|null} 選択されたアンケート
 */
function selectSurveyFromList(surveys) {
  const ui = SpreadsheetApp.getUi();
  
  // アンケート選択用のメッセージを作成
  const surveyList = surveys
    .map((survey, index) => `${index + 1}: ${survey.title} (作成日: ${Utilities.formatDate(survey.createdDate, 'Asia/Tokyo', 'yyyy-MM-dd')})`)
    .join('\n');
  
  const response = ui.prompt(
    'アンケート選択',
    `利用可能なアンケート:\n\n${surveyList}\n\n選択するアンケートの番号を入力してください:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  const surveyIndex = parseInt(response.getResponseText().trim()) - 1;
  if (isNaN(surveyIndex) || surveyIndex < 0 || surveyIndex >= surveys.length) {
    ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
    return null;
  }
  
  return surveys[surveyIndex];
}

/**
 * アンケート配信用のメール本文を作成する
 * @param {Object} survey - アンケート情報
 * @returns {string} メール本文
 */
function createSurveyEmailBody(survey) {
  const dueDateStr = Utilities.formatDate(survey.dueDate, 'Asia/Tokyo', 'yyyy年MM月dd日');
  
  return `PTA関係者の皆様

いつもPTA活動にご協力いただき、ありがとうございます。

アンケートのお願いをさせていただきます。
皆様のご意見をお聞かせください。

【アンケートタイトル】
${survey.title}

【回答期限】
${dueDateStr}

【アンケートURL】
${survey.formUrl}

※このアンケートはお一人様一回までの回答となっております。
※回答には5-10分程度お時間をいただきます。

ご協力のほど、よろしくお願いいたします。`;
}

/**
 * アンケート結果を取得する
 * @param {Object} survey - アンケート情報
 * @returns {Object} アンケート結果
 */
function getSurveyResults(survey) {
  try {
    // フォームIDをURLから抽出
    const formId = extractFormIdFromUrl(survey.formUrl);
    if (!formId) {
      throw new Error('フォームIDを取得できませんでした');
    }
    
    const form = FormApp.openById(formId);
    const responses = form.getResponses();
    
    return {
      responseCount: responses.length,
      responses: responses.map(response => ({
        timestamp: response.getTimestamp(),
        email: response.getRespondentEmail(),
        answers: response.getItemResponses().map(itemResponse => ({
          question: itemResponse.getItem().getTitle(),
          answer: itemResponse.getResponse()
        }))
      }))
    };
    
  } catch (error) {
    logError('アンケート結果取得中にエラーが発生しました', error);
    return {
      responseCount: 0,
      responses: [],
      error: error.toString()
    };
  }
}

/**
 * URLからフォームIDを抽出する
 * @param {string} url - フォームURL
 * @returns {string|null} フォームID
 */
function extractFormIdFromUrl(url) {
  const match = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * アンケート結果を表示する
 * @param {Object} survey - アンケート情報
 * @param {Object} results - アンケート結果
 */
function displaySurveyResults(survey, results) {
  const ui = SpreadsheetApp.getUi();
  
  if (results.error) {
    ui.alert(
      'エラー',
      `アンケート結果の取得中にエラーが発生しました: ${results.error}`,
      ui.ButtonSet.OK
    );
    return;
  }
  
  if (results.responseCount === 0) {
    ui.alert(
      'アンケート結果',
      `アンケート「${survey.title}」にはまだ回答がありません。`,
      ui.ButtonSet.OK
    );
    return;
  }
  
  // 結果をスプレッドシートに出力
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheetName = `Results_${survey.id}`;
  let resultSheet = ss.getSheetByName(resultSheetName);
  
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
  } else {
    resultSheet.clear();
  }
  
  // ヘッダー行を作成
  const headers = ['回答日時', 'メールアドレス'];
  if (results.responses.length > 0) {
    results.responses[0].answers.forEach(answer => {
      headers.push(answer.question);
    });
  }
  
  resultSheet.appendRow(headers);
  resultSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // データ行を追加
  results.responses.forEach(response => {
    const row = [
      response.timestamp,
      response.email
    ];
    
    response.answers.forEach(answer => {
      row.push(answer.answer);
    });
    
    resultSheet.appendRow(row);
  });
  
  // 列幅を自動調整
  resultSheet.autoResizeColumns(1, headers.length);
  
  // 結果シートをアクティブにする
  ss.setActiveSheet(resultSheet);
  
  ui.alert(
    'アンケート結果',
    `アンケート結果を「${resultSheetName}」シートに出力しました。\n\n回答数: ${results.responseCount}件`,
    ui.ButtonSet.OK
  );
}

/**
 * アンケート結果概要配信用のメール本文を作成する
 * @param {Object} survey - アンケート情報
 * @param {Object} results - アンケート結果
 * @returns {string} メール本文
 */
function createSurveyResultsEmailBody(survey, results) {
  const dueDateStr = Utilities.formatDate(survey.dueDate, 'Asia/Tokyo', 'yyyy年MM月dd日');
  
  let summaryText = '';
  
  // 簡単な統計情報を生成
  if (results.responses.length > 0) {
    const questionStats = {};
    
    results.responses.forEach(response => {
      response.answers.forEach(answer => {
        if (!questionStats[answer.question]) {
          questionStats[answer.question] = {};
        }
        
        const answerKey = answer.answer.toString();
        if (!questionStats[answer.question][answerKey]) {
          questionStats[answer.question][answerKey] = 0;
        }
        questionStats[answer.question][answerKey]++;
      });
    });
    
    // 統計情報をテキストに変換
    Object.keys(questionStats).forEach(question => {
      summaryText += `\n【${question}】\n`;
      Object.keys(questionStats[question]).forEach(answer => {
        const count = questionStats[question][answer];
        const percentage = Math.round((count / results.responseCount) * 100);
        summaryText += `- ${answer}: ${count}件 (${percentage}%)\n`;
      });
    });
  }
  
  return `PTA関係者の皆様

いつもPTA活動にご協力いただき、ありがとうございます。

先日お願いいたしましたアンケートの結果をご報告いたします。

【アンケートタイトル】
${survey.title}

【実施期間】
${Utilities.formatDate(survey.createdDate, 'Asia/Tokyo', 'yyyy年MM月dd日')} 〜 ${dueDateStr}

【回答数】
${results.responseCount}件

【結果概要】${summaryText}

詳細な結果については、別途資料をご確認ください。
貴重なご意見をいただき、ありがとうございました。

今後ともPTA活動へのご協力をよろしくお願いいたします。`;
}