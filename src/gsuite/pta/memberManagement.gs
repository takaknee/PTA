/**
 * PTAメンバー管理機能
 * メンバーの追加、更新、削除、検索機能を提供します
 */

/**
 * メンバー一覧を表示する
 */
function showMemberList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
    
    if (!memberSheet) {
      SpreadsheetApp.getUi().alert(
        'エラー',
        'メンバーシートが見つかりません。先に初期設定を実行してください。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // メンバーシートをアクティブにする
    ss.setActiveSheet(memberSheet);
    
    // 統計情報を取得
    const stats = getMemberStatistics();
    
    SpreadsheetApp.getUi().alert(
      'メンバー一覧',
      `メンバー一覧を表示しました。\n\n統計情報:\n- 総メンバー数: ${stats.total}人\n- アクティブ: ${stats.active}人\n- 非アクティブ: ${stats.inactive}人\n- 退会済み: ${stats.withdrawn}人`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    logInfo('メンバー一覧を表示しました');
    
  } catch (error) {
    logError('メンバー一覧の表示中にエラーが発生しました', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `メンバー一覧の表示中にエラーが発生しました: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 新しいメンバーを追加する
 */
function addNewMember() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 氏名の入力
    const nameResponse = ui.prompt(
      '新しいメンバーを追加',
      '氏名を入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (nameResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const name = nameResponse.getResponseText().trim();
    if (!name) {
      ui.alert('エラー', '氏名を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // メールアドレスの入力
    const emailResponse = ui.prompt(
      '新しいメンバーを追加',
      'メールアドレスを入力してください:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (emailResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const email = emailResponse.getResponseText().trim();
    if (!email) {
      ui.alert('エラー', 'メールアドレスを入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // メールアドレス形式の検証
    if (!isValidEmail(email)) {
      ui.alert('エラー', '有効なメールアドレスを入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    // 重複チェック
    if (isEmailExists(email)) {
      ui.alert('エラー', 'このメールアドレスは既に登録されています。', ui.ButtonSet.OK);
      return;
    }
    
    // メンバーを追加
    const memberId = addMemberToSheet(name, email);
    
    ui.alert(
      '追加完了',
      `新しいメンバーを追加しました。\n\n氏名: ${name}\nメールアドレス: ${email}\nメンバーID: ${memberId}`,
      ui.ButtonSet.OK
    );
    
    logInfo(`新しいメンバーを追加しました: ${name} (${email})`);
    
  } catch (error) {
    logError('メンバー追加中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `メンバー追加中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * メンバーのステータスを更新する
 */
function updateMemberStatus() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // メンバー選択
    const member = selectMemberByEmail();
    if (!member) {
      return;
    }
    
    // 新しいステータスの選択
    const statusResponse = ui.prompt(
      'ステータス更新',
      `${member.name} (${member.email}) のステータスを選択してください:\n\n1: アクティブ\n2: 非アクティブ\n3: 退会\n\n番号を入力してください:`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (statusResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const statusChoice = statusResponse.getResponseText().trim();
    let newStatus;
    
    switch (statusChoice) {
      case '1':
        newStatus = 'アクティブ';
        break;
      case '2':
        newStatus = '非アクティブ';
        break;
      case '3':
        newStatus = '退会';
        break;
      default:
        ui.alert('エラー', '有効な番号を入力してください。', ui.ButtonSet.OK);
        return;
    }
    
    // ステータスを更新
    updateMemberStatusInSheet(member.row, newStatus);
    
    ui.alert(
      '更新完了',
      `${member.name} のステータスを「${newStatus}」に更新しました。`,
      ui.ButtonSet.OK
    );
    
    logInfo(`メンバーステータスを更新しました: ${member.name} -> ${newStatus}`);
    
  } catch (error) {
    logError('メンバーステータス更新中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `ステータス更新中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * メンバーリストをCSVファイルとしてエクスポートする
 */
function exportMemberList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
    
    if (!memberSheet) {
      SpreadsheetApp.getUi().alert(
        'エラー',
        'メンバーシートが見つかりません。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // データの取得
    const lastRow = memberSheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert(
        '情報',
        'エクスポートするメンバーデータがありません。',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const data = memberSheet.getRange(1, 1, lastRow, 7).getValues();
    
    // CSVファイルを作成
    const csvContent = data.map(row => 
      row.map(cell => `"${cell.toString().replace(/"/g, '""')}"`).join(',')
    ).join('\n');
    
    // ファイル名の生成
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const fileName = `PTA_Members_${timestamp}.csv`;
    
    // ドライブにファイルを作成
    const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    const file = DriveApp.createFile(blob);
    
    SpreadsheetApp.getUi().alert(
      'エクスポート完了',
      `メンバーリストをエクスポートしました。\n\nファイル名: ${fileName}\nファイルID: ${file.getId()}\n\nGoogleドライブで確認できます。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    logInfo(`メンバーリストをエクスポートしました: ${fileName}`);
    
  } catch (error) {
    logError('メンバーリストエクスポート中にエラーが発生しました', error);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `エクスポート中にエラーが発生しました: ${error.toString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * メンバー統計情報を取得する
 * @returns {Object} 統計情報
 */
function getMemberStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
  
  if (!memberSheet) {
    return { total: 0, active: 0, inactive: 0, withdrawn: 0 };
  }
  
  const lastRow = memberSheet.getLastRow();
  if (lastRow <= 1) {
    return { total: 0, active: 0, inactive: 0, withdrawn: 0 };
  }
  
  // ステータス列のデータを取得（4列目）
  const statusData = memberSheet.getRange(2, 4, lastRow - 1, 1).getValues();
  
  const stats = {
    total: statusData.length,
    active: 0,
    inactive: 0,
    withdrawn: 0
  };
  
  statusData.forEach(row => {
    const status = row[0];
    switch (status) {
      case 'アクティブ':
        stats.active++;
        break;
      case '非アクティブ':
        stats.inactive++;
        break;
      case '退会':
        stats.withdrawn++;
        break;
    }
  });
  
  return stats;
}

/**
 * メールアドレスでメンバーを選択する
 * @returns {Object|null} 選択されたメンバー情報
 */
function selectMemberByEmail() {
  const ui = SpreadsheetApp.getUi();
  
  const emailResponse = ui.prompt(
    'メンバー選択',
    'メールアドレスを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  const email = emailResponse.getResponseText().trim();
  if (!email) {
    ui.alert('エラー', 'メールアドレスを入力してください。', ui.ButtonSet.OK);
    return null;
  }
  
  const member = findMemberByEmail(email);
  if (!member) {
    ui.alert('エラー', 'このメールアドレスのメンバーが見つかりません。', ui.ButtonSet.OK);
    return null;
  }
  
  return member;
}

/**
 * メールアドレスでメンバーを検索する
 * @param {string} email - メールアドレス
 * @returns {Object|null} メンバー情報
 */
function findMemberByEmail(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
  
  if (!memberSheet) {
    return null;
  }
  
  const lastRow = memberSheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }
  
  const data = memberSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][2] === email) { // メールアドレスは3列目（インデックス2）
      return {
        row: i + 2, // シートの行番号（ヘッダー行を考慮）
        id: data[i][0],
        name: data[i][1],
        email: data[i][2],
        status: data[i][3],
        joinDate: data[i][4],
        leaveDate: data[i][5],
        notes: data[i][6]
      };
    }
  }
  
  return null;
}

/**
 * メールアドレスの重複をチェックする
 * @param {string} email - メールアドレス
 * @returns {boolean} 既に存在する場合はtrue
 */
function isEmailExists(email) {
  return findMemberByEmail(email) !== null;
}

/**
 * メールアドレスの形式を検証する
 * @param {string} email - メールアドレス
 * @returns {boolean} 有効な形式の場合はtrue
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * メンバーをシートに追加する
 * @param {string} name - 氏名
 * @param {string} email - メールアドレス
 * @returns {string} 生成されたメンバーID
 */
function addMemberToSheet(name, email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
  
  if (!memberSheet) {
    throw new Error('メンバーシートが見つかりません');
  }
  
  // 新しいIDを生成
  const lastRow = memberSheet.getLastRow();
  const memberId = `PTA${String(lastRow).padStart(4, '0')}`;
  
  // 現在の日時
  const joinDate = new Date();
  
  // 新しい行を追加
  memberSheet.appendRow([
    memberId,     // ID
    name,         // 氏名
    email,        // メールアドレス
    'アクティブ',   // ステータス
    joinDate,     // 登録日
    '',           // 退会日
    ''            // 備考
  ]);
  
  return memberId;
}

/**
 * メンバーのステータスをシートで更新する
 * @param {number} row - 行番号
 * @param {string} status - 新しいステータス
 */
function updateMemberStatusInSheet(row, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
  
  if (!memberSheet) {
    throw new Error('メンバーシートが見つかりません');
  }
  
  // ステータスを更新（4列目）
  memberSheet.getRange(row, 4).setValue(status);
  
  // 退会の場合は退会日を設定（6列目）
  if (status === '退会') {
    memberSheet.getRange(row, 6).setValue(new Date());
  } else {
    // 退会以外の場合は退会日をクリア
    memberSheet.getRange(row, 6).setValue('');
  }
}

/**
 * アクティブなメンバーのリストを取得する
 * @returns {Array} アクティブなメンバーの配列
 */
function getActiveMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(PTA_CONFIG.MEMBER_SHEET_NAME);
  
  if (!memberSheet) {
    return [];
  }
  
  const lastRow = memberSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  
  const data = memberSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  return data
    .filter(row => row[3] === 'アクティブ') // ステータスがアクティブ
    .map(row => ({
      id: row[0],
      name: row[1],
      email: row[2],
      status: row[3],
      joinDate: row[4],
      leaveDate: row[5],
      notes: row[6]
    }));
}