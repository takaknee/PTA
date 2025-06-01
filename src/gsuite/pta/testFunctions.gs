/**
 * PTAシステムの基本的なテスト
 * このファイルは開発・デバッグ用であり、本番環境では不要です
 */

/**
 * 基本機能のテスト
 */
function testBasicFunctionality() {
  console.log('=== PTA情報配信システム 基本機能テスト ===');
  
  try {
    // 設定の確認
    console.log('設定確認:', PTA_CONFIG);
    
    // ログ機能のテスト
    logInfo('テストログ: 情報レベル');
    logError('テストログ: エラーレベル', new Error('テストエラー'));
    
    // メンバー統計の確認（シートが存在しない場合でもエラーにならないことを確認）
    const stats = getMemberStatistics();
    console.log('メンバー統計:', stats);
    
    // メールテンプレート取得のテスト
    const templates = getEmailTemplates();
    console.log('メールテンプレート数:', templates.length);
    
    // アンケート取得のテスト
    const surveys = getAllSurveys();
    console.log('アンケート数:', surveys.length);
    
    // 定期配信設定の取得テスト
    const distributions = getScheduledDistributions();
    console.log('定期配信設定数:', distributions.length);
    
    console.log('=== 基本機能テスト完了 ===');
    
  } catch (error) {
    console.error('テスト中にエラーが発生しました:', error);
  }
}

/**
 * メール形式のテスト
 */
function testEmailValidation() {
  console.log('=== メール形式テスト ===');
  
  const testEmails = [
    'test@example.com',
    'invalid-email',
    'user.name+tag@domain.co.jp',
    '@invalid.com',
    'test@',
    ''
  ];
  
  testEmails.forEach(email => {
    const isValid = isValidEmail(email);
    console.log(`${email}: ${isValid ? '有効' : '無効'}`);
  });
  
  console.log('=== メール形式テスト完了 ===');
}

/**
 * 設定値の確認
 */
function testConfiguration() {
  console.log('=== 設定値確認 ===');
  
  // 必須設定の確認
  const requiredConfigs = [
    'MEMBER_SHEET_NAME',
    'DISTRIBUTION_HISTORY_SHEET_NAME', 
    'SURVEY_SHEET_NAME',
    'LOG_SHEET_NAME',
    'MAX_EMAIL_BATCH_SIZE'
  ];
  
  requiredConfigs.forEach(config => {
    if (PTA_CONFIG[config]) {
      console.log(`${config}: OK (${PTA_CONFIG[config]})`);
    } else {
      console.error(`${config}: 設定されていません`);
    }
  });
  
  console.log('=== 設定値確認完了 ===');
}

/**
 * 全テストの実行
 */
function runAllTests() {
  console.log('PTA情報配信システム - 全テスト開始');
  
  testConfiguration();
  testEmailValidation(); 
  testBasicFunctionality();
  
  console.log('PTA情報配信システム - 全テスト完了');
}

/**
 * デモデータの作成（開発・テスト用）
 */
function createDemoData() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'デモデータ作成',
    'デモ用のサンプルデータを作成しますか？\n（テスト用途のみ）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // 初期設定を実行
    initializePTASystem();
    
    // サンプルメンバーを追加
    const sampleMembers = [
      { name: 'テスト太郎', email: 'test.taro@example.com' },
      { name: 'サンプル花子', email: 'sample.hanako@example.com' },
      { name: 'デモ次郎', email: 'demo.jiro@example.com' }
    ];
    
    sampleMembers.forEach(member => {
      addMemberToSheet(member.name, member.email);
    });
    
    // サンプルテンプレートを作成
    saveEmailTemplate(
      'お知らせテンプレート',
      '【PTA】重要なお知らせ',
      'PTA関係者の皆様\n\nいつもお世話になっております。\n重要なお知らせがあります。\n\nよろしくお願いいたします。'
    );
    
    ui.alert(
      '作成完了',
      'デモデータの作成が完了しました。\n\n- サンプルメンバー: 3名\n- サンプルテンプレート: 1件',
      ui.ButtonSet.OK
    );
    
    logInfo('デモデータを作成しました');
    
  } catch (error) {
    logError('デモデータ作成中にエラーが発生しました', error);
    ui.alert(
      'エラー',
      `デモデータ作成中にエラーが発生しました: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}