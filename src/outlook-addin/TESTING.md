# Office Add-in テストファイル

## 概要

PTA Outlook Add-inの機能をテストするためのファイルです。

## テスト方法

### 1. 基本機能テスト

```javascript
// ブラウザの開発者ツールのコンソールで実行

// Office APIの初期化確認
console.log('Office ready:', Office.isSetSupported('Mailbox', '1.1'));

// 設定機能テスト
testSettings();

// メール解析テスト（メール選択が必要）
testEmailAnalysis();

// AI API接続テスト
testAPIConnection();
```

### 2. API統合テスト

#### OpenAI API テスト

```javascript
async function testOpenAIAPI() {
    const settings = {
        provider: 'openai',
        apiKey: 'your-api-key-here', // 実際のAPIキーに置換
        model: 'gpt-3.5-turbo'
    };
    
    const testPrompt = '「こんにちは」と日本語で返答してください。';
    
    try {
        const result = await callAIAPI(testPrompt, settings);
        console.log('OpenAI API テスト成功:', result);
        return true;
    } catch (error) {
        console.error('OpenAI API テスト失敗:', error);
        return false;
    }
}
```

#### Azure OpenAI API テスト

```javascript
async function testAzureOpenAIAPI() {
    const settings = {
        provider: 'azure',
        apiKey: 'your-azure-api-key',
        model: 'gpt-35-turbo',
        azureEndpoint: 'https://your-resource.openai.azure.com'
    };
    
    const testPrompt = '「接続成功」と日本語で返答してください。';
    
    try {
        const result = await callAIAPI(testPrompt, settings);
        console.log('Azure OpenAI API テスト成功:', result);
        return true;
    } catch (error) {
        console.error('Azure OpenAI API テスト失敗:', error);
        return false;
    }
}
```

### 3. UI機能テスト

#### タブ切り替えテスト

```javascript
function testTabSwitching() {
    const tabs = ['analyze', 'settings', 'history'];
    
    tabs.forEach((tab, index) => {
        setTimeout(() => {
            const tabHeader = document.querySelector(`[data-tab="${tab}"]`);
            if (tabHeader) {
                tabHeader.click();
                console.log(`タブ切り替えテスト: ${tab} タブ`);
            }
        }, index * 1000);
    });
}
```

#### フォーム入力テスト

```javascript
function testFormInputs() {
    const testData = {
        'analysis-type': 'summary',
        'custom-prompt': 'このメールの重要なポイントを教えてください。',
        'api-provider': 'openai',
        'model': 'gpt-3.5-turbo',
        'email-type': 'announcement',
        'message-content': 'PTA総会の開催についてお知らせします。',
        'tone': 'formal'
    };
    
    Object.entries(testData).forEach(([id, value]) => {
        const element = document.getElementById(id);
        if (element) {
            element.value = value;
            console.log(`フォーム入力テスト: ${id} = ${value}`);
        }
    });
}
```

### 4. エラーハンドリングテスト

#### 無効なAPIキーテスト

```javascript
async function testInvalidAPIKey() {
    const settings = {
        provider: 'openai',
        apiKey: 'invalid-key',
        model: 'gpt-3.5-turbo'
    };
    
    try {
        await callAIAPI('テストプロンプト', settings);
        console.error('エラーハンドリングテスト失敗: エラーが発生するべき');
        return false;
    } catch (error) {
        console.log('エラーハンドリングテスト成功:', error.message);
        return true;
    }
}
```

#### ネットワークエラーテスト

```javascript
async function testNetworkError() {
    const settings = {
        provider: 'openai',
        apiKey: 'test-key',
        model: 'gpt-3.5-turbo'
    };
    
    // 無効なエンドポイントでテスト
    const originalFetch = window.fetch;
    window.fetch = () => Promise.reject(new Error('Network error'));
    
    try {
        await callAIAPI('テストプロンプト', settings);
        console.error('ネットワークエラーテスト失敗');
        return false;
    } catch (error) {
        console.log('ネットワークエラーテスト成功:', error.message);
        return true;
    } finally {
        window.fetch = originalFetch;
    }
}
```

### 5. パフォーマンステスト

#### レスポンス時間測定

```javascript
async function measureResponseTime(operation, fn) {
    const start = performance.now();
    
    try {
        const result = await fn();
        const duration = performance.now() - start;
        console.log(`${operation} 実行時間: ${duration.toFixed(2)}ms`);
        return { success: true, duration, result };
    } catch (error) {
        const duration = performance.now() - start;
        console.log(`${operation} エラー時間: ${duration.toFixed(2)}ms`);
        return { success: false, duration, error };
    }
}

// 使用例
measureResponseTime('メール解析', () => analyzeCurrentEmail());
measureResponseTime('メール生成', () => generateEmailContent());
```

#### メモリ使用量チェック

```javascript
function checkMemoryUsage() {
    if (performance.memory) {
        const memory = performance.memory;
        console.log('メモリ使用状況:');
        console.log(`使用中: ${(memory.usedJSHeapSize / 1024 / 1024).toFixed(2)} MB`);
        console.log(`制限: ${(memory.jsHeapSizeLimit / 1024 / 1024).toFixed(2)} MB`);
        console.log(`割り当て済み: ${(memory.totalJSHeapSize / 1024 / 1024).toFixed(2)} MB`);
    } else {
        console.log('メモリ情報は利用できません');
    }
}
```

### 6. 設定テスト

#### 設定保存・読み込みテスト

```javascript
async function testSettingsPersistence() {
    const testSettings = {
        provider: 'openai',
        apiKey: 'test-key-12345',
        model: 'gpt-4',
        customEndpoint: 'https://test.example.com'
    };
    
    // 設定保存
    Office.context.roamingSettings.set('ptaSettings', JSON.stringify(testSettings));
    Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('設定保存テスト成功');
            
            // 設定読み込み
            Office.context.roamingSettings.get('ptaSettings', (savedSettings) => {
                const parsed = JSON.parse(savedSettings);
                const isMatch = JSON.stringify(parsed) === JSON.stringify(testSettings);
                console.log('設定読み込みテスト:', isMatch ? '成功' : '失敗');
                console.log('保存された設定:', parsed);
            });
        } else {
            console.error('設定保存テスト失敗:', result.error);
        }
    });
}
```

### 7. 履歴テスト

#### 履歴保存・クリアテスト

```javascript
function testHistoryFunctions() {
    // テスト履歴データ
    const testHistory = [
        {
            id: Date.now(),
            type: 'analysis',
            timestamp: new Date().toLocaleString('ja-JP'),
            subject: 'テストメール1',
            result: 'テスト解析結果1'
        },
        {
            id: Date.now() + 1,
            type: 'generation',
            timestamp: new Date().toLocaleString('ja-JP'),
            emailType: 'announcement',
            result: 'テスト生成結果2'
        }
    ];
    
    // 履歴保存
    Office.context.roamingSettings.set('ptaHistory', JSON.stringify(testHistory));
    Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('履歴保存テスト成功');
            
            // 履歴読み込み
            loadHistoryUI();
            
            // 履歴クリア
            setTimeout(() => {
                clearHistory();
                console.log('履歴クリアテスト実行');
            }, 2000);
        }
    });
}
```

### 8. 統合テスト

#### 全機能テストスイート

```javascript
async function runFullTestSuite() {
    console.log('=== PTA Outlook Add-in 統合テスト開始 ===');
    
    const tests = [
        { name: '設定保存・読み込み', fn: testSettingsPersistence },
        { name: 'タブ切り替え', fn: testTabSwitching },
        { name: 'フォーム入力', fn: testFormInputs },
        { name: '履歴機能', fn: testHistoryFunctions },
        { name: '無効APIキー', fn: testInvalidAPIKey },
        { name: 'ネットワークエラー', fn: testNetworkError }
    ];
    
    const results = [];
    
    for (const test of tests) {
        try {
            console.log(`テスト実行中: ${test.name}`);
            const result = await test.fn();
            results.push({ name: test.name, success: true, result });
            console.log(`✅ ${test.name}: 成功`);
        } catch (error) {
            results.push({ name: test.name, success: false, error });
            console.error(`❌ ${test.name}: 失敗`, error);
        }
        
        // テスト間の待機
        await new Promise(resolve => setTimeout(resolve, 1000));
    }
    
    // 結果サマリー
    const successCount = results.filter(r => r.success).length;
    console.log(`\n=== テスト結果 ===`);
    console.log(`成功: ${successCount}/${tests.length}`);
    console.log(`失敗: ${tests.length - successCount}/${tests.length}`);
    
    return results;
}
```

## テスト実行手順

### 1. 開発環境でのテスト

```bash
# アドインの起動
cd src/outlook-addin
npm install
npm start

# ブラウザでテストページを開く
# https://localhost:3000/src/taskpane.html
```

### 2. Outlookでのテスト

1. manifest.xmlをOutlookにサイドロード
2. メールを選択して各機能をテスト
3. ブラウザの開発者ツールでコンソールを確認

### 3. 自動テストの実行

ブラウザのコンソールで実行：

```javascript
// 全テスト実行
runFullTestSuite();

// 個別テスト実行
testSettingsPersistence();
testAPIConnection();
```

## 期待される結果

### 正常系

- API呼び出しが成功し、適切な応答が返る
- 設定が正しく保存・読み込みされる
- 履歴が適切に管理される
- UI要素が正常に動作する

### 異常系

- 無効なAPIキーでエラーが適切にハンドリングされる
- ネットワークエラーが適切に処理される
- フォームバリデーションが機能する

## トラブルシューティング

### よくある問題

1. **Office APIエラー**: Officeが初期化されていない
2. **CORSエラー**: HTTPSでない、または許可されていないドメイン
3. **認証エラー**: APIキーが無効または期限切れ

### デバッグのヒント

- ブラウザの開発者ツールのNetworkタブでAPIリクエストを確認
- Consoleタブでエラーメッセージを確認
- Office.context.diagnostics で詳細情報を取得