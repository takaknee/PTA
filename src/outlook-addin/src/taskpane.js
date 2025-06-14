/*
 * PTA Outlook アドイン - タスクペイン機能
 * Copyright (c) 2024 PTA Development Team
 */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.addEventListener("DOMContentLoaded", function() {
            initializeTaskPane();
        });
    }
});

/**
 * タスクペインの初期化
 */
function initializeTaskPane() {
    // タブ切り替えイベント
    const tabHeaders = document.querySelectorAll('.tab-header');
    tabHeaders.forEach(header => {
        header.addEventListener('click', switchTab);
    });
    
    // 解析タブのイベント
    document.getElementById('analysis-type').onchange = handleAnalysisTypeChange;
    document.getElementById('analyze-email-button').onclick = analyzeEmail;
    
    // 設定タブのイベント
    document.getElementById('api-provider').onchange = handleProviderChange;
    document.getElementById('save-settings-button').onclick = saveSettings;
    document.getElementById('test-api-button').onclick = testAPIConnection;
    
    // 履歴タブのイベント
    document.getElementById('clear-history-button').onclick = clearHistory;
    
    // 初期データ読み込み
    loadSettingsUI();
    loadHistoryUI();
}

/**
 * タブ切り替え
 */
function switchTab(event) {
    const targetTab = event.target.getAttribute('data-tab');
    
    // すべてのタブヘッダーとコンテンツを非アクティブに
    document.querySelectorAll('.tab-header').forEach(h => h.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    // 選択されたタブをアクティブに
    event.target.classList.add('active');
    document.getElementById(targetTab + '-tab').classList.add('active');
}

/**
 * 解析タイプ変更時の処理
 */
function handleAnalysisTypeChange() {
    const analysisType = document.getElementById('analysis-type').value;
    const customPromptGroup = document.getElementById('custom-prompt-group');
    
    if (analysisType === 'custom') {
        customPromptGroup.style.display = 'block';
    } else {
        customPromptGroup.style.display = 'none';
    }
}

/**
 * メール解析実行
 */
async function analyzeEmail() {
    try {
        showLoading(true);
        
        // メールデータ取得
        const emailData = await getEmailData();
        
        // 解析タイプ取得
        const analysisType = document.getElementById('analysis-type').value;
        const customPrompt = document.getElementById('custom-prompt').value;
        
        // プロンプト生成
        const prompt = createAnalysisPrompt(emailData, analysisType, customPrompt);
        
        // AI解析実行
        const settings = await loadSettings();
        const result = await callAIAPI(prompt, settings);
        
        // 結果表示
        displayAnalysisResults(result, analysisType);
        
        // 履歴保存
        saveAnalysisToHistory(emailData, analysisType, result);
        
    } catch (error) {
        console.error('解析エラー:', error);
        showError('解析中にエラーが発生しました: ' + error.message);
    } finally {
        showLoading(false);
    }
}

/**
 * メールデータを取得
 */
function getEmailData() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const emailData = {
                    subject: Office.context.mailbox.item.subject,
                    body: result.value,
                    sender: Office.context.mailbox.item.sender ? Office.context.mailbox.item.sender.emailAddress : '',
                    dateTime: Office.context.mailbox.item.dateTimeCreated
                };
                resolve(emailData);
            } else {
                reject(new Error(result.error.message));
            }
        });
    });
}

/**
 * 解析用プロンプトを作成
 */
function createAnalysisPrompt(emailData, analysisType, customPrompt) {
    let specificInstruction = '';
    
    switch (analysisType) {
        case 'summary':
            specificInstruction = `
このメールの要約を以下の形式で作成してください：
- 主な内容（2-3行）
- 重要なポイント（箇条書き）
- 結論や次のステップ
`;
            break;
            
        case 'action-items':
            specificInstruction = `
このメールから必要なアクション項目を抽出してください：
- 実行すべきタスク
- 期限や締切
- 担当者や連絡先
- 優先度
`;
            break;
            
        case 'sentiment':
            specificInstruction = `
このメールの感情分析を行ってください：
- 送信者の感情や態度
- メッセージのトーン
- 緊急度や重要度の評価
- 適切な返信トーンの提案
`;
            break;
            
        case 'priority':
            specificInstruction = `
このメールの優先度を評価してください：
- 優先度レベル（高/中/低）
- 評価理由
- 対応の推奨タイミング
- 関連する人やプロジェクト
`;
            break;
            
        case 'custom':
            specificInstruction = customPrompt || 'このメールについて分析してください。';
            break;
    }
    
    return `
あなたはPTA活動を支援するAIアシスタントです。以下のメールを分析し、日本語で回答してください。

件名: ${emailData.subject}
送信者: ${emailData.sender}
日時: ${emailData.dateTime}

本文:
${emailData.body}

分析指示:
${specificInstruction}

PTA活動の文脈を考慮し、実用的で具体的な分析を提供してください。
`;
}

/**
 * 解析結果を表示
 */
function displayAnalysisResults(result, analysisType) {
    const resultsContainer = document.getElementById('analysis-results');
    const timestamp = new Date().toLocaleString('ja-JP');
    
    resultsContainer.innerHTML = `
        <div class="analysis-result-item">
            <h4>${getAnalysisTypeLabel(analysisType)} (${timestamp})</h4>
            <div class="result-content">
                <pre style="white-space: pre-wrap; font-family: inherit;">${result}</pre>
            </div>
        </div>
    `;
}

/**
 * 解析タイプのラベルを取得
 */
function getAnalysisTypeLabel(type) {
    const labels = {
        'summary': '要約',
        'action-items': 'アクション項目',
        'sentiment': '感情分析',
        'priority': '優先度評価',
        'custom': 'カスタム解析'
    };
    return labels[type] || '解析結果';
}

/**
 * プロバイダー変更時の処理
 */
function handleProviderChange() {
    const provider = document.getElementById('api-provider').value;
    const modelSelect = document.getElementById('model');
    
    // モデル選択肢を更新
    modelSelect.innerHTML = '';
    let options = [];
    
    switch (provider) {
        case 'openai':
            options = [
                { value: 'gpt-4', text: 'GPT-4' },
                { value: 'gpt-3.5-turbo', text: 'GPT-3.5 Turbo' }
            ];
            break;
        case 'azure':
            options = [
                { value: 'gpt-4', text: 'GPT-4' },
                { value: 'gpt-35-turbo', text: 'GPT-3.5 Turbo' }
            ];
            break;
        case 'claude':
            options = [
                { value: 'claude-3-opus', text: 'Claude 3 Opus' },
                { value: 'claude-3-sonnet', text: 'Claude 3 Sonnet' }
            ];
            break;
    }
    
    options.forEach(option => {
        const optionElement = document.createElement('option');
        optionElement.value = option.value;
        optionElement.textContent = option.text;
        modelSelect.appendChild(optionElement);
    });
}

/**
 * 設定を保存
 */
async function saveSettings() {
    try {
        const settings = {
            provider: document.getElementById('api-provider').value,
            apiKey: document.getElementById('api-key').value,
            model: document.getElementById('model').value
        };
        
        // Azure用の追加設定
        if (settings.provider === 'azure') {
            settings.azureEndpoint = prompt('Azure OpenAI エンドポイントを入力してください:') || '';
        }
        
        // 設定保存
        Office.context.roamingSettings.set('ptaSettings', JSON.stringify(settings));
        Office.context.roamingSettings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showSuccess('設定を保存しました。');
            } else {
                showError('設定の保存に失敗しました: ' + result.error.message);
            }
        });
        
    } catch (error) {
        console.error('設定保存エラー:', error);
        showError('設定保存中にエラーが発生しました: ' + error.message);
    }
}

/**
 * API接続テスト
 */
async function testAPIConnection() {
    try {
        showLoading(true);
        
        const settings = await loadSettings();
        
        if (!settings.apiKey) {
            showError('APIキーが設定されていません。');
            return;
        }
        
        // テスト用プロンプト
        const testPrompt = 'こんにちは。接続テストです。「接続成功」と日本語で返答してください。';
        
        const response = await callAIAPI(testPrompt, settings);
        
        if (response && response.includes('接続成功')) {
            showSuccess('API接続テストが成功しました。');
        } else {
            showSuccess('API接続は成功しましたが、予期しない応答です: ' + response);
        }
        
    } catch (error) {
        console.error('接続テストエラー:', error);
        showError('接続テストに失敗しました: ' + error.message);
    } finally {
        showLoading(false);
    }
}

/**
 * 設定UIを読み込み
 */
async function loadSettingsUI() {
    try {
        const settings = await loadSettings();
        
        document.getElementById('api-provider').value = settings.provider || 'openai';
        document.getElementById('api-key').value = settings.apiKey || '';
        
        // プロバイダー変更を発火してモデル選択肢を更新
        handleProviderChange();
        
        document.getElementById('model').value = settings.model || 'gpt-3.5-turbo';
        
    } catch (error) {
        console.error('設定読み込みエラー:', error);
    }
}

/**
 * 履歴UIを読み込み
 */
function loadHistoryUI() {
    Office.context.roamingSettings.get('ptaHistory', (historyData) => {
        const historyList = document.getElementById('history-list');
        
        if (!historyData) {
            historyList.innerHTML = '<p class="ms-font-m">履歴はまだありません。</p>';
            return;
        }
        
        const history = JSON.parse(historyData);
        
        if (history.length === 0) {
            historyList.innerHTML = '<p class="ms-font-m">履歴はまだありません。</p>';
            return;
        }
        
        let historyHTML = '';
        history.forEach(item => {
            const typeLabel = item.type === 'analysis' ? '解析' : '生成';
            historyHTML += `
                <div class="history-item">
                    <div class="history-header">
                        <span class="history-type">${typeLabel}</span>
                        <span class="history-timestamp">${item.timestamp}</span>
                    </div>
                    <div class="history-subject">${item.subject || item.emailType || '不明'}</div>
                    <div class="history-preview">${truncateText(item.result, 100)}</div>
                </div>
            `;
        });
        
        historyList.innerHTML = historyHTML;
    });
}

/**
 * 履歴をクリア
 */
function clearHistory() {
    if (confirm('履歴をすべて削除しますか？')) {
        Office.context.roamingSettings.set('ptaHistory', JSON.stringify([]));
        Office.context.roamingSettings.saveAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                loadHistoryUI();
                showSuccess('履歴をクリアしました。');
            } else {
                showError('履歴のクリアに失敗しました: ' + result.error.message);
            }
        });
    }
}

/**
 * 解析を履歴に保存
 */
function saveAnalysisToHistory(emailData, analysisType, result) {
    Office.context.roamingSettings.get('ptaHistory', (historyData) => {
        let history = [];
        if (historyData) {
            history = JSON.parse(historyData);
        }
        
        history.unshift({
            id: Date.now(),
            type: 'analysis',
            timestamp: new Date().toLocaleString('ja-JP'),
            subject: emailData.subject,
            analysisType: analysisType,
            result: result
        });
        
        // 最新50件のみ保持
        if (history.length > 50) {
            history = history.slice(0, 50);
        }
        
        Office.context.roamingSettings.set('ptaHistory', JSON.stringify(history));
        Office.context.roamingSettings.saveAsync();
    });
}

// 共通関数（他のファイルから再利用）

/**
 * AI APIを呼び出し
 */
async function callAIAPI(prompt, settings) {
    const apiKey = settings.apiKey;
    const provider = settings.provider || 'openai';
    const model = settings.model || 'gpt-3.5-turbo';
    
    if (!apiKey) {
        throw new Error('APIキーが設定されていません。');
    }
    
    let endpoint, headers, body;
    
    switch (provider) {
        case 'openai':
            endpoint = 'https://api.openai.com/v1/chat/completions';
            headers = {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            };
            body = JSON.stringify({
                model: model,
                messages: [
                    {
                        role: 'system',
                        content: 'あなたはPTA活動を支援するAIアシスタントです。日本語で丁寧に回答してください。'
                    },
                    {
                        role: 'user',
                        content: prompt
                    }
                ],
                max_tokens: 1500,
                temperature: 0.7
            });
            break;
            
        case 'azure':
            const azureEndpoint = settings.azureEndpoint || '';
            endpoint = `${azureEndpoint}/openai/deployments/${model}/chat/completions?api-version=2023-05-15`;
            headers = {
                'Content-Type': 'application/json',
                'api-key': apiKey
            };
            body = JSON.stringify({
                messages: [
                    {
                        role: 'system',
                        content: 'あなたはPTA活動を支援するAIアシスタントです。日本語で丁寧に回答してください。'
                    },
                    {
                        role: 'user',
                        content: prompt
                    }
                ],
                max_tokens: 1500,
                temperature: 0.7
            });
            break;
            
        default:
            throw new Error('サポートされていないAIプロバイダーです: ' + provider);
    }
    
    const response = await fetch(endpoint, {
        method: 'POST',
        headers: headers,
        body: body
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`APIエラー (${response.status}): ${errorText}`);
    }
    
    const data = await response.json();
    
    if (provider === 'openai' || provider === 'azure') {
        if (data.choices && data.choices.length > 0) {
            return data.choices[0].message.content.trim();
        } else {
            throw new Error('AIからの有効な応答が得られませんでした');
        }
    }
    
    throw new Error('予期しないAPI応答形式です');
}

/**
 * 設定を読み込み
 */
async function loadSettings() {
    return new Promise((resolve) => {
        Office.context.roamingSettings.get('ptaSettings', (result) => {
            if (result) {
                resolve(JSON.parse(result));
            } else {
                resolve({
                    provider: 'openai',
                    model: 'gpt-3.5-turbo',
                    apiKey: ''
                });
            }
        });
    });
}

/**
 * テキストを切り詰め
 */
function truncateText(text, maxLength) {
    if (text.length <= maxLength) {
        return text;
    }
    return text.substring(0, maxLength) + '...';
}

/**
 * ローディング表示制御
 */
function showLoading(show) {
    const loadingContainer = document.getElementById('loading-container');
    const tabsContainer = document.querySelector('.tabs-container');
    
    if (show) {
        loadingContainer.style.display = 'block';
        tabsContainer.style.display = 'none';
    } else {
        loadingContainer.style.display = 'none';
        tabsContainer.style.display = 'block';
    }
}

/**
 * エラー表示
 */
function showError(message) {
    // 簡単な通知表示（実際のUIでは適切な通知システムを使用）
    alert('エラー: ' + message);
}

/**
 * 成功メッセージ表示
 */
function showSuccess(message) {
    // 簡単な通知表示（実際のUIでは適切な通知システムを使用）
    alert('成功: ' + message);
}