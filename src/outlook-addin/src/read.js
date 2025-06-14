/*
 * PTA Outlook アドイン - メール読み取り機能
 * Copyright (c) 2024 PTA Development Team
 */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.addEventListener("DOMContentLoaded", function() {
            // ボタンイベント設定
            document.getElementById("analyze-button").onclick = analyzeCurrentEmail;
            document.getElementById("taskpane-button").onclick = openTaskPane;
        });
    }
});

/**
 * 現在のメールをAIで解析する
 */
async function analyzeCurrentEmail() {
    try {
        showLoading(true);
        
        // メール情報取得
        const emailData = await getEmailData();
        
        // AI解析実行
        const analysis = await performAIAnalysis(emailData);
        
        // 結果表示
        displayAnalysisResult(analysis);
        
    } catch (error) {
        console.error('メール解析エラー:', error);
        showError('メール解析中にエラーが発生しました: ' + error.message);
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
 * AI解析を実行
 */
async function performAIAnalysis(emailData) {
    // 設定読み込み
    const settings = await loadSettings();
    
    const prompt = `
以下のメールを解析し、重要な情報を日本語で抽出してください：

件名: ${emailData.subject}
送信者: ${emailData.sender}
本文:
${emailData.body}

以下の形式で回答してください：
1. 要約
2. 重要なポイント
3. 必要なアクション
4. 優先度（高/中/低）
`;

    try {
        const response = await callAIAPI(prompt, settings);
        return {
            summary: response,
            timestamp: new Date().toLocaleString('ja-JP'),
            emailSubject: emailData.subject
        };
    } catch (error) {
        throw new Error('AI APIエラー: ' + error.message);
    }
}

/**
 * AI APIを呼び出し
 */
async function callAIAPI(prompt, settings) {
    const apiKey = settings.apiKey;
    const provider = settings.provider || 'openai';
    const model = settings.model || 'gpt-3.5-turbo';
    
    if (!apiKey) {
        throw new Error('APIキーが設定されていません。設定タブでAPIキーを設定してください。');
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
                max_tokens: 1000,
                temperature: 0.7
            });
            break;
            
        case 'azure':
            // Azure OpenAI用の設定
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
                max_tokens: 1000,
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
 * 解析結果を表示
 */
function displayAnalysisResult(analysis) {
    const resultContainer = document.getElementById('result-container');
    const resultDiv = document.getElementById('analysis-result');
    
    resultDiv.innerHTML = `
        <div class="analysis-result-item">
            <h4>解析結果 (${analysis.timestamp})</h4>
            <div class="result-content">
                ${analysis.summary.replace(/\n/g, '<br>')}
            </div>
        </div>
    `;
    
    resultContainer.style.display = 'block';
    
    // 履歴に保存
    saveToHistory(analysis);
}

/**
 * 履歴に保存
 */
function saveToHistory(analysis) {
    Office.context.roamingSettings.get('ptaHistory', (historyData) => {
        let history = [];
        if (historyData) {
            history = JSON.parse(historyData);
        }
        
        history.unshift({
            id: Date.now(),
            type: 'analysis',
            timestamp: analysis.timestamp,
            subject: analysis.emailSubject,
            result: analysis.summary
        });
        
        // 最新50件のみ保持
        if (history.length > 50) {
            history = history.slice(0, 50);
        }
        
        Office.context.roamingSettings.set('ptaHistory', JSON.stringify(history));
        Office.context.roamingSettings.saveAsync();
    });
}

/**
 * タスクペインを開く
 */
function openTaskPane() {
    Office.ribbon.requestUpdate({
        tabs: [{
            id: "TabDefault",
            controls: [{
                id: "msgReadOpenPaneButton",
                enabled: true
            }]
        }]
    });
}

/**
 * ローディング表示制御
 */
function showLoading(show) {
    const loadingContainer = document.getElementById('loading-container');
    const controlsContainer = document.querySelector('.controls-container');
    
    if (show) {
        loadingContainer.style.display = 'block';
        controlsContainer.style.display = 'none';
    } else {
        loadingContainer.style.display = 'none';
        controlsContainer.style.display = 'block';
    }
}

/**
 * エラー表示
 */
function showError(message) {
    const notificationContainer = document.querySelector('.notification-container');
    notificationContainer.innerHTML = `
        <div class="error-message">
            <p class="ms-font-m" style="color: #d13438;">エラー: ${message}</p>
        </div>
    `;
}