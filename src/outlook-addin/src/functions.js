/*
 * PTA Outlook アドイン - 関数実行
 * Copyright (c) 2024 PTA Development Team
 */

Office.onReady(() => {
    // 関数定義の初期化
});

/**
 * メール解析関数（リボンボタンから呼び出し）
 */
function analyzeEmail(event) {
    try {
        // 現在のメールアイテムを取得
        const item = Office.context.mailbox.item;
        
        if (!item) {
            showNotification('メールが選択されていません。');
            event.completed();
            return;
        }
        
        // メール本文を取得
        item.body.getAsync(Office.CoercionType.Text, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                try {
                    // 設定読み込み
                    const settings = await loadSettings();
                    
                    if (!settings.apiKey) {
                        showNotification('APIキーが設定されていません。詳細パネルで設定してください。');
                        event.completed();
                        return;
                    }
                    
                    // AI解析実行
                    const emailData = {
                        subject: item.subject,
                        body: result.value,
                        sender: item.sender ? item.sender.emailAddress : '',
                        dateTime: item.dateTimeCreated
                    };
                    
                    const analysisResult = await performQuickAnalysis(emailData, settings);
                    
                    // 結果を通知として表示
                    showNotification(`解析完了: ${analysisResult.substring(0, 100)}...`);
                    
                    // 履歴に保存
                    saveToHistory({
                        id: Date.now(),
                        type: 'quick-analysis',
                        timestamp: new Date().toLocaleString('ja-JP'),
                        subject: emailData.subject,
                        result: analysisResult
                    });
                    
                } catch (error) {
                    console.error('解析エラー:', error);
                    showNotification('解析中にエラーが発生しました: ' + error.message);
                }
            } else {
                showNotification('メール内容の取得に失敗しました。');
            }
            
            event.completed();
        });
        
    } catch (error) {
        console.error('関数実行エラー:', error);
        showNotification('エラーが発生しました: ' + error.message);
        event.completed();
    }
}

/**
 * メール作成支援関数（リボンボタンから呼び出し）
 */
function assistCompose(event) {
    try {
        // 作成支援のダイアログを表示
        const dialogUrl = Office.context.requirements.isSetSupported('DialogApi', '1.1') ?
            'https://localhost:3000/src/compose-dialog.html' : null;
        
        if (dialogUrl) {
            Office.context.ui.displayDialogAsync(
                dialogUrl,
                { height: 60, width: 50 },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const dialog = result.value;
                        
                        // ダイアログからのメッセージを処理
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                            try {
                                const message = JSON.parse(args.message);
                                
                                if (message.type === 'generated-content') {
                                    // 生成された内容をメール本文に挿入
                                    insertContentToCompose(message.content);
                                }
                                
                                dialog.close();
                            } catch (e) {
                                console.error('ダイアログメッセージエラー:', e);
                            }
                        });
                    }
                }
            );
        } else {
            // ダイアログAPIがサポートされていない場合の代替処理
            showNotification('作成支援機能は詳細パネルから利用してください。');
        }
        
        event.completed();
        
    } catch (error) {
        console.error('作成支援エラー:', error);
        showNotification('作成支援中にエラーが発生しました: ' + error.message);
        event.completed();
    }
}

/**
 * クイック解析を実行
 */
async function performQuickAnalysis(emailData, settings) {
    const prompt = `
以下のメールを簡潔に解析してください：

件名: ${emailData.subject}
送信者: ${emailData.sender}
本文: ${emailData.body.substring(0, 1000)}...

以下の形式で回答してください：
- 要約（1-2行）
- 重要度（高/中/低）
- 必要なアクション（あれば）
`;

    return await callAIAPI(prompt, settings);
}

/**
 * 生成された内容をメール本文に挿入
 */
function insertContentToCompose(content) {
    const item = Office.context.mailbox.item;
    
    if (!item) {
        showNotification('メールアイテムが見つかりません。');
        return;
    }
    
    // 件名と本文を分離
    const lines = content.split('\n');
    let subject = '';
    let body = '';
    
    const subjectLine = lines.find(line => line.startsWith('件名:'));
    if (subjectLine) {
        subject = subjectLine.replace('件名:', '').trim();
        body = lines.filter(line => !line.startsWith('件名:')).join('\n').trim();
    } else {
        body = content;
    }
    
    // 件名設定
    if (subject) {
        item.subject.setAsync(subject, (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.warn('件名設定エラー:', result.error);
            }
        });
    }
    
    // 本文設定
    item.body.setAsync(
        body,
        { coercionType: Office.CoercionType.Text },
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showNotification('メール内容を挿入しました。');
            } else {
                showNotification('本文挿入エラー: ' + result.error.message);
            }
        }
    );
}

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
                        content: 'あなたはPTA活動を支援するAIアシスタントです。日本語で簡潔に回答してください。'
                    },
                    {
                        role: 'user',
                        content: prompt
                    }
                ],
                max_tokens: 500,
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
                        content: 'あなたはPTA活動を支援するAIアシスタントです。日本語で簡潔に回答してください。'
                    },
                    {
                        role: 'user',
                        content: prompt
                    }
                ],
                max_tokens: 500,
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
 * 履歴に保存
 */
function saveToHistory(historyItem) {
    Office.context.roamingSettings.get('ptaHistory', (historyData) => {
        let history = [];
        if (historyData) {
            history = JSON.parse(historyData);
        }
        
        history.unshift(historyItem);
        
        // 最新50件のみ保持
        if (history.length > 50) {
            history = history.slice(0, 50);
        }
        
        Office.context.roamingSettings.set('ptaHistory', JSON.stringify(history));
        Office.context.roamingSettings.saveAsync();
    });
}

/**
 * 通知を表示
 */
function showNotification(message) {
    Office.context.mailbox.item.notificationMessages.addAsync(
        'ptaNotification',
        {
            type: 'informationalMessage',
            message: message,
            icon: 'icon-16',
            persistent: false
        }
    );
}