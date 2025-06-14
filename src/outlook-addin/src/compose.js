/*
 * PTA Outlook アドイン - メール作成支援機能
 * Copyright (c) 2024 PTA Development Team
 */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.addEventListener("DOMContentLoaded", function() {
            // ボタンイベント設定
            document.getElementById("generate-button").onclick = generateEmailContent;
            document.getElementById("insert-button").onclick = insertGeneratedContent;
            
            // カスタムタイプ選択時の処理
            document.getElementById("email-type").onchange = handleEmailTypeChange;
        });
    }
});

let generatedContent = '';

/**
 * メール内容を生成
 */
async function generateEmailContent() {
    try {
        showLoading(true);
        
        // フォームデータ取得
        const formData = getFormData();
        
        // 入力値検証
        if (!formData.content.trim()) {
            showError('メッセージ内容を入力してください。');
            return;
        }
        
        // AI文章生成実行
        generatedContent = await generateAIContent(formData);
        
        // 結果表示
        displayGeneratedContent(generatedContent);
        
    } catch (error) {
        console.error('文章生成エラー:', error);
        showError('文章生成中にエラーが発生しました: ' + error.message);
    } finally {
        showLoading(false);
    }
}

/**
 * フォームデータを取得
 */
function getFormData() {
    return {
        type: document.getElementById("email-type").value,
        content: document.getElementById("message-content").value,
        tone: document.getElementById("tone").value
    };
}

/**
 * AI文章生成を実行
 */
async function generateAIContent(formData) {
    // 設定読み込み
    const settings = await loadSettings();
    
    // メールタイプに応じたプロンプト生成
    const basePrompt = createPromptFromType(formData.type, formData.content, formData.tone);
    
    try {
        const response = await callAIAPI(basePrompt, settings);
        return response;
    } catch (error) {
        throw new Error('AI APIエラー: ' + error.message);
    }
}

/**
 * メールタイプに応じたプロンプトを生成
 */
function createPromptFromType(type, content, tone) {
    let typeInstruction = '';
    
    switch (type) {
        case 'announcement':
            typeInstruction = 'PTA会員向けのお知らせメールとして、重要な情報を分かりやすく伝える';
            break;
        case 'request':
            typeInstruction = 'PTA活動への協力依頼として、お願いの理由と具体的な内容を丁寧に説明する';
            break;
        case 'reminder':
            typeInstruction = 'イベントや締切のリマインダーとして、必要な情報を整理して伝える';
            break;
        case 'response':
            typeInstruction = '問い合わせや要望への返信として、適切で建設的な回答を提供する';
            break;
        case 'custom':
            typeInstruction = 'PTA関連のメールとして適切な内容で';
            break;
        default:
            typeInstruction = 'PTA関連のメールとして';
    }
    
    let toneInstruction = '';
    switch (tone) {
        case 'formal':
            toneInstruction = '丁寧で正式な文体を使用し、敬語を適切に使って';
            break;
        case 'friendly':
            toneInstruction = '親しみやすく温かい文体を使用し、親近感を持てるように';
            break;
        case 'business':
            toneInstruction = 'ビジネスライクで簡潔な文体を使用し、要点を明確に';
            break;
        default:
            toneInstruction = '適切な文体で';
    }
    
    return `
あなたはPTA活動を支援するAIアシスタントです。以下の指示に従って、日本語でメール文章を作成してください。

目的: ${typeInstruction}
文体: ${toneInstruction}
内容: ${content}

以下の要件を満たしてください：
1. 件名も含めて作成する
2. 挨拶と締めの言葉を含める
3. 読みやすい段落構成にする
4. PTA活動に適した丁寧な表現を使用する
5. 必要に応じて具体的な行動を促す文言を含める

形式:
件名: [件名をここに]

[メール本文をここに]

よろしくお願いいたします。
`;
}

/**
 * AI APIを呼び出し（read.jsと共通関数）
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
 * 設定を読み込み（read.jsと共通関数）
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
 * 生成された内容を表示
 */
function displayGeneratedContent(content) {
    const resultContainer = document.getElementById('result-container');
    const resultDiv = document.getElementById('generated-text');
    
    resultDiv.innerHTML = `
        <div class="generated-content">
            <pre style="white-space: pre-wrap; font-family: inherit;">${content}</pre>
        </div>
    `;
    
    resultContainer.style.display = 'block';
    document.getElementById('insert-button').style.display = 'inline-block';
}

/**
 * 生成された内容をメール本文に挿入
 */
async function insertGeneratedContent() {
    if (!generatedContent) {
        showError('挿入する内容がありません。');
        return;
    }
    
    try {
        // 件名と本文を分離
        const lines = generatedContent.split('\n');
        let subject = '';
        let body = '';
        
        // 件名を抽出
        const subjectLine = lines.find(line => line.startsWith('件名:'));
        if (subjectLine) {
            subject = subjectLine.replace('件名:', '').trim();
            // 件名行を除去して本文を作成
            body = lines.filter(line => !line.startsWith('件名:')).join('\n').trim();
        } else {
            body = generatedContent;
        }
        
        // 件名設定
        if (subject) {
            Office.context.mailbox.item.subject.setAsync(subject, (result) => {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    console.warn('件名設定エラー:', result.error);
                }
            });
        }
        
        // 本文設定
        Office.context.mailbox.item.body.setAsync(
            body,
            { coercionType: Office.CoercionType.Text },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showSuccess('メール内容を挿入しました。');
                    
                    // 履歴に保存
                    saveGenerationToHistory();
                } else {
                    showError('本文挿入エラー: ' + result.error.message);
                }
            }
        );
        
    } catch (error) {
        console.error('挿入エラー:', error);
        showError('内容の挿入中にエラーが発生しました: ' + error.message);
    }
}

/**
 * 生成履歴を保存
 */
function saveGenerationToHistory() {
    const formData = getFormData();
    
    Office.context.roamingSettings.get('ptaHistory', (historyData) => {
        let history = [];
        if (historyData) {
            history = JSON.parse(historyData);
        }
        
        history.unshift({
            id: Date.now(),
            type: 'generation',
            timestamp: new Date().toLocaleString('ja-JP'),
            emailType: formData.type,
            tone: formData.tone,
            input: formData.content,
            result: generatedContent
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
 * メールタイプ変更時の処理
 */
function handleEmailTypeChange() {
    const type = document.getElementById("email-type").value;
    const messageContent = document.getElementById("message-content");
    
    // プレースホルダーをタイプに応じて変更
    let placeholder = '';
    switch (type) {
        case 'announcement':
            placeholder = '例: 次回PTA総会の開催について、日時と議題をお知らせします';
            break;
        case 'request':
            placeholder = '例: 運動会のボランティア募集について、お手伝いをお願いしたいと思います';
            break;
        case 'reminder':
            placeholder = '例: 来週の参観日について、忘れ物や注意事項をお知らせします';
            break;
        case 'response':
            placeholder = '例: ご質問いただいた件について、回答いたします';
            break;
        case 'custom':
            placeholder = '伝えたい内容を具体的に入力してください';
            break;
    }
    
    messageContent.placeholder = placeholder;
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

/**
 * 成功メッセージ表示
 */
function showSuccess(message) {
    const notificationContainer = document.querySelector('.notification-container');
    notificationContainer.innerHTML = `
        <div class="success-message">
            <p class="ms-font-m" style="color: #107c10;">成功: ${message}</p>
        </div>
    `;
}