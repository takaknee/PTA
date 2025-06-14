/*
 * PTA Edge拡張機能 - バックグラウンドサービスワーカー
 * Copyright (c) 2024 PTA Development Team
 */

// 拡張機能のインストール時
chrome.runtime.onInstalled.addListener((details) => {
    console.log('PTA支援ツールがインストールされました');
    
    // 初期設定を保存
    if (details.reason === 'install') {
        chrome.storage.local.set({
            'pta_settings': {
                provider: 'azure',
                model: 'gpt-4',
                apiKey: '',
                azureEndpoint: '',
                initialized: false
            }
        });
    }
    
    // コンテキストメニューを作成
    createContextMenus();
});

/**
 * コンテキストメニューを作成
 */
function createContextMenus() {
    // 既存のメニューをクリア
    chrome.contextMenus.removeAll(() => {
        // 選択テキスト用メニュー
        chrome.contextMenus.create({
            id: 'pta-analyze-selection',
            title: '🏫 選択文をPTA支援ツールで分析',
            contexts: ['selection']
        });
        
        // ページ全体用メニュー
        chrome.contextMenus.create({
            id: 'pta-analyze-page',
            title: '🏫 このページをPTA支援ツールで要約',
            contexts: ['page']
        });
    });
}

// コンテキストメニューのクリックハンドラー
chrome.contextMenus.onClicked.addListener((info, tab) => {
    switch (info.menuItemId) {
        case 'pta-analyze-selection':
            // 選択されたテキストを分析
            chrome.tabs.sendMessage(tab.id, {
                action: 'analyzeSelection',
                data: {
                    selectedText: info.selectionText,
                    pageUrl: info.pageUrl,
                    pageTitle: tab.title
                }
            });
            break;
            
        case 'pta-analyze-page':
            // ページ全体を要約
            chrome.tabs.sendMessage(tab.id, {
                action: 'analyzePage',
                data: {
                    pageUrl: info.pageUrl,
                    pageTitle: tab.title
                }
            });
            break;
    }
});

// メッセージハンドラー
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
        case 'analyzeEmail':
            handleEmailAnalysis(message.data, sendResponse);
            return true; // 非同期レスポンス
            
        case 'analyzePage':
            handlePageAnalysis(message.data, sendResponse);
            return true; // 非同期レスポンス
            
        case 'analyzeSelection':
            handleSelectionAnalysis(message.data, sendResponse);
            return true; // 非同期レスポンス
            
        case 'composeEmail':
            handleEmailComposition(message.data, sendResponse);
            return true; // 非同期レスポンス
            
        case 'testApiConnection':
            handleApiTest(message.data, sendResponse);
            return true; // 非同期レスポンス
            
        default:
            sendResponse({ error: 'サポートされていないアクションです' });
    }
});

/**
 * メール解析処理
 */
async function handleEmailAnalysis(data, sendResponse) {
    try {
        const settings = await getSettings();
        
        if (!settings.apiKey) {
            sendResponse({ error: 'APIキーが設定されていません' });
            return;
        }
        
        const prompt = createAnalysisPrompt(data);
        const result = await callAIAPI(prompt, settings);
        
        // 履歴に保存
        await saveToHistory({
            type: 'analysis',
            timestamp: new Date().toISOString(),
            emailSubject: data.subject,
            result: result
        });
        
        sendResponse({ success: true, result: result });
    } catch (error) {
        console.error('メール解析エラー:', error);
        sendResponse({ error: error.message });
    }
}

/**
 * ページ解析処理
 */
async function handlePageAnalysis(data, sendResponse) {
    try {
        const settings = await getSettings();
        
        if (!settings.apiKey) {
            sendResponse({ error: 'APIキーが設定されていません' });
            return;
        }
        
        const prompt = createPageAnalysisPrompt(data);
        const result = await callAIAPI(prompt, settings);
        
        // 履歴に保存
        await saveToHistory({
            type: 'page_analysis',
            timestamp: new Date().toISOString(),
            pageTitle: data.pageTitle,
            pageUrl: data.pageUrl,
            result: result
        });
        
        sendResponse({ success: true, result: result });
    } catch (error) {
        console.error('ページ解析エラー:', error);
        sendResponse({ error: error.message });
    }
}

/**
 * 選択テキスト解析処理
 */
async function handleSelectionAnalysis(data, sendResponse) {
    try {
        const settings = await getSettings();
        
        if (!settings.apiKey) {
            sendResponse({ error: 'APIキーが設定されていません' });
            return;
        }
        
        const prompt = createSelectionAnalysisPrompt(data);
        const result = await callAIAPI(prompt, settings);
        
        // 履歴に保存
        await saveToHistory({
            type: 'selection_analysis',
            timestamp: new Date().toISOString(),
            pageTitle: data.pageTitle,
            pageUrl: data.pageUrl,
            selectedText: data.selectedText.substring(0, 100) + '...',
            result: result
        });
        
        sendResponse({ success: true, result: result });
    } catch (error) {
        console.error('選択テキスト解析エラー:', error);
        sendResponse({ error: error.message });
    }
}
async function handleEmailComposition(data, sendResponse) {
    try {
        const settings = await getSettings();
        
        if (!settings.apiKey) {
            sendResponse({ error: 'APIキーが設定されていません' });
            return;
        }
        
        const prompt = createCompositionPrompt(data);
        const result = await callAIAPI(prompt, settings);
        
        // 履歴に保存
        await saveToHistory({
            type: 'composition',
            timestamp: new Date().toISOString(),
            requestType: data.type,
            result: result
        });
        
        sendResponse({ success: true, result: result });
    } catch (error) {
        console.error('メール作成エラー:', error);
        sendResponse({ error: error.message });
    }
}

/**
 * API接続テスト
 */
async function handleApiTest(data, sendResponse) {
    try {
        const testPrompt = 'こんにちは。API接続テストです。簡単に挨拶をしてください。';
        await callAIAPI(testPrompt, data);
        
        sendResponse({ 
            success: true, 
            result: 'API接続テストが成功しました。' 
        });
    } catch (error) {
        console.error('API接続テストエラー:', error);
        sendResponse({ error: error.message });
    }
}

/**
 * AI APIを呼び出し（既存のoutlook-addinコードを移植）
 */
async function callAIAPI(prompt, settings) {
    const apiKey = settings.apiKey;
    const provider = settings.provider || 'azure';
    const model = settings.model || 'gpt-4';
    
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
            
        case 'azure': {
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
        }
            
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
 * 設定取得
 */
async function getSettings() {
    return new Promise((resolve) => {
        chrome.storage.local.get(['pta_settings'], (result) => {
            resolve(result.pta_settings || {});
        });
    });
}

/**
 * 履歴保存
 */
async function saveToHistory(entry) {
    return new Promise((resolve) => {
        chrome.storage.local.get(['pta_history'], (result) => {
            let history = result.pta_history || [];
            
            // 新しいエントリを先頭に追加
            history.unshift({
                ...entry,
                id: Date.now()
            });
            
            // 最新50件のみ保持
            if (history.length > 50) {
                history = history.slice(0, 50);
            }
            
            chrome.storage.local.set({ 'pta_history': history }, resolve);
        });
    });
}

/**
 * ページ解析用プロンプト作成
 */
function createPageAnalysisPrompt(data) {
    return `
以下のWebページを要約・分析してください：

ページタイトル: ${data.pageTitle || '（タイトルなし）'}
URL: ${data.pageUrl || ''}
ページ内容: ${data.pageContent || '（内容を取得中...）'}

以下の形式で回答してください：
## 📄 ページ要約
- このページの主要な内容を3-5行で要約

## 🎯 重要なポイント
- 特に注目すべき情報やデータ（箇条書き）

## 🏫 PTA活動への関連性
- PTA活動や学校関連業務に役立つ情報があれば指摘
- 特に関連がない場合は「直接的な関連性は低い」と記載

## 💡 アクション提案
- このページの情報を活用するための具体的な提案（あれば）
`;
}

/**
 * 選択テキスト解析用プロンプト作成
 */
function createSelectionAnalysisPrompt(data) {
    return `
以下の選択されたテキストを分析してください：

ページタイトル: ${data.pageTitle || '（タイトルなし）'}
URL: ${data.pageUrl || ''}
選択されたテキスト:
${data.selectedText || '（テキストなし）'}

以下の形式で回答してください：
## 📝 選択テキストの要約
- 選択された内容の要点を2-3行で要約

## 🔍 詳細分析
- 重要な情報やキーワードの解説
- 背景情報や補足説明（必要に応じて）

## 🏫 PTA活動への活用
- この情報がPTA活動にどう役立つか
- 学校関連業務での活用方法
- 特に関連がない場合は「直接的な関連性は低い」と記載

## ⚡ 次のアクション
- この情報を受けて取るべき行動があれば提案
`;
}
function createAnalysisPrompt(emailData) {
    return `
以下のメールを分析してください：

件名: ${emailData.subject || '（件名なし）'}
送信者: ${emailData.sender || '（送信者不明）'}
本文:
${emailData.body || '（本文なし）'}

以下の形式で回答してください：
## 📧 メール要約
- 主要な内容（2-3行）

## ⚠️ 重要度
重要度: 高/中/低

## 📋 必要なアクション
- アクション項目（あれば）

## 💡 PTA観点でのコメント
- PTA活動に関連する重要な情報や注意点
`;
}

/**
 * メール作成用プロンプト作成
 */
function createCompositionPrompt(requestData) {
    const basePrompt = 'PTA活動に関するメールを作成してください。';
    
    switch (requestData.type) {
        case 'notice':
            return `${basePrompt}
内容: ${requestData.content}
種類: お知らせメール
要件: 丁寧で分かりやすい文面で、PTA会員向けのお知らせを作成してください。`;
            
        case 'reminder':
            return `${basePrompt}
内容: ${requestData.content}
種類: リマインダーメール
要件: 緊急度を適切に表現し、期限や重要な情報を強調してください。`;
            
        case 'survey':
            return `${basePrompt}
内容: ${requestData.content}
種類: アンケート依頼メール
要件: 協力をお願いする丁寧な文面で、回答方法を明確に示してください。`;
            
        default:
            return `${basePrompt}
内容: ${requestData.content}
要件: PTA活動に適した丁寧な文面で作成してください。`;
    }
}