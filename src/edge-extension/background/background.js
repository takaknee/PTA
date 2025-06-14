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
 * システム診断機能
 */
async function runSystemDiagnostics() {
    const diagnostics = {
        timestamp: new Date().toISOString(),
        chrome: {
            version: navigator.userAgent,
            offscreenSupport: typeof chrome.offscreen !== 'undefined',
            runtimeSupport: typeof chrome.runtime !== 'undefined'
        },
        permissions: {},
        offscreenDocument: {
            canCreate: false,
            exists: false,
            error: null
        },
        network: {
            basicConnectivity: false,
            openaiReachable: false,
            corsTest: false
        }
    };

    try {
        // 権限チェック
        const permissions = await chrome.permissions.getAll();
        diagnostics.permissions = permissions;

        // Offscreen document診断
        if (chrome.offscreen) {
            try {
                const contexts = await chrome.runtime.getContexts({
                    contextTypes: ['OFFSCREEN_DOCUMENT']
                });
                diagnostics.offscreenDocument.exists = contexts.length > 0;
                diagnostics.offscreenDocument.canCreate = true;
            } catch (error) {
                diagnostics.offscreenDocument.error = error.message;
            }
        }

        // 基本的なネットワーク接続テスト
        try {
            const response = await fetch('https://www.google.com', {
                method: 'HEAD',
                mode: 'no-cors'
            });
            diagnostics.network.basicConnectivity = true;
        } catch (error) {
            diagnostics.network.basicConnectivity = false;
        }

        // OpenAI API 到達性テスト
        try {
            const response = await fetch('https://api.openai.com', {
                method: 'HEAD',
                mode: 'no-cors'
            });
            diagnostics.network.openaiReachable = true;
        } catch (error) {
            diagnostics.network.openaiReachable = false;
        }

    } catch (error) {
        diagnostics.error = error.message;
    }

    console.log('システム診断結果:', diagnostics);
    return diagnostics;
}

/**
 * 改良版API接続テスト（診断機能付き）
 */
async function handleApiTest(data, sendResponse) {
    try {
        console.log('API接続テスト開始:', { provider: data.provider, model: data.model });

        // システム診断を実行
        const diagnostics = await runSystemDiagnostics();
        console.log('診断結果:', diagnostics);

        // 設定の基本検証
        if (!data.apiKey) {
            throw new Error('APIキーが設定されていません。');
        }

        if (data.provider === 'azure' && !data.azureEndpoint) {
            throw new Error('Azure エンドポイントが設定されていません。');
        }        // URLの妥当性チェック（詳細バリデーション）
        if (data.provider === 'azure') {
            const endpoint = data.azureEndpoint;

            if (!endpoint) {
                throw new Error('Azure エンドポイントが設定されていません。');
            }

            try {
                const url = new URL(endpoint);
                console.log('Azure エンドポイント解析:', {
                    protocol: url.protocol,
                    hostname: url.hostname,
                    pathname: url.pathname,
                    fullUrl: endpoint
                });

                // Azure OpenAI エンドポイントの形式チェック
                if (!url.hostname.includes('.openai.azure.com')) {
                    throw new Error(
                        `無効なAzure OpenAIエンドポイントです。\n\n` +
                        `入力値: ${endpoint}\n\n` +
                        `正しい形式: https://your-resource-name.openai.azure.com\n\n` +
                        `例: https://my-openai-resource.openai.azure.com`
                    );
                }

                if (url.protocol !== 'https:') {
                    throw new Error('Azure OpenAIエンドポイントはHTTPS形式である必要があります。');
                }

            } catch (urlError) {
                if (urlError.message.includes('無効なAzure OpenAI')) {
                    throw urlError; // 詳細なエラーメッセージをそのまま使用
                }

                throw new Error(
                    `Azure エンドポイントのURLが無効です。\n\n` +
                    `入力値: ${endpoint}\n\n` +
                    `エラー詳細: ${urlError.message}\n\n` +
                    `正しい形式: https://your-resource-name.openai.azure.com`
                );
            }
        }

        // Offscreen documentの状態確認
        if (!diagnostics.offscreenDocument.canCreate) {
            throw new Error('Offscreen document機能が利用できません。拡張機能の再インストールを試してください。');
        }

        const testPrompt = 'こんにちは。API接続テストです。「OK」とだけ返答してください。';

        console.log('API呼び出し開始...');
        const result = await callAIAPI(testPrompt, data);
        console.log('API接続テスト成功:', result);

        sendResponse({
            success: true,
            result: `API接続テストが成功しました。応答: ${result.substring(0, 50)}...`,
            diagnostics: diagnostics
        });
    } catch (error) {
        console.error('API接続テストエラー:', error);
        console.error('エラースタック:', error.stack);

        // 診断情報を取得
        const diagnostics = await runSystemDiagnostics();

        // より詳細なエラー情報を提供
        let errorMessage = error.message;

        if (error.message.includes('Failed to fetch') || error.message.includes('ネットワーク')) {
            errorMessage = `ネットワーク接続エラー: ${error.message}\n\n🔍 詳細診断:\n`;

            if (!diagnostics.network.basicConnectivity) {
                errorMessage += '• インターネット接続に問題があります\n';
            }

            if (!diagnostics.offscreenDocument.canCreate) {
                errorMessage += '• Offscreen document機能に問題があります\n';
            }

            if (diagnostics.offscreenDocument.error) {
                errorMessage += `• Offscreen document エラー: ${diagnostics.offscreenDocument.error}\n`;
            }

            errorMessage += '\n📋 推奨対策:\n';
            errorMessage += '• 拡張機能を無効化→有効化\n';
            errorMessage += '• ブラウザを再起動\n';
            errorMessage += '• 拡張機能を再インストール\n';
            errorMessage += '• ネットワーク設定を確認\n';
        }

        sendResponse({
            success: false,
            error: errorMessage,
            diagnostics: diagnostics
        });
    }
}

/**
 * AI APIを呼び出し（CORS制限回避版）
 */
async function callAIAPI(prompt, settings) {
    const apiKey = settings.apiKey;
    const provider = settings.provider || 'azure';
    const model = settings.model || 'gpt-4o-mini';

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
            if (!azureEndpoint) {
                throw new Error('Azure エンドポイントが設定されていません。');
            }
            endpoint = `${azureEndpoint}/openai/deployments/${model}/chat/completions?api-version=2024-02-15-preview`;
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

    // Offscreen documentを使用してCORS制限を回避
    return await fetchWithOffscreen({
        endpoint,
        headers,
        body,
        provider
    });
}

/**
 * Offscreen documentを使用したfetch
 */
async function fetchWithOffscreen(requestData) {
    console.log('fetchWithOffscreen開始:', requestData);

    try {
        // Offscreen documentを作成または取得
        await ensureOffscreenDocument();

        console.log('Offscreen documentに送信するデータ:', requestData);

        // Offscreen documentにメッセージを送信
        const response = await new Promise((resolve, reject) => {
            const timeout = setTimeout(() => {
                reject(new Error('Offscreen document通信タイムアウト（30秒）'));
            }, 30000);

            chrome.runtime.sendMessage({
                action: 'fetchAPI',
                data: requestData
            }, (response) => {
                clearTimeout(timeout);

                if (chrome.runtime.lastError) {
                    console.error('Runtime error:', chrome.runtime.lastError);
                    reject(new Error('Chrome runtime error: ' + chrome.runtime.lastError.message));
                    return;
                }

                if (!response) {
                    reject(new Error('Offscreen documentからの応答がありません'));
                    return;
                }

                resolve(response);
            });
        });

        console.log('Offscreen documentからの応答:', response);

        if (response.success) {
            return response.data;
        } else {
            throw new Error(response.error || 'Offscreen documentでエラーが発生しました');
        }
    } catch (error) {
        console.error('fetchWithOffscreen エラー:', error);

        // デバッグ用：フォールバック処理の詳細ログ
        console.log('フォールバック処理を開始します...');
        try {
            return await fallbackDirectFetch(requestData);
        } catch (fallbackError) {
            console.error('フォールバック処理も失敗:', fallbackError);
            throw new Error(`主処理とフォールバック処理の両方が失敗しました:\n主処理: ${error.message}\nフォールバック: ${fallbackError.message}`);
        }
    }
}

/**
 * Offscreen documentの作成・管理
 */
async function ensureOffscreenDocument() {
    const offscreenUrl = chrome.runtime.getURL('offscreen/offscreen.html');
    console.log('Offscreen document URL:', offscreenUrl);

    try {
        // 既存のoffscreen documentがあるかチェック
        const existingContexts = await chrome.runtime.getContexts({
            contextTypes: ['OFFSCREEN_DOCUMENT'],
            documentUrls: [offscreenUrl]
        });

        console.log('既存のOffscreen document:', existingContexts.length);

        if (existingContexts.length > 0) {
            console.log('Offscreen document既に存在します');
            return; // 既に存在する
        }

        console.log('新しいOffscreen documentを作成中...');

        // 新しいoffscreen documentを作成
        await chrome.offscreen.createDocument({
            url: offscreenUrl,
            reasons: [chrome.offscreen.Reason.WORKERS], // API呼び出し用に修正
            justification: 'CORS制限回避のためのAPI呼び出し処理'
        });

        console.log('Offscreen document作成完了');

        // 少し待機して初期化を完了させる
        await new Promise(resolve => setTimeout(resolve, 1000));

    } catch (error) {
        console.error('Offscreen document作成エラー:', error);
        console.error('Error details:', error.stack);
        throw new Error('Offscreen document作成に失敗しました: ' + error.message);
    }
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

/**
 * フォールバック: 直接fetchを試行（デバッグ・テスト用）
 */
async function fallbackDirectFetch(requestData) {
    console.log('フォールバック: 直接fetch実行');
    console.log('Request data:', requestData);

    const { endpoint, headers, body, provider } = requestData;

    try {
        // まずはCORSありで試行
        console.log('CORS有効でのfetch試行中...');
        let response = await fetch(endpoint, {
            method: 'POST',
            headers: headers,
            body: body,
            mode: 'cors',
            credentials: 'omit',
            cache: 'no-cache'
        });

        if (response.ok) {
            const data = await response.json();
            console.log('直接fetch成功（CORS有効）:', data);

            // OpenAI/Azure OpenAI の応答解析
            if (provider === 'openai' || provider === 'azure') {
                if (data.choices && data.choices.length > 0) {
                    return data.choices[0].message.content.trim();
                } else {
                    throw new Error('AIからの有効な応答が得られませんでした');
                }
            }

            return data;
        } else {
            const errorText = await response.text();
            throw new Error(`API呼び出しエラー (${response.status}): ${errorText}`);
        }

    } catch (corsError) {
        console.log('CORS有効でのfetch失敗:', corsError.message);

        // CORS無効で再試行
        try {
            console.log('CORS無効でのfetch試行中...');
            const response = await fetch(endpoint, {
                method: 'POST',
                headers: headers,
                body: body,
                mode: 'no-cors', // CORSを無効化（レスポンスは opaque になる）
            });

            // no-corsモードでは詳細なレスポンス情報を取得できない
            console.log('フォールバック fetch完了（no-corsモード）');
            throw new Error('no-corsモードでは応答内容を確認できません。適切なCORS設定またはOffscreen documentが必要です。');

        } catch (noCorsError) {
            console.error('フォールバック fetch エラー:', noCorsError);

            // 詳細なエラー情報を提供
            if (corsError.name === 'TypeError' && corsError.message.includes('Failed to fetch')) {
                throw new Error(
                    'ネットワーク接続エラーが発生しました:\n\n' +
                    '■ 考えられる原因:\n' +
                    '• インターネット接続の問題\n' +
                    '• APIエンドポイントURLの間違い\n' +
                    '• プロキシ設定またはファイアウォールによるブロック\n' +
                    '• VPNやセキュリティソフトの影響\n' +
                    '• CORS制限\n\n' +
                    '■ 対策:\n' +
                    '• ネットワーク接続を確認\n' +
                    '• 拡張機能を一度無効化→有効化\n' +
                    '• ブラウザを再起動\n' +
                    '• 設定画面でAPIキーとエンドポイントを再確認\n\n' +
                    `詳細エラー: ${corsError.message}`
                );
            }

            throw new Error(`API呼び出しに失敗しました: ${corsError.message}`);
        }
    }
}