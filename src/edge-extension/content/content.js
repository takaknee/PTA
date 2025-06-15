/*
 * AI Edge拡張機能 - コンテンツスクリプト（修正版）
 * Copyright (c) 2024 AI Development Team
 */

// 現在のメールサービスを判定
let currentService = 'unknown';
if (window.location.hostname.includes('outlook.office.com') || window.location.hostname.includes('outlook.live.com')) {
    currentService = 'outlook';
} else if (window.location.hostname.includes('mail.google.com')) {
    currentService = 'gmail';
} else {
    currentService = 'general'; // 一般的なWebページ
}

// AI支援ボタンを追加
let aiButton = null;

// バックグラウンドスクリプトからのメッセージを受信
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
        case 'analyzePage':
            handlePageAnalysis(message.data);
            break;
        case 'analyzeSelection':
            handleSelectionAnalysis(message.data);
            break;
    }
});

// ページ読み込み完了を待機
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize);
} else {
    initialize();
}

/**
 * 初期化処理
 */
function initialize() {
    console.log('AI支援ツール初期化開始:', currentService);

    // テーマ検出と適用（最初に実行）
    detectAndApplyTheme();

    // サービスに応じた初期化
    if (currentService === 'outlook') {
        initializeOutlook();
    } else if (currentService === 'gmail') {
        initializeGmail();
    } else {
        // 一般的なWebページの場合は常にボタンを表示
        initializeGeneralPage();
    }

    // URLの変更を監視（SPA対応）
    observeUrlChanges();
}

/**
 * 一般的なWebページの初期化
 */
function initializeGeneralPage() {
    // フローティングボタンを追加
    addAISupportButton();
}

/**
 * Outlookの初期化
 */
function initializeOutlook() {
    // Outlookのメール読み込み完了を待機
    const observer = new MutationObserver(() => {
        const emailContent = document.querySelector('[role="main"] [role="document"]');
        if (emailContent && !aiButton) {
            addAISupportButton();
        }
    });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });

    // 初回チェック
    setTimeout(() => {
        const emailContent = document.querySelector('[role="main"] [role="document"]');
        if (emailContent && !aiButton) {
            addAISupportButton();
        }
    }, 2000);
}

/**
 * Gmailの初期化
 */
function initializeGmail() {
    // Gmailのメール読み込み完了を待機
    const observer = new MutationObserver(() => {
        const emailContent = document.querySelector('.ii.gt .a3s.aiL');
        if (emailContent && !aiButton) {
            addAISupportButton();
        }
    });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });

    // 初回チェック
    setTimeout(() => {
        const emailContent = document.querySelector('.ii.gt .a3s.aiL');
        if (emailContent && !aiButton) {
            addAISupportButton();
        }
    }, 2000);
}

/**
 * AI支援ボタンを追加
 */
function addAISupportButton() {
    // 既存のボタンがあれば削除
    if (aiButton) {
        aiButton.remove();
    }

    // ボタンを作成
    aiButton = document.createElement('button');
    aiButton.className = 'ai-support-button';
    aiButton.innerHTML = `
        <div class="ai-button-content">
            <span class="ai-icon">🏫</span>
            <span class="ai-text">AI支援</span>
        </div>
    `;

    // 基本スタイルを設定
    aiButton.style.cssText = `
        position: fixed !important;
        top: 20px !important;
        right: 20px !important;
        z-index: 2147483647 !important;
        background: linear-gradient(135deg, #2196F3, #1976D2) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 10px 16px !important;
        cursor: pointer !important;
        font-family: 'Segoe UI', sans-serif !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        box-shadow: 0 4px 12px rgba(33, 150, 243, 0.3) !important;
        transition: all 0.3s ease !important;
    `;

    // ホバー効果
    aiButton.addEventListener('mouseenter', () => {
        aiButton.style.transform = 'translateY(-2px)';
        aiButton.style.boxShadow = '0 6px 16px rgba(33, 150, 243, 0.4)';
    });

    aiButton.addEventListener('mouseleave', () => {
        aiButton.style.transform = 'translateY(0)';
        aiButton.style.boxShadow = '0 4px 12px rgba(33, 150, 243, 0.3)';
    });

    // クリックイベント
    aiButton.addEventListener('click', showAiDialog);

    // ページに追加
    document.body.appendChild(aiButton);

    console.log('AI支援ボタンを追加しました');
}

/**
 * AIダイアログを表示
 */
function showAiDialog() {
    const selectedText = getSelectedText();
    const pageContent = extractPageContent();

    const dialogData = {
        pageTitle: document.title,
        pageUrl: window.location.href,
        pageContent: pageContent,
        selectedText: selectedText,
        currentService: currentService
    };

    // メール情報が利用可能な場合は追加
    if (currentService === 'outlook' || currentService === 'gmail') {
        const emailData = getCurrentEmailData();
        Object.assign(dialogData, emailData);
    }

    createAiDialog(dialogData);
}

/**
 * AIダイアログを作成（確実にモーダル表示）
 */
function createAiDialog(dialogData) {
    // 既存のダイアログを削除
    const existingDialog = document.getElementById('ai-dialog');
    if (existingDialog) {
        existingDialog.remove();
    }

    // ダイアログコンテナを作成
    const dialog = document.createElement('div');
    dialog.id = 'ai-dialog';
    dialog.dialogData = dialogData; // データを保存

    // 強制的にモーダルスタイル適用
    dialog.style.cssText = `
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        z-index: 2147483647 !important;
        background: rgba(0, 0, 0, 0.7) !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-family: 'Segoe UI', sans-serif !important;
    `;    // テーマを検出
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const prefersHighContrast = window.matchMedia('(prefers-contrast: high)').matches;

    console.log('🎨 ダイアログテーマ検出:', { prefersDark, prefersHighContrast });

    // テーマに応じた色設定
    let dialogBg, dialogText, headerBg, borderColor;

    if (prefersHighContrast && prefersDark) {
        dialogBg = '#000000';
        dialogText = '#ffffff';
        headerBg = '#000000';
        borderColor = '#ffffff';
    } else if (prefersHighContrast) {
        dialogBg = '#ffffff';
        dialogText = '#000000';
        headerBg = '#ffffff';
        borderColor = '#000000';
    } else if (prefersDark) {
        dialogBg = '#2d2d2d';
        dialogText = '#ffffff';
        headerBg = 'linear-gradient(135deg, #3a3a3a, #2d2d2d)';
        borderColor = '#555555';
    } else {
        dialogBg = '#ffffff';
        dialogText = '#333333';
        headerBg = 'linear-gradient(135deg, #2196F3, #1976D2)';
        borderColor = '#e0e0e0';
    }

    // テーマに応じた追加色設定
    let textMuted, infoBg, headerTextColor;

    if (prefersDark) {
        textMuted = '#cccccc';
        infoBg = '#404040';
        headerTextColor = '#ffffff';
    } else {
        textMuted = '#666666';
        infoBg = '#f0f8ff';
        headerTextColor = '#ffffff';
    }

    // ダイアログコンテンツを作成
    const content = document.createElement('div');
    content.style.cssText = `
        background: ${dialogBg} !important;
        color: ${dialogText} !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3) !important;
        max-width: 600px !important;
        width: 90% !important;
        max-height: 80vh !important;
        overflow-y: auto !important;
        position: relative !important;
    `;
    content.innerHTML = `
        <div class="ai-dialog-header" style="
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 20px;            border-bottom: 1px solid ${borderColor};
            background: ${headerBg};
            color: white;
            border-radius: 12px 12px 0 0;
        ">            <h3 style="margin: 0; font-size: 18px; font-weight: 600; color: ${headerTextColor};">🏫 AI支援ツール</h3>
            <button class="ai-close-btn" style="
                background: none;
                border: none;
                color: ${headerTextColor};
                font-size: 24px;
                cursor: pointer;
                width: 32px;
                height: 32px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 50%;
            ">×</button>
        </div>
        <div style="padding: 20px;">
            <div style="margin-bottom: 20px;">                <h4 style="margin: 0 0 8px 0; color: ${dialogText};">📄 ${dialogData.pageTitle || 'ページ情報'}</h4>
                <p style="margin: 0 0 16px 0; color: ${textMuted}; font-size: 12px; word-break: break-all;">${dialogData.pageUrl || ''}</p>
                ${dialogData.selectedText ? `<div style="background: ${infoBg}; padding: 12px; border-radius: 6px; border-left: 4px solid #2196F3; margin-bottom: 16px; color: ${dialogText};"><strong>選択テキスト:</strong> ${dialogData.selectedText.substring(0, 100)}...</div>` : ''}
            </div>
            
            <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px;">
                ${currentService === 'outlook' || currentService === 'gmail' ? `
                    <button class="ai-analyze-email-btn" style="
                        background: linear-gradient(135deg, #2196F3, #1976D2);
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 12px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                    ">📧 メール解析</button>
                    <button class="ai-compose-reply-btn" style="
                        background: linear-gradient(135deg, #4CAF50, #388E3C);
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 12px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                    ">📝 返信作成</button>
                ` : `
                    <button class="ai-analyze-page-btn" style="
                        background: linear-gradient(135deg, #2196F3, #1976D2);
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 12px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                    ">📄 ページ要約</button>
                `}
                ${dialogData.selectedText ? `
                    <button class="ai-analyze-selection-btn" style="
                        background: linear-gradient(135deg, #FF9800, #F57C00);
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 12px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                    ">🔍 選択テキスト分析</button>
                ` : ''}                <button class="ai-open-settings-btn" style="
                    background: ${prefersDark ? '#555555' : '#f5f5f5'};
                    color: ${prefersDark ? '#cccccc' : '#666666'};
                    border: 1px solid ${prefersDark ? '#777777' : '#dddddd'};
                    border-radius: 6px;
                    padding: 12px 16px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 500;
                ">⚙️ 設定</button>
            </div>
            
            <div>                <div id="ai-loading" style="display: none; text-align: center; padding: 20px; color: ${textMuted};">
                    <div style="
                        border: 4px solid ${prefersDark ? '#555555' : '#f3f3f3'};
                        border-top: 4px solid #2196F3;
                        border-radius: 50%;
                        width: 40px;
                        height: 40px;
                        animation: spin 1s linear infinite;
                        margin: 0 auto 16px;
                    "></div>
                    <span>AI処理中...</span>
                </div>
                <div id="ai-result" style="display: none; background: ${prefersDark ? '#404040' : '#f9f9f9'}; padding: 16px; border-radius: 8px; border-left: 4px solid #2196F3; color: ${dialogText};"></div>
            </div>
        </div>
    `;

    dialog.appendChild(content);

    // スピナーアニメーション用のCSSを追加
    if (!document.getElementById('ai-spinner-style')) {
        const style = document.createElement('style');
        style.id = 'ai-spinner-style';
        style.textContent = `
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        `;
        document.head.appendChild(style);
    }
    // bodyに直接追加してiframeを回避
    document.body.appendChild(dialog);

    // デバッグ: 挿入されたHTMLの内容を確認
    console.log('ダイアログHTML:', content.innerHTML);

    // CSP準拠のイベントリスナー設定
    setupDialogEventListeners(dialog);

    // オーバーレイクリックで閉じる
    dialog.addEventListener('click', (e) => {
        if (e.target === dialog) {
            closeAiDialog();
        }
    });

    // ESCキーで閉じる
    document.addEventListener('keydown', handleEscapeKey);

    // ダイアログにテーマクラスを適用
    applyThemeToDialog(dialog);

    console.log('AIダイアログを作成しました（モーダル表示）');
}

/**
 * ダイアログにテーマクラスを適用
 */
function applyThemeToDialog(dialogElement) {
    // 現在のテーマを検出
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const prefersHighContrast = window.matchMedia('(prefers-contrast: high)').matches;

    // テーマクラスを追加
    dialogElement.classList.remove('ai-theme-light', 'ai-theme-dark', 'ai-theme-high-contrast');

    if (prefersHighContrast) {
        dialogElement.classList.add('ai-theme-high-contrast');
    } else if (prefersDark) {
        dialogElement.classList.add('ai-theme-dark');
    } else {
        dialogElement.classList.add('ai-theme-light');
    }

    console.log('ダイアログテーマ適用:', { prefersDark, prefersHighContrast });
}

/**
 * ダイアログのイベントリスナーを設定（CSP準拠）
 */
function setupDialogEventListeners(dialog) {
    console.log('setupDialogEventListeners 開始');

    // 閉じるボタン
    const closeBtn = dialog.querySelector('.ai-close-btn');
    if (closeBtn) {
        closeBtn.addEventListener('click', closeAiDialog);
        console.log('閉じるボタンのイベントリスナー設定完了');
    }

    // メール解析ボタン
    const analyzeEmailBtn = dialog.querySelector('.ai-analyze-email-btn');
    if (analyzeEmailBtn) {
        analyzeEmailBtn.addEventListener('click', analyzeEmail);
        console.log('メール解析ボタンのイベントリスナー設定完了');
    }

    // 返信作成ボタン
    const composeReplyBtn = dialog.querySelector('.ai-compose-reply-btn');
    if (composeReplyBtn) {
        composeReplyBtn.addEventListener('click', composeReply);
        console.log('返信作成ボタンのイベントリスナー設定完了');
    }

    // ページ解析ボタン
    const analyzePageBtn = dialog.querySelector('.ai-analyze-page-btn');
    if (analyzePageBtn) {
        analyzePageBtn.addEventListener('click', analyzePage);
        console.log('ページ解析ボタンのイベントリスナー設定完了');
    } else {
        console.log('ページ解析ボタンが見つかりません');
    }

    // 選択テキスト解析ボタン
    const analyzeSelectionBtn = dialog.querySelector('.ai-analyze-selection-btn');
    if (analyzeSelectionBtn) {
        analyzeSelectionBtn.addEventListener('click', analyzeSelection);
        console.log('選択テキスト解析ボタンのイベントリスナー設定完了');
    }

    // 設定ボタン
    const openSettingsBtn = dialog.querySelector('.ai-open-settings-btn');
    if (openSettingsBtn) {
        openSettingsBtn.addEventListener('click', openSettings);
        console.log('設定ボタンのイベントリスナー設定完了');
    }

    console.log('setupDialogEventListeners 完了');
}

/**
 * ダイアログを閉じる
 */
function closeAiDialog() {
    const dialog = document.getElementById('ai-dialog');
    if (dialog) {
        dialog.remove();
    }

    // ESCキーリスナーを削除
    document.removeEventListener('keydown', handleEscapeKey);
}

/**
 * ESCキー処理
 */
function handleEscapeKey(event) {
    if (event.key === 'Escape') {
        closeAiDialog();
    }
}

/**
 * メールの現在のデータを取得
 */
function getCurrentEmailData() {
    const data = {
        service: currentService,
        subject: '',
        body: '',
        sender: '',
        recipients: []
    };

    if (currentService === 'outlook') {
        // Outlookからデータを取得
        const subjectElement = document.querySelector('[role="main"] h1, [aria-label*="件名"]');
        if (subjectElement) {
            data.subject = subjectElement.textContent || subjectElement.innerText;
        }

        const bodyElement = document.querySelector('[role="main"] [role="document"], .allowTextSelection');
        if (bodyElement) {
            data.body = bodyElement.textContent || bodyElement.innerText;
        }

        const senderElement = document.querySelector('[aria-label*="送信者"], .PersonaCard');
        if (senderElement) {
            data.sender = senderElement.textContent || senderElement.innerText;
        }
    } else if (currentService === 'gmail') {
        // Gmailからデータを取得
        const subjectElement = document.querySelector('h2[data-legacy-thread-id]');
        if (subjectElement) {
            data.subject = subjectElement.textContent || subjectElement.innerText;
        }

        const bodyElement = document.querySelector('.ii.gt .a3s.aiL');
        if (bodyElement) {
            data.body = bodyElement.textContent || bodyElement.innerText;
        }

        const senderElement = document.querySelector('.go .gD');
        if (senderElement) {
            data.sender = senderElement.getAttribute('email') || senderElement.textContent;
        }
    }

    return data;
}

/**
 * ページ上の選択されたテキストを取得
 */
function getSelectedText() {
    const selection = window.getSelection();
    return selection.toString().trim();
}

/**
 * 通知を表示
 */
function showNotification(message, type = 'info') {
    // 既存の通知を削除
    const existingNotification = document.getElementById('ai-notification');
    if (existingNotification) {
        existingNotification.remove();
    }

    const notification = document.createElement('div');
    notification.id = 'ai-notification';
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 8px;
        color: white;
        font-family: 'Segoe UI', sans-serif;
        font-size: 14px;
        z-index: 2147483647;
        max-width: 300px;
        word-wrap: break-word;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
        transition: all 0.3s ease;
        ${type === 'success' ? 'background: #4CAF50;' :
            type === 'error' ? 'background: #f44336;' :
                'background: #2196F3;'}
    `;
    notification.textContent = message;
    document.body.appendChild(notification);

    // 3秒後に自動削除
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
    }, 3000);
}

/**
 * ローディング表示
 */
function showLoading() {
    const loadingElement = document.getElementById('ai-loading');
    const resultElement = document.getElementById('ai-result');

    if (loadingElement) {
        loadingElement.style.display = 'block';
    }
    if (resultElement) {
        resultElement.style.display = 'none';
    }
}

/**
 * ローディング非表示
 */
function hideLoading() {
    const loadingElement = document.getElementById('ai-loading');

    if (loadingElement) {
        loadingElement.style.display = 'none';
    }
}

/**
 * 結果を表示
 */
function showResult(result) {
    const resultElement = document.getElementById('ai-result');

    if (resultElement) {
        resultElement.innerHTML = result;
        resultElement.style.display = 'block';
    }
}

/**
 * メール解析を実行
 */
function analyzeEmail() {
    const dialog = document.getElementById('ai-dialog');
    const emailData = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzeEmail',
        data: emailData
    }, (response) => {
        hideLoading();

        if (response && response.success) {
            showResult(response.result);
            showNotification('メール解析が完了しました', 'success');
        } else {
            const errorMessage = response ? response.error : '不明なエラーが発生しました';
            showResult(`<div style="color: #f44336;">❌ エラー: ${errorMessage}</div>`);
            showNotification('メール解析に失敗しました', 'error');
        }
    });
}

/**
 * ページ解析を実行
 */
function analyzePage() {
    console.log('analyzePage 関数が呼び出されました');

    const dialog = document.getElementById('ai-dialog');
    if (!dialog) {
        console.error('ダイアログが見つかりません');
        return;
    }

    const pageData = dialog.dialogData;
    console.log('ページデータ:', pageData);

    // pageContentの値を詳細にチェック
    if (!pageData.pageContent) {
        console.error('⚠️ pageContent が未定義です!');
    } else if (pageData.pageContent === 'undefined') {
        console.error('⚠️ pageContent が文字列の "undefined" です!');
    } else if (pageData.pageContent.trim() === '') {
        console.error('⚠️ pageContent が空文字列です!');
    } else {
        console.log('✅ pageContent が正常に設定されています:', pageData.pageContent.substring(0, 200) + '...');
    }

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzePage',
        data: pageData
    }, (response) => {
        hideLoading();

        if (response && response.success) {
            showResult(response.result);
            showNotification('ページ解析が完了しました', 'success');
        } else {
            const errorMessage = response ? response.error : '不明なエラーが発生しました';
            showResult(`<div style="color: #f44336;">❌ エラー: ${errorMessage}</div>`);
            showNotification('ページ解析に失敗しました', 'error');
        }
    });
}

/**
 * 選択テキスト解析を実行
 */
function analyzeSelection() {
    const dialog = document.getElementById('ai-dialog');
    const selectionData = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzeSelection',
        data: selectionData
    }, (response) => {
        hideLoading();

        if (response && response.success) {
            showResult(response.result);
            showNotification('選択テキスト解析が完了しました', 'success');
        } else {
            const errorMessage = response ? response.error : '不明なエラーが発生しました';
            showResult(`<div style="color: #f44336;">❌ エラー: ${errorMessage}</div>`);
            showNotification('選択テキスト解析に失敗しました', 'error');
        }
    });
}

/**
 * 返信作成を実行
 */
function composeReply() {
    const dialog = document.getElementById('ai-dialog');
    const emailData = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'composeEmail',
        data: emailData
    }, (response) => {
        hideLoading();

        if (response && response.success) {
            showResult(response.result);
            showNotification('返信作成が完了しました', 'success');
        } else {
            const errorMessage = response ? response.error : '不明なエラーが発生しました';
            showResult(`<div style="color: #f44336;">❌ エラー: ${errorMessage}</div>`);
            showNotification('返信作成に失敗しました', 'error');
        }
    });
}

/**
 * 設定画面を開く
 */
function openSettings() {
    chrome.runtime.openOptionsPage();
}

/**
 * ページ解析ハンドラー
 */
function handlePageAnalysis(data) {
    const pageContent = extractPageContent();

    const pageData = {
        pageTitle: document.title,
        pageUrl: window.location.href,
        pageContent: pageContent
    };

    createAiDialog(pageData);
}

/**
 * 選択テキスト解析ハンドラー
 */
function handleSelectionAnalysis(data) {
    const selectedText = getSelectedText();
    if (!selectedText) {
        showNotification('テキストが選択されていません', 'error');
        return;
    }

    const selectionData = {
        pageTitle: document.title,
        pageUrl: window.location.href,
        selectedText: selectedText
    };

    createAiDialog(selectionData);
}

/**
 * URL変更を監視（SPA対応）
 */
function observeUrlChanges() {
    let currentUrl = window.location.href;

    const observer = new MutationObserver(() => {
        if (window.location.href !== currentUrl) {
            currentUrl = window.location.href;

            // URL変更時に既存のボタンを削除
            if (aiButton) {
                aiButton.remove();
                aiButton = null;
            }

            // 新しいページでボタンを再追加
            setTimeout(() => {
                if (currentService === 'outlook') {
                    initializeOutlook();
                } else if (currentService === 'gmail') {
                    initializeGmail();
                } else {
                    initializeGeneralPage();
                }
            }, 1000);
        }
    });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });
}

/**
 * テーマ検出とクラス設定
 */
function detectAndApplyTheme() {
    // システムのテーマ設定を検出
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const prefersHighContrast = window.matchMedia('(prefers-contrast: high)').matches;

    console.log('テーマ検出:', { prefersDark, prefersHighContrast });

    // body要素にテーマクラスを追加
    document.body.classList.remove('ai-theme-light', 'ai-theme-dark', 'ai-theme-high-contrast');

    if (prefersHighContrast) {
        document.body.classList.add('ai-theme-high-contrast');
    } else if (prefersDark) {
        document.body.classList.add('ai-theme-dark');
    } else {
        document.body.classList.add('ai-theme-light');
    }

    // テーマ変更を監視
    const darkModeQuery = window.matchMedia('(prefers-color-scheme: dark)');
    const highContrastQuery = window.matchMedia('(prefers-contrast: high)');

    darkModeQuery.addEventListener('change', detectAndApplyTheme);
    highContrastQuery.addEventListener('change', detectAndApplyTheme);
}

// グローバル関数として公開（ダイアログから呼び出し可能にするため）
window.analyzeEmail = analyzeEmail;
window.analyzePage = analyzePage;
window.analyzeSelection = analyzeSelection;
window.composeReply = composeReply;
window.openSettings = openSettings;
window.closeAiDialog = closeAiDialog;

console.log('AI支援ツール - コンテンツスクリプト読み込み完了');
console.log('公開されたグローバル関数:', {
    analyzeEmail: typeof window.analyzeEmail,
    analyzePage: typeof window.analyzePage,
    analyzeSelection: typeof window.analyzeSelection,
    composeReply: typeof window.composeReply,
    openSettings: typeof window.openSettings,
    closeAiDialog: typeof window.closeAiDialog
});

/**
 * ページコンテンツを安全に抽出する共通関数
 * 複数の方法でフォールバックして確実にコンテンツを取得
 */
function extractPageContent() {
    let pageContent = '';

    try {
        // 方法1: メインコンテンツエリアを探す
        const mainSelectors = [
            'main',
            'article',
            '.content',
            '#content',
            '.main-content',
            '.post-content',
            '.entry-content',
            '[data-testid="article-body"]', // Qiita等の記事サイト
            '.markdown-body'  // GitHub等
        ];

        for (const selector of mainSelectors) {
            const element = document.querySelector(selector);
            if (element && element.innerText && element.innerText.trim()) {
                pageContent = element.innerText.trim();
                console.log(`コンテンツ抽出成功: ${selector}`);
                break;
            }
        }

        // 方法2: body全体から抽出（スクリプトやスタイルを除外）
        if (!pageContent.trim()) {
            const bodyClone = document.body.cloneNode(true);

            // 不要な要素を削除
            const unwantedSelectors = [
                'script', 'style', 'noscript',
                'nav', 'header', 'footer',
                '.menu', '.sidebar', '.advertisement', '.ad',
                '.social-share', '.comments', '.related-posts',
                '[class*="menu"]', '[class*="nav"]', '[class*="sidebar"]'
            ];

            unwantedSelectors.forEach(selector => {
                try {
                    const elements = bodyClone.querySelectorAll(selector);
                    elements.forEach(el => el.remove());
                } catch (e) {
                    // セレクタエラーを無視
                }
            });

            pageContent = bodyClone.innerText || bodyClone.textContent || '';
            console.log('body全体からコンテンツ抽出');
        }

        // 方法3: 空の場合のフォールバック
        if (!pageContent.trim()) {
            pageContent = document.body.innerText || document.body.textContent || '';
            console.log('フォールバックでコンテンツ抽出');
        }

        // 長すぎる場合は切り詰め（最初の8000文字）
        if (pageContent.length > 8000) {
            pageContent = pageContent.substring(0, 8000).trim();
            console.log('コンテンツを8000文字に切り詰め');
        }

        // 最終チェック
        if (!pageContent.trim()) {
            pageContent = '（ページコンテンツを取得できませんでした）';
            console.warn('ページコンテンツが空です');
        }

    } catch (error) {
        console.error('ページコンテンツ抽出エラー:', error);
        // エラーの場合は最低限の情報を返す
        try {
            pageContent = document.body.innerText || document.body.textContent || 'コンテンツを取得できませんでした';
        } catch (e) {
            pageContent = 'コンテンツ取得中にエラーが発生しました';
        }
    }

    console.log(`最終的に抽出されたコンテンツ（最初の200文字）:`, pageContent.substring(0, 200));
    return pageContent;
}
