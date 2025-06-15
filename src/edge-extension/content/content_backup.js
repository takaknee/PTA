/*
 * Shima Edge拡張機能 - コンテンツスクリプト
 * Copyright (c) 2024 Shima Development Team
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

// PTA支援ボタンを追加
let ptaButton = null;

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
    console.log('PTA支援ツール初期化開始:', currentService);

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
        if (emailContent && !ptaButton) {
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
        if (emailContent && !ptaButton) {
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
        if (emailContent && !ptaButton) {
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
        if (emailContent && !ptaButton) {
            addAISupportButton();
        }
    }, 2000);
}

/**
 * PTA支援ボタンを追加
 */
function addAISupportButton() {
    // 既存のボタンを削除
    const existingButton = document.getElementById('pta-support-button');
    if (existingButton) {
        existingButton.remove();
    }

    // ボタンを作成
    ptaButton = document.createElement('div');
    ptaButton.id = 'pta-support-button';
    ptaButton.className = 'pta-support-button';

    // サービスに応じてボタンテキストを変更
    let buttonText = 'PTA支援';
    if (currentService === 'outlook' || currentService === 'gmail') {
        buttonText = 'メール支援';
    } else {
        buttonText = 'ページ分析';
    }

    ptaButton.innerHTML = `
        <div class="pta-button-content">
            <span class="pta-icon">🏫</span>
            <span class="pta-text">${buttonText}</span>
        </div>
    `;

    // ボタンクリックイベント
    ptaButton.addEventListener('click', openPTADialog);

    // ボタンを適切な位置に配置
    insertPTAButton();
}

/**
 * サービスに応じてボタンを配置
 */
function insertPTAButton() {
    let targetElement = null;

    if (currentService === 'outlook') {
        // Outlookの場合はツールバーに追加
        targetElement = document.querySelector('[role="main"] [role="toolbar"]');
        if (!targetElement) {
            targetElement = document.querySelector('[role="main"] .ms-CommandBar');
        }
    } else if (currentService === 'gmail') {
        // Gmailの場合はメールヘッダーに追加
        targetElement = document.querySelector('.gE.iv.gt .gH .go .gK');
        if (!targetElement) {
            targetElement = document.querySelector('.gE.iv.gt .gH .go');
        }
    }

    if (targetElement) {
        // フローティングボタンとして配置
        ptaButton.style.position = 'fixed';
        ptaButton.style.top = '10px';
        ptaButton.style.right = '10px';
        ptaButton.style.zIndex = '10000';
        document.body.appendChild(ptaButton);
    } else {
        // フォールバック: 画面右上に配置
        ptaButton.style.position = 'fixed';
        ptaButton.style.top = '10px';
        ptaButton.style.right = '10px';
        ptaButton.style.zIndex = '10000';
        document.body.appendChild(ptaButton);
    }
}

/**
 * PTAダイアログを開く
 */
function openPTADialog() {
    // ページの種類に応じてデータを取得
    let dialogData = {};

    if (currentService === 'outlook' || currentService === 'gmail') {
        // メールサービスの場合は既存のメール取得ロジック
        dialogData = getCurrentEmailData();
        if (!dialogData.body) {
            showNotification('メール内容を取得できませんでした。', 'error');
            return;
        }
    } else {
        // 一般的なWebページの場合はページ情報を取得
        dialogData = getCurrentPageData();
    }

    // ダイアログを作成
    createPTADialog(dialogData);
}

/**
 * 現在のページデータを取得
 */
function getCurrentPageData() {
    return {
        pageTitle: document.title,
        pageUrl: window.location.href,
        pageContent: getPageContent(),
        service: currentService,
        selectedText: getSelectedText()
    };
}

/**
 * ページのメインコンテンツを取得
 */
function getPageContent() {
    // より良いコンテンツ抽出のため複数の候補を試す
    const selectors = [
        'main',
        'article',
        '.content',
        '.main-content',
        '#content',
        '#main',
        'body'
    ];

    for (const selector of selectors) {
        const element = document.querySelector(selector);
        if (element) {
            const text = element.innerText || element.textContent || '';
            if (text.length > 100) { // 最小限の内容があることを確認
                return text.substring(0, 3000); // 3000文字に制限
            }
        }
    }

    // フォールバック: body全体から取得
    return (document.body.innerText || document.body.textContent || '').substring(0, 3000);
}

/**
 * 現在のメールデータを取得
 */
function getCurrentEmailData() {
    let subject = '';
    let body = '';
    let sender = '';

    if (currentService === 'outlook') {
        // Outlook用の抽出ロジック
        const subjectElement = document.querySelector('[role="main"] [data-testid="message-subject"]');
        if (subjectElement) {
            subject = subjectElement.textContent.trim();
        }

        const bodyElement = document.querySelector('[role="main"] [data-testid="message-body"]');
        if (bodyElement) {
            body = bodyElement.textContent.trim();
        }

        const senderElement = document.querySelector('[role="main"] [data-testid="message-sender"]');
        if (senderElement) {
            sender = senderElement.textContent.trim();
        }
    } else if (currentService === 'gmail') {
        // Gmail用の抽出ロジック
        const subjectElement = document.querySelector('.hP');
        if (subjectElement) {
            subject = subjectElement.textContent.trim();
        }

        const bodyElement = document.querySelector('.ii.gt .a3s.aiL');
        if (bodyElement) {
            body = bodyElement.textContent.trim();
        }

        const senderElement = document.querySelector('.gD');
        if (senderElement) {
            sender = senderElement.getAttribute('email') || senderElement.textContent.trim();
        }
    }

    return {
        subject: subject,
        body: body,
        sender: sender,
        service: currentService,
        pageUrl: window.location.href,
        pageTitle: document.title
    };
}

/**
 * PTAダイアlogを作成（確実にモーダル表示）
 */
function createPTADialog(dialogData) {
    // 既存のダイアログを削除
    const existingDialog = document.getElementById('pta-dialog');
    if (existingDialog) {
        existingDialog.remove();
    }

    // ダイアログコンテナを作成
    const dialog = document.createElement('div');
    dialog.id = 'pta-dialog';
    dialog.className = 'pta-dialog';
    dialog.dialogData = dialogData; // データを保存

    // 強制的にbodyの最後に追加（iframeを回避）
    dialog.style.cssText = `
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        z-index: 2147483647 !important;
        pointer-events: auto !important;
        box-sizing: border-box !important;
        font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif !important;
        background: rgba(0, 0, 0, 0.7) !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
    `;

    // ダイアログHTML構造を作成
    dialog.innerHTML = `
        <div class="pta-dialog-content" style="
            background: white !important;
            border-radius: 12px !important;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3) !important;
            max-width: 600px !important;
            width: 90% !important;
            max-height: 80vh !important;
            overflow-y: auto !important;
            z-index: 2147483648 !important;
            pointer-events: auto !important;
            box-sizing: border-box !important;
            position: relative !important;
        ">
            <div class="pta-dialog-header" style="
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 16px 20px;
                border-bottom: 1px solid #e0e0e0;
                background: linear-gradient(135deg, #f5f5f5, #e8e8e8);
                border-radius: 12px 12px 0 0;
            ">
                <h3 style="margin: 0; color: #333; font-size: 18px; font-weight: 600;">🏫 PTA支援ツール</h3>
                <button class="pta-dialog-close" onclick="closePTADialog()" style="
                    background: none;
                    border: none;
                    font-size: 24px;
                    cursor: pointer;
                    color: #666;
                    padding: 0;
                    width: 32px;
                    height: 32px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    border-radius: 50%;
                    transition: background-color 0.2s;
                ">×</button>
            </div>
            <div class="pta-dialog-body" style="padding: 20px;">
                <div class="pta-dialog-info">
                    <h4 style="margin: 0 0 8px 0; color: #333;">📄 ${dialogData.pageTitle || 'ページ情報'}</h4>
                    <p class="pta-url" style="margin: 0 0 16px 0; color: #666; font-size: 12px; word-break: break-all;">${dialogData.pageUrl || ''}</p>
                    ${dialogData.selectedText ? `<div class="pta-selected-text" style="background: #f0f0f0; padding: 8px; border-radius: 4px; margin-bottom: 16px;"><strong>選択テキスト:</strong> ${dialogData.selectedText.substring(0, 100)}...</div>` : ''}
                </div>
                
                <div class="pta-dialog-actions" style="display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 16px;">
                    ${currentService === 'outlook' || currentService === 'gmail' ? `
                        <button class="pta-action-btn" onclick="analyzeEmail()" style="
                            background: linear-gradient(135deg, #2196F3, #1976D2);
                            color: white;
                            border: none;
                            border-radius: 6px;
                            padding: 10px 16px;
                            cursor: pointer;
                            font-size: 14px;
                            font-weight: 500;
                            transition: all 0.2s;
                        ">📧 メール解析</button>
                        <button class="pta-action-btn" onclick="composeReply()" style="
                            background: linear-gradient(135deg, #4CAF50, #388E3C);
                            color: white;
                            border: none;
                            border-radius: 6px;
                            padding: 10px 16px;
                            cursor: pointer;
                            font-size: 14px;
                            font-weight: 500;
                            transition: all 0.2s;
                        ">📝 返信作成</button>
                    ` : `
                        <button class="pta-action-btn" onclick="analyzePage()" style="
                            background: linear-gradient(135deg, #2196F3, #1976D2);
                            color: white;
                            border: none;
                            border-radius: 6px;
                            padding: 10px 16px;
                            cursor: pointer;
                            font-size: 14px;
                            font-weight: 500;
                            transition: all 0.2s;
                        ">📄 ページ要約</button>
                    `}
                    ${dialogData.selectedText ? `
                        <button class="pta-action-btn" onclick="analyzeSelection()" style="
                            background: linear-gradient(135deg, #FF9800, #F57C00);
                            color: white;
                            border: none;
                            border-radius: 6px;
                            padding: 10px 16px;
                            cursor: pointer;
                            font-size: 14px;
                            font-weight: 500;
                            transition: all 0.2s;
                        ">🔍 選択テキスト分析</button>
                    ` : ''}
                    <button class="pta-action-btn secondary" onclick="openSettings()" style="
                        background: #f5f5f5;
                        color: #666;
                        border: 1px solid #ddd;
                        border-radius: 6px;
                        padding: 10px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                        transition: all 0.2s;
                    ">⚙️ 設定</button>
                </div>
                
                <div class="pta-dialog-result">
                    <div id="pta-loading" class="pta-loading" style="display: none; text-align: center; padding: 20px;">
                        <div class="pta-spinner" style="
                            border: 4px solid #f3f3f3;
                            border-top: 4px solid #2196F3;
                            border-radius: 50%;
                            width: 40px;
                            height: 40px;
                            animation: spin 1s linear infinite;
                            margin: 0 auto 16px;
                        "></div>
                        <span style="color: #666;">AI処理中...</span>
                    </div>
                    <div id="pta-result" class="pta-result" style="display: none; background: #f9f9f9; padding: 16px; border-radius: 8px; border-left: 4px solid #2196F3;"></div>
                </div>
            </div>
        </div>
    `;

    // bodyに直接追加してiframeを回避
    document.body.appendChild(dialog);

    // スピナーアニメーション用のCSSを追加
    if (!document.getElementById('pta-spinner-style')) {
        const style = document.createElement('style');
        style.id = 'pta-spinner-style';
        style.textContent = `
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
        `;
        document.head.appendChild(style);
    } console.log('PTAダイアログを作成しました（モーダル表示）');
}

/**
 * ダイアログを閉じる
 */

// ESCキーで閉じる
document.addEventListener('keydown', handleEscapeKey);
}

/**
 * PTAダイアログを閉じる
 */
function closePTADialog() {
    const dialog = document.getElementById('pta-dialog');
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
        closePTADialog();
    }
}

/**
 * PTAダイアログのスタイルを追加
 */
function addPTADialogStyles() {
    // 既存のスタイルが存在する場合は何もしない
    if (document.getElementById('pta-dialog-styles')) {
        return;
    }

    const style = document.createElement('style');
    style.id = 'pta-dialog-styles';
    style.textContent = `
        .pta-dialog {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: 2147483647; /* 最高レベルのz-index */
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
        }

        .pta-dialog-overlay {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            backdrop-filter: blur(2px);
        }

        .pta-dialog-content {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 90%;
            max-width: 600px;
            max-height: 80vh;
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
            animation: ptaDialogShow 0.3s ease;
        }

        @keyframes ptaDialogShow {
            from {
                opacity: 0;
                transform: translate(-50%, -60%);
            }
            to {
                opacity: 1;
                transform: translate(-50%, -50%);
            }
        }

        .pta-dialog-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 20px 24px;
            background: linear-gradient(135deg, #2196F3, #1976D2);
            color: white;
        }

        .pta-dialog-header h3 {
            margin: 0;
            font-size: 18px;
            font-weight: 600;
        }

        .pta-dialog-close {
            background: none;
            border: none;
            color: white;
            font-size: 24px;
            cursor: pointer;
            padding: 0;
            width: 32px;
            height: 32px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background 0.2s;
        }

        .pta-dialog-close:hover {
            background: rgba(255, 255, 255, 0.1);
        }

        .pta-dialog-body {
            padding: 24px;
            max-height: 60vh;
            overflow-y: auto;
        }

        .pta-dialog-info {
            margin-bottom: 24px;
            padding-bottom: 16px;
            border-bottom: 1px solid #eee;
        }

        .pta-dialog-info h4 {
            margin: 0 0 8px 0;
            font-size: 16px;
            color: #333;
        }

        .pta-url {
            font-size: 14px;
            color: #666;
            margin: 0 0 12px 0;
            word-break: break-all;
        }

        .pta-selected-text {
            background: #f8f9fa;
            padding: 12px;
            border-radius: 6px;
            border-left: 3px solid #2196F3;
            font-size: 14px;
            color: #555;
        }

        .pta-dialog-actions {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 12px;
            margin-bottom: 24px;
        }

        .pta-action-btn {
            padding: 12px 16px;
            border: none;
            border-radius: 8px;
            background: #2196F3;
            color: white;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }

        .pta-action-btn:hover {
            background: #1976D2;
            transform: translateY(-1px);
        }

        .pta-action-btn.secondary {
            background: #666;
        }

        .pta-action-btn.secondary:hover {
            background: #555;
        }

        .pta-loading {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 12px;
            padding: 32px;
            color: #666;
        }

        .pta-spinner {
            width: 20px;
            height: 20px;
            border: 2px solid #e0e0e0;
            border-top: 2px solid #2196F3;
            border-radius: 50%;
            animation: ptaSpinner 1s linear infinite;
        }

        @keyframes ptaSpinner {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .pta-result {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 16px;
            font-size: 14px;
            line-height: 1.6;
            white-space: pre-wrap;
            max-height: 300px;
            overflow-y: auto;
        }

        .pta-result.error {
            background: #ffebee;
            color: #c62828;
            border: 1px solid #ffcdd2;
        }

        /* 支援ボタンのスタイル */
        .pta-support-button {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 2147483646;
            background: linear-gradient(135deg, #2196F3, #1976D2);
            color: white;
            border: none;
            border-radius: 50px;
            padding: 12px 20px;
            cursor: pointer;
            box-shadow: 0 4px 16px rgba(33, 150, 243, 0.3);
            transition: all 0.3s ease;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
        }

        .pta-support-button:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 24px rgba(33, 150, 243, 0.4);
        }

        .pta-button-content {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 14px;
            font-weight: 500;
        }

        .pta-icon {
            font-size: 16px;
        }
    `;

    document.head.appendChild(style);
}

/**
 * メール解析を実行
 */
function analyzeEmail() {
    const dialog = document.getElementById('pta-dialog');
    const emailData = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzeEmail',
        data: emailData
    }, (response) => {
        hideLoading();

        if (chrome.runtime.lastError) {
            showResult(`通信エラー: ${chrome.runtime.lastError.message}`, 'error');
            return;
        }

        if (response && response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response ? response.error : '不明なエラー'}`, 'error');
        }
    });
}

/**
 * ページ解析を実行
 */
function analyzePage() {
    const dialog = document.getElementById('pta-dialog');
    const pageData = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzePage',
        data: pageData
    }, (response) => {
        hideLoading();

        if (chrome.runtime.lastError) {
            showResult(`通信エラー: ${chrome.runtime.lastError.message}`, 'error');
            return;
        }

        if (response && response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response ? response.error : '不明なエラー'}`, 'error');
        }
    });
}

/**
 * 選択テキスト解析を実行
 */
function analyzeSelection() {
    const dialog = document.getElementById('pta-dialog');
    const pageData = dialog.dialogData;

    if (!pageData.selectedText) {
        showResult('選択されたテキストがありません。', 'error');
        return;
    }

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzeSelection',
        data: pageData
    }, (response) => {
        hideLoading();

        if (chrome.runtime.lastError) {
            showResult(`通信エラー: ${chrome.runtime.lastError.message}`, 'error');
            return;
        }

        if (response && response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response ? response.error : '不明なエラー'}`, 'error');
        }
    });
}

/**
 * 返信作成を実行
 */
function composeReply() {
    const dialog = document.getElementById('pta-dialog');
    const emailData = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'composeEmail',
        data: {
            type: 'reply',
            content: `件名「${emailData.subject}」への返信を作成してください。`,
            originalEmail: emailData
        }
    }, (response) => {
        hideLoading();

        if (chrome.runtime.lastError) {
            showResult(`通信エラー: ${chrome.runtime.lastError.message}`, 'error');
            return;
        }

        if (response && response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response ? response.error : '不明なエラー'}`, 'error');
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
 * ローディング表示
 */
function showLoading() {
    const loadingElement = document.getElementById('pta-loading');
    const resultElement = document.getElementById('pta-result');

    if (loadingElement) {
        loadingElement.style.display = 'flex';
    }
    if (resultElement) {
        resultElement.style.display = 'none';
    }
}

/**
 * ローディング非表示
 */
function hideLoading() {
    const loadingElement = document.getElementById('pta-loading');

    if (loadingElement) {
        loadingElement.style.display = 'none';
    }
}

/**
 * 結果表示
 */
function showResult(result, type = 'success') {
    const resultElement = document.getElementById('pta-result');

    if (resultElement) {
        resultElement.textContent = result;
        resultElement.className = `pta-result ${type}`;
        resultElement.style.display = 'block';
    }
}

/**
 * 通知を表示
 */
function showNotification(message, type = 'info') {
    // 既存の通知を削除
    const existingNotification = document.getElementById('pta-notification');
    if (existingNotification) {
        existingNotification.remove();
    }

    // 通知要素を作成
    const notification = document.createElement('div');
    notification.id = 'pta-notification';
    notification.className = `pta-notification ${type}`;
    notification.textContent = message;

    // 通知スタイルを追加
    if (!document.getElementById('pta-notification-styles')) {
        const style = document.createElement('style');
        style.id = 'pta-notification-styles';
        style.textContent = `
            .pta-notification {
                position: fixed;
                top: 20px;
                right: 20px;
                z-index: 2147483647;
                padding: 12px 20px;
                border-radius: 6px;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
                font-size: 14px;
                font-weight: 500;
                box-shadow: 0 4px 16px rgba(0, 0, 0, 0.15);
                animation: ptaNotificationSlide 0.3s ease;
            }

            @keyframes ptaNotificationSlide {
                from {
                    transform: translateX(100%);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }

            .pta-notification.success {
                background: #e8f5e8;
                color: #2e7d32;
                border: 1px solid #4CAF50;
            }

            .pta-notification.error {
                background: #ffebee;
                color: #c62828;
                border: 1px solid #f44336;
            }

            .pta-notification.info {
                background: #e3f2fd;
                color: #1565c0;
                border: 1px solid #2196F3;
            }
        `;
        document.head.appendChild(style);
    }

    // 通知を表示
    document.body.appendChild(notification);

    // 3秒後に自動削除
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
    }, 3000);
}

/**
 * ページ解析ハンドラー（コンテキストメニューから呼び出し）
 */
function handlePageAnalysis(data) {
    const pageData = {
        pageTitle: data.pageTitle,
        pageUrl: data.pageUrl,
        pageContent: getPageContent(),
        service: currentService
    };

    // ダイアログを作成してページ解析を実行
    createPTADialog(pageData);

    // 自動的にページ解析を開始
    setTimeout(() => {
        analyzePage();
    }, 500);
}

/**
 * 選択テキスト解析ハンドラー（コンテキストメニューから呼び出し）
 */
function handleSelectionAnalysis(data) {
    const pageData = {
        pageTitle: data.pageTitle,
        pageUrl: data.pageUrl,
        selectedText: data.selectedText,
        service: currentService
    };

    // ダイアログを作成して選択テキスト解析を実行
    createPTADialog(pageData);

    // 自動的に選択テキスト解析を開始
    setTimeout(() => {
        analyzeSelection();
    }, 500);
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
            if (ptaButton) {
                ptaButton.remove();
                ptaButton = null;
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

// グローバル関数として公開（ダイアログから呼び出し可能にするため）
window.analyzeEmail = analyzeEmail;
window.analyzePage = analyzePage;
window.analyzeSelection = analyzeSelection;
window.composeReply = composeReply;
window.openSettings = openSettings;