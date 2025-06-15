/*
 * PTA Edge拡張機能 - コンテンツスクリプト（修正版）
 * Copyright (c) 2024 PTA Development Team
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
 * AI支援ボタンを追加
 */
function addAISupportButton() {
    // 既存のボタンがあれば削除
    if (ptaButton) {
        ptaButton.remove();
    }

    // ボタンを作成
    ptaButton = document.createElement('button');
    ptaButton.className = 'ai-support-button';
    ptaButton.innerHTML = `
        <div class="ai-button-content">
            <span class="ai-icon">🏫</span>
            <span class="ai-text">PTA支援</span>
        </div>
    `;

    // 基本スタイルを設定
    ptaButton.style.cssText = `
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
    ptaButton.addEventListener('mouseenter', () => {
        ptaButton.style.transform = 'translateY(-2px)';
        ptaButton.style.boxShadow = '0 6px 16px rgba(33, 150, 243, 0.4)';
    });

    ptaButton.addEventListener('mouseleave', () => {
        ptaButton.style.transform = 'translateY(0)';
        ptaButton.style.boxShadow = '0 4px 12px rgba(33, 150, 243, 0.3)';
    });

    // クリックイベント
    ptaButton.addEventListener('click', showPTADialog);

    // ページに追加
    document.body.appendChild(ptaButton);

    console.log('PTA支援ボタンを追加しました');
}

/**
 * PTAダイアログを表示
 */
function showPTADialog() {
    const selectedText = getSelectedText();
    const dialogData = {
        pageTitle: document.title,
        pageUrl: window.location.href,
        selectedText: selectedText,
        currentService: currentService
    };

    // メール情報が利用可能な場合は追加
    if (currentService === 'outlook' || currentService === 'gmail') {
        const emailData = getCurrentEmailData();
        Object.assign(dialogData, emailData);
    }

    createPTADialog(dialogData);
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
    `;

    // ダイアログコンテンツを作成
    const content = document.createElement('div');
    content.style.cssText = `
        background: white !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3) !important;
        max-width: 600px !important;
        width: 90% !important;
        max-height: 80vh !important;
        overflow-y: auto !important;
        position: relative !important;
    `;
    content.innerHTML = `
        <div class="pta-dialog-header" style="
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 20px;
            border-bottom: 1px solid #e0e0e0;
            background: linear-gradient(135deg, #2196F3, #1976D2);
            color: white;
            border-radius: 12px 12px 0 0;
        ">
            <h3 style="margin: 0; font-size: 18px; font-weight: 600;">🏫 PTA支援ツール</h3>
            <button class="pta-close-btn" style="
                background: none;
                border: none;
                color: white;
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
            <div style="margin-bottom: 20px;">
                <h4 style="margin: 0 0 8px 0; color: #333;">📄 ${dialogData.pageTitle || 'ページ情報'}</h4>
                <p style="margin: 0 0 16px 0; color: #666; font-size: 12px; word-break: break-all;">${dialogData.pageUrl || ''}</p>
                ${dialogData.selectedText ? `<div style="background: #f0f8ff; padding: 12px; border-radius: 6px; border-left: 4px solid #2196F3; margin-bottom: 16px;"><strong>選択テキスト:</strong> ${dialogData.selectedText.substring(0, 100)}...</div>` : ''}
            </div>
            
            <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px;">
                ${currentService === 'outlook' || currentService === 'gmail' ? `
                    <button class="pta-analyze-email-btn" style="
                        background: linear-gradient(135deg, #2196F3, #1976D2);
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 12px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                    ">📧 メール解析</button>
                    <button class="pta-compose-reply-btn" style="
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
                    <button class="pta-analyze-page-btn" style="
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
                    <button class="pta-analyze-selection-btn" style="
                        background: linear-gradient(135deg, #FF9800, #F57C00);
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 12px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                    ">🔍 選択テキスト分析</button>
                ` : ''}
                <button class="pta-open-settings-btn" style="
                    background: #f5f5f5;
                    color: #666;
                    border: 1px solid #ddd;
                    border-radius: 6px;
                    padding: 12px 16px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 500;
                ">⚙️ 設定</button>
            </div>
            
            <div>
                <div id="pta-loading" style="display: none; text-align: center; padding: 20px; color: #666;">
                    <div style="
                        border: 4px solid #f3f3f3;
                        border-top: 4px solid #2196F3;
                        border-radius: 50%;
                        width: 40px;
                        height: 40px;
                        animation: spin 1s linear infinite;
                        margin: 0 auto 16px;
                    "></div>
                    <span>AI処理中...</span>
                </div>
                <div id="pta-result" style="display: none; background: #f9f9f9; padding: 16px; border-radius: 8px; border-left: 4px solid #2196F3;"></div>
            </div>
        </div>
    `;

    dialog.appendChild(content);

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
    }
    // bodyに直接追加してiframeを回避
    document.body.appendChild(dialog);

    // CSP準拠のイベントリスナー設定
    setupDialogEventListeners(dialog);

    // オーバーレイクリックで閉じる
    dialog.addEventListener('click', (e) => {
        if (e.target === dialog) {
            closePTADialog();
        }
    });

    // ESCキーで閉じる
    document.addEventListener('keydown', handleEscapeKey);

    console.log('PTAダイアログを作成しました（モーダル表示）');
}

/**
 * ダイアログのイベントリスナーを設定（CSP準拠）
 */
function setupDialogEventListeners(dialog) {
    // 閉じるボタン
    const closeBtn = dialog.querySelector('.pta-close-btn');
    if (closeBtn) {
        closeBtn.addEventListener('click', closePTADialog);
    }

    // メール解析ボタン
    const analyzeEmailBtn = dialog.querySelector('.pta-analyze-email-btn');
    if (analyzeEmailBtn) {
        analyzeEmailBtn.addEventListener('click', analyzeEmail);
    }

    // 返信作成ボタン
    const composeReplyBtn = dialog.querySelector('.pta-compose-reply-btn');
    if (composeReplyBtn) {
        composeReplyBtn.addEventListener('click', composeReply);
    }

    // ページ解析ボタン
    const analyzePageBtn = dialog.querySelector('.pta-analyze-page-btn');
    if (analyzePageBtn) {
        analyzePageBtn.addEventListener('click', analyzePage);
    }

    // 選択テキスト解析ボタン
    const analyzeSelectionBtn = dialog.querySelector('.pta-analyze-selection-btn');
    if (analyzeSelectionBtn) {
        analyzeSelectionBtn.addEventListener('click', analyzeSelection);
    }
    // 設定ボタン
    const openSettingsBtn = dialog.querySelector('.pta-open-settings-btn');
    if (openSettingsBtn) {
        openSettingsBtn.addEventListener('click', openSettings);
    }
}

/**
 * ダイアログを閉じる
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
    const existingNotification = document.getElementById('pta-notification');
    if (existingNotification) {
        existingNotification.remove();
    }

    const notification = document.createElement('div');
    notification.id = 'pta-notification';
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
    const loadingElement = document.getElementById('pta-loading');
    const resultElement = document.getElementById('pta-result');

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
    const loadingElement = document.getElementById('pta-loading');

    if (loadingElement) {
        loadingElement.style.display = 'none';
    }
}

/**
 * 結果を表示
 */
function showResult(result) {
    const resultElement = document.getElementById('pta-result');

    if (resultElement) {
        resultElement.innerHTML = result;
        resultElement.style.display = 'block';
    }
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
    const dialog = document.getElementById('pta-dialog');
    const pageData = dialog.dialogData;

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
    const dialog = document.getElementById('pta-dialog');
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
    const dialog = document.getElementById('pta-dialog');
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
    const pageData = {
        pageTitle: document.title,
        pageUrl: window.location.href,
        pageContent: document.body.innerText.substring(0, 5000) // 最初の5000字
    };

    createPTADialog(pageData);
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

    createPTADialog(selectionData);
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
window.closePTADialog = closePTADialog;

console.log('PTA支援ツール - コンテンツスクリプト読み込み完了');
