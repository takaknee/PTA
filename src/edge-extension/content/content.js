/*
 * PTA Edge拡張機能 - コンテンツスクリプト
 * Copyright (c) 2024 PTA Development Team
 */

// 現在のメールサービスを判定
let currentService = 'unknown';
if (window.location.hostname.includes('outlook.office.com') || window.location.hostname.includes('outlook.live.com')) {
    currentService = 'outlook';
} else if (window.location.hostname.includes('mail.google.com')) {
    currentService = 'gmail';
}

// PTA支援ボタンを追加
let ptaButton = null;

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
    }
    
    // URLの変更を監視（SPA対応）
    observeUrlChanges();
}

/**
 * Outlookの初期化
 */
function initializeOutlook() {
    // Outlookのメール読み込み完了を待機
    const observer = new MutationObserver(() => {
        const emailContent = document.querySelector('[role="main"] [role="document"]');
        if (emailContent && !ptaButton) {
            addPTAButton();
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
            addPTAButton();
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
            addPTAButton();
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
            addPTAButton();
        }
    }, 2000);
}

/**
 * PTA支援ボタンを追加
 */
function addPTAButton() {
    // 既存のボタンを削除
    const existingButton = document.getElementById('pta-support-button');
    if (existingButton) {
        existingButton.remove();
    }
    
    // ボタンを作成
    ptaButton = document.createElement('div');
    ptaButton.id = 'pta-support-button';
    ptaButton.className = 'pta-support-button';
    ptaButton.innerHTML = `
        <div class="pta-button-content">
            <span class="pta-icon">🏫</span>
            <span class="pta-text">PTA支援</span>
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
    // 現在のメール内容を取得
    const emailData = getCurrentEmailData();
    
    if (!emailData.body) {
        showNotification('メール内容を取得できませんでした。', 'error');
        return;
    }
    
    // ダイアログを作成
    createPTADialog(emailData);
}

/**
 * 現在のメールデータを取得
 */
function getCurrentEmailData() {
    let emailData = {
        subject: '',
        sender: '',
        body: '',
        service: currentService
    };
    
    if (currentService === 'outlook') {
        // Outlookからメールデータを取得
        const subjectElement = document.querySelector('[role="main"] h1');
        const senderElement = document.querySelector('[role="main"] [data-testid="sender-name"]');
        const bodyElement = document.querySelector('[role="main"] [role="document"]');
        
        emailData.subject = subjectElement ? subjectElement.textContent.trim() : '';
        emailData.sender = senderElement ? senderElement.textContent.trim() : '';
        emailData.body = bodyElement ? bodyElement.textContent.trim() : '';
        
    } else if (currentService === 'gmail') {
        // Gmailからメールデータを取得
        const subjectElement = document.querySelector('.hP');
        const senderElement = document.querySelector('.go .g2 .gD');
        const bodyElement = document.querySelector('.ii.gt .a3s.aiL');
        
        emailData.subject = subjectElement ? subjectElement.textContent.trim() : '';
        emailData.sender = senderElement ? senderElement.textContent.trim() : '';
        emailData.body = bodyElement ? bodyElement.textContent.trim() : '';
    }
    
    return emailData;
}

/**
 * PTAダイアログを作成
 */
function createPTADialog(emailData) {
    // 既存のダイアログを削除
    const existingDialog = document.getElementById('pta-dialog');
    if (existingDialog) {
        existingDialog.remove();
    }
    
    // ダイアログを作成
    const dialog = document.createElement('div');
    dialog.id = 'pta-dialog';
    dialog.className = 'pta-dialog';
    dialog.innerHTML = `
        <div class="pta-dialog-content">
            <div class="pta-dialog-header">
                <h2>🏫 PTA支援ツール</h2>
                <button class="pta-close-button" onclick="this.closest('.pta-dialog').remove()">×</button>
            </div>
            <div class="pta-dialog-body">
                <div class="pta-email-info">
                    <h3>📧 メール情報</h3>
                    <p><strong>件名:</strong> ${emailData.subject || '（件名なし）'}</p>
                    <p><strong>送信者:</strong> ${emailData.sender || '（送信者不明）'}</p>
                    <p><strong>本文:</strong> ${emailData.body.substring(0, 100)}${emailData.body.length > 100 ? '...' : ''}</p>
                </div>
                <div class="pta-actions">
                    <button class="pta-action-button" onclick="analyzeEmail()">📊 メール解析</button>
                    <button class="pta-action-button" onclick="composeReply()">✍️ 返信作成</button>
                    <button class="pta-action-button" onclick="openSettings()">⚙️ 設定</button>
                </div>
                <div class="pta-result" id="pta-result" style="display: none;">
                    <h3>結果</h3>
                    <div id="pta-result-content"></div>
                </div>
                <div class="pta-loading" id="pta-loading" style="display: none;">
                    <div class="pta-spinner"></div>
                    <p>AI処理中...</p>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(dialog);
    
    // ダイアログに現在のメールデータを保存
    dialog.emailData = emailData;
}

/**
 * メール解析を実行
 */
function analyzeEmail() {
    const dialog = document.getElementById('pta-dialog');
    const emailData = dialog.emailData;
    
    showLoading();
    
    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzeEmail',
        data: emailData
    }, (response) => {
        hideLoading();
        
        if (response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response.error}`, 'error');
        }
    });
}

/**
 * 返信作成を実行
 */
function composeReply() {
    const dialog = document.getElementById('pta-dialog');
    const emailData = dialog.emailData;
    
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
        
        if (response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response.error}`, 'error');
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
    
    loadingElement.style.display = 'block';
    resultElement.style.display = 'none';
}

/**
 * ローディング非表示
 */
function hideLoading() {
    const loadingElement = document.getElementById('pta-loading');
    loadingElement.style.display = 'none';
}

/**
 * 結果表示
 */
function showResult(content, type = 'success') {
    const resultElement = document.getElementById('pta-result');
    const resultContent = document.getElementById('pta-result-content');
    
    resultContent.innerHTML = `<pre>${content}</pre>`;
    resultElement.className = `pta-result ${type}`;
    resultElement.style.display = 'block';
}

/**
 * 通知表示
 */
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `pta-notification ${type}`;
    notification.textContent = message;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.remove();
    }, 3000);
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
window.composeReply = composeReply;
window.openSettings = openSettings;