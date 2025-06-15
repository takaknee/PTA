/*
 * AI業務支援ツール Edge拡張機能 - コンテンツスクリプト
 * Copyright (c) 2024 AI Business Support Team
 */

// 現在のサービス/ページを判定
let currentService = 'unknown';
if (window.location.hostname.includes('outlook.office.com') || window.location.hostname.includes('outlook.live.com')) {
    currentService = 'outlook';
} else if (window.location.hostname.includes('mail.google.com')) {
    currentService = 'gmail';
} else {
    currentService = 'general'; // 一般的なWebページ
}

// AI業務支援ボタンを追加
let aiSupportButton = null;

// バックグラウンドスクリプトからのメッセージを受信
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
        case 'analyzePage':
            handlePageAnalysis(message.data);
            break;
        case 'analyzeSelection':
            handleSelectionAnalysis(message.data);
            break;
        case 'translateSelection':
            handleSelectionTranslation(message.data);
            break;
        case 'translatePage':
            handlePageTranslation(message.data);
            break;
        case 'extractUrls':
            handleUrlExtraction(message.data);
            break;
        case 'copyPageInfo':
            handlePageInfoCopy(message.data);
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
    console.log('AI業務支援ツール初期化開始:', currentService);

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
        if (emailContent && !aiSupportButton) {
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
        if (emailContent && !aiSupportButton) {
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
        if (emailContent && !aiSupportButton) {
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
        if (emailContent && !aiSupportButton) {
            addAISupportButton();
        }
    }, 2000);
}

/**
 * AI業務支援ボタンを追加
 */
function addAISupportButton() {
    // 既存のボタンを削除
    const existingButton = document.getElementById('ai-support-button');
    if (existingButton) {
        existingButton.remove();
    }

    // ボタンを作成
    aiSupportButton = document.createElement('div');
    aiSupportButton.id = 'ai-support-button';
    aiSupportButton.className = 'ai-support-button';

    // サービスに応じてボタンテキストを変更
    let buttonText = 'AI支援';
    if (currentService === 'outlook' || currentService === 'gmail') {
        buttonText = 'メール支援';
    } else {
        buttonText = 'ページ分析';
    }

    aiSupportButton.innerHTML = `
        <div class="ai-button-content">
            <span class="ai-icon">🤖</span>
            <span class="ai-text">${buttonText}</span>
        </div>
    `;

    // ボタンクリックイベント
    aiSupportButton.addEventListener('click', openAIDialog);

    // ボタンを適切な位置に配置
    insertAISupportButton();
}

/**
 * サービスに応じてボタンを配置
 */
function insertAISupportButton() {
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
        aiSupportButton.style.position = 'fixed';
        aiSupportButton.style.top = '10px';
        aiSupportButton.style.right = '10px';
        aiSupportButton.style.zIndex = '10000';
        document.body.appendChild(aiSupportButton);
    } else {
        // フォールバック: 画面右上に配置
        aiSupportButton.style.position = 'fixed';
        aiSupportButton.style.top = '10px';
        aiSupportButton.style.right = '10px';
        aiSupportButton.style.zIndex = '10000';
        document.body.appendChild(aiSupportButton);
    }
}

/**
 * AIダイアログを開く
 */
function openAIDialog() {
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
    createAIDialog(dialogData);
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
 * 選択されたテキストを取得
 */
function getSelectedText() {
    const selection = window.getSelection();
    return selection.toString().trim();
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
 * AIダイアログを作成
 */
function createAIDialog(data) {
    // 既存のダイアログを削除
    const existingDialog = document.getElementById('ai-dialog-overlay');
    if (existingDialog) {
        existingDialog.remove();
    }
    // ダイアログを作成
    const dialog = document.createElement('div');
    dialog.id = 'ai-dialog-overlay';
    dialog.className = 'ai-dialog-overlay';

    let contentHtml = '';
    let actionsHtml = '';

    if (currentService === 'outlook' || currentService === 'gmail') {
        // メールサービスの場合
        contentHtml = `
            <div class="ai-content-info">
                <h3>📧 メール情報</h3>
                <p><strong>件名:</strong> ${data.subject || '（件名なし）'}</p>
                <p><strong>送信者:</strong> ${data.sender || '（送信者不明）'}</p>
                <p><strong>本文:</strong> ${(data.body || '').substring(0, 100)}${(data.body || '').length > 100 ? '...' : ''}</p>
            </div>
        `;
        actionsHtml = `
            <button class="ai-action-button" onclick="analyzeEmail()">📊 メール解析</button>
            <button class="ai-action-button" onclick="composeReply()">✍️ 返信作成</button>
        `;
    } else {
        // 一般的なWebページの場合
        const hasSelection = data.selectedText && data.selectedText.length > 0;
        contentHtml = `
            <div class="ai-content-info">
                <h3>📄 ページ情報</h3>
                <p><strong>タイトル:</strong> ${data.pageTitle || data.title || '（タイトルなし）'}</p>
                <p><strong>URL:</strong> ${data.pageUrl || data.url || ''}</p>
                ${hasSelection ? `<p><strong>選択テキスト:</strong> ${data.selectedText.substring(0, 100)}${data.selectedText.length > 100 ? '...' : ''}</p>` : ''}
                ${data.content ? `<p><strong>コンテンツ:</strong> ${data.content.substring(0, 100)}${data.content.length > 100 ? '...' : ''}</p>` : ''}
            </div>
        `;

        if (data.action === 'translate') {
            // 翻訳モードの場合
            actionsHtml = `
                <button class="ai-action-button" onclick="translateText()">🌐 翻訳実行</button>
            `;
        } else if (data.action === 'translatePage') {
            // ページ翻訳モードの場合
            actionsHtml = `
                <button class="ai-action-button" onclick="translatePage()">🌐 ページ翻訳</button>
            `;        } else {
            // 通常の分析モード
            actionsHtml = `
                <button class="ai-action-button" onclick="analyzePage()">📄 ページ要約</button>
                ${hasSelection ? '<button class="ai-action-button" onclick="analyzeSelection()">📝 選択文分析</button>' : ''}
            `;
        }
    }

    dialog.innerHTML = `
        <div class="ai-dialog-content">
            <div class="ai-dialog-header">
                <h2>🤖 AI業務支援ツール</h2>
                <button class="ai-close-button" onclick="this.closest('.ai-dialog').remove()">×</button>
            </div>
            <div class="ai-dialog-body">
                ${contentHtml}
                <div class="ai-actions">
                    ${actionsHtml}
                    <button class="ai-action-button" onclick="openSettings()">⚙️ 設定</button>
                </div>
                <div class="ai-result" id="ai-result" style="display: none;">
                    <h3>結果</h3>
                    <div id="ai-result-content"></div>
                </div>                <div class="ai-loading" id="ai-loading" style="display: none;">
                    <div class="ai-spinner"></div>
                    <p>AI処理中...</p>
                </div>
            </div>
        </div>
    `;

    document.body.appendChild(dialog);

    // 閉じるボタンのイベントリスナーを追加（右クリックメニューから開いたダイアログが閉じられない問題を修正）
    const closeButton = dialog.querySelector('#ai-close-btn');
    if (closeButton) {
        closeButton.addEventListener('click', function () {
            dialog.remove();
        });
    }

    // オーバーレイクリックで閉じる
    dialog.addEventListener('click', function (e) {
        if (e.target === dialog) {
            dialog.remove();
        }
    });

    // ESCキーで閉じる
    const escHandler = function (e) {
        if (e.key === 'Escape') {
            dialog.remove();
            document.removeEventListener('keydown', escHandler);
        }
    };
    document.addEventListener('keydown', escHandler);

    // ダイアログに現在のデータを保存
    dialog.dialogData = data;
}

/**
 * メール解析を実行
 */
function analyzeEmail() {
    const dialog = document.getElementById('ai-dialog-overlay');
    const emailData = dialog.dialogData;

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
 * ページ解析を実行
 */
function analyzePage() {
    const dialog = document.getElementById('ai-dialog');
    const pageData = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'analyzePage',
        data: pageData
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
 * 選択テキスト解析を実行
 */
function analyzeSelection() {
    const dialog = document.getElementById('ai-dialog');
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
    const dialog = document.getElementById('ai-dialog');
    const emailData = dialog.dialogData;

    showLoading();    // バックグラウンドスクリプトにメッセージを送信
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
    const loadingElement = document.getElementById('ai-loading');
    const resultElement = document.getElementById('ai-result');

    if (loadingElement) loadingElement.style.display = 'block';
    if (resultElement) resultElement.style.display = 'none';
}

/**
 * ローディング非表示
 */
function hideLoading() {
    const loadingElement = document.getElementById('ai-loading');
    if (loadingElement) loadingElement.style.display = 'none';
}

/**
 * 結果表示
 */
function showResult(content, type = 'success') {
    const resultElement = document.getElementById('ai-result');
    const resultContent = document.getElementById('ai-result-content');

    if (resultContent) {
        resultContent.innerHTML = `<pre>${content}</pre>`;
    }
    if (resultElement) {
        resultElement.className = `ai-result ${type}`;
        resultElement.style.display = 'block';
    }
}

/**
 * 通知表示
 */
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `ai-notification ${type}`;
    notification.textContent = message;

    document.body.appendChild(notification);

    setTimeout(() => {
        notification.remove();
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
    };    // ダイアログを作成してページ解析を実行
    createAIDialog(pageData);

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

    // ダイアログを作成して選択テキスト解析を実行    createAIDialog(pageData);

    // 自動的に選択テキスト解析を開始
    setTimeout(() => {
        analyzeSelection();
    }, 500);
}

/**
 * 選択テキストの翻訳処理
 */
function handleSelectionTranslation(data) {
    if (!data.selectedText) return;

    // 翻訳用のダイアログデータを作成
    const pageData = {
        title: data.pageTitle || document.title,
        url: data.pageUrl || window.location.href,
        type: 'translation',
        content: data.selectedText,
        action: 'translate'
    };

    // 翻訳処理を実行
    createAndShowDialog(pageData);
}

/**
 * ページ全体の翻訳処理
 */
function handlePageTranslation(data) {
    // ページ全体のテキストを取得
    let pageText = getPageContent();

    const pageData = {
        title: data.pageTitle || document.title,
        url: data.pageUrl || window.location.href,
        type: 'translation',
        content: pageText,
        action: 'translatePage'
    };

    createAndShowDialog(pageData);
}

/**
 * URL抽出＆コピー処理
 */
function handleUrlExtraction(data) {
    // ページ内のリンクを抽出
    const links = document.querySelectorAll('a[href]');
    const urls = [];

    links.forEach(link => {
        const href = link.href;
        const text = link.textContent.trim();
        if (href && href.startsWith('http') && text) {
            urls.push({
                url: href,
                text: text
            });
        }
    });

    // Markdown形式に変換
    let markdownText = '# 抽出されたURL一覧\n\n';
    urls.forEach(link => {
        markdownText += `- [${link.text}](${link.url})\n`;
    });

    // クリップボードにコピー
    copyToClipboard(markdownText);
    showNotification('URLをMarkdown形式でクリップボードにコピーしました');
}

/**
 * ページ情報コピー処理
 */
function handlePageInfoCopy(data) {
    const title = data.pageTitle || document.title;
    const url = data.pageUrl || window.location.href;
    const summary = getPageSummary(); // ページの要約を取得

    // Markdown形式でページ情報を作成
    const markdownText = `# ${title}

**URL**: ${url}

## 要約
${summary}
`;

    // クリップボードにコピー
    copyToClipboard(markdownText);
    showNotification('ページ情報をMarkdown形式でクリップボードにコピーしました');
}

/**
 * ダイアログを作成して表示する（右クリックメニュー用）
 */
function createAndShowDialog(pageData) {
    // 既存のダイアログがあれば削除
    const existingDialog = document.getElementById('ai-dialog-overlay');
    if (existingDialog) {
        existingDialog.remove();
    }    // ダイアログを作成
    createAIDialog(pageData);

    // 自動的に処理を開始
    setTimeout(() => {
        if (pageData.action === 'translate') {
            translateText();
        } else if (pageData.action === 'translatePage') {
            translatePage();
        }
    }, 500);
}

/**
 * ページの簡単な要約を取得
 */
function getPageSummary() {
    // メインコンテンツを取得
    let content = '';
    const mainElements = document.querySelectorAll('main, article, .content, .main-content, [role="main"]');

    if (mainElements.length > 0) {
        content = mainElements[0].textContent;
    } else {
        // フォールバック：bodyの最初の数段落
        const paragraphs = document.querySelectorAll('p');
        for (let i = 0; i < Math.min(3, paragraphs.length); i++) {
            content += paragraphs[i].textContent + '\n';
        }
    }

    // 300文字に制限
    return content.slice(0, 300).trim() + (content.length > 300 ? '...' : '');
}

/**
 * クリップボードにテキストをコピー
 */
async function copyToClipboard(text) {
    try {
        await navigator.clipboard.writeText(text);
    } catch (err) {
        // フォールバック：テキストエリアを使用
        const textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
    }
}

/**
 * 通知を表示
 */
function showNotification(message) {
    // 既存の通知があれば削除
    const existingNotification = document.getElementById('pta-notification');
    if (existingNotification) {
        existingNotification.remove();
    }

    // 通知要素を作成
    const notification = document.createElement('div');
    notification.id = 'pta-notification';
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #4CAF50;
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 10001;
        font-size: 14px;
        max-width: 300px;
        word-wrap: break-word;
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
 * URL変更を監視（SPA対応）
 */
function observeUrlChanges() {
    let currentUrl = window.location.href;

    const observer = new MutationObserver(() => {
        if (window.location.href !== currentUrl) {
            currentUrl = window.location.href;
            // URL変更時に既存のボタンを削除
            if (aiSupportButton) {
                aiSupportButton.remove();
                aiSupportButton = null;
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

/**
 * 翻訳実行（選択テキスト用）
 */
function translateText() {
    const dialog = document.getElementById('ai-dialog-overlay');
    const data = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'translateSelection',
        data: {
            selectedText: data.selectedText || data.content,
            pageUrl: data.url || data.pageUrl,
            pageTitle: data.title || data.pageTitle
        }
    }, (response) => {
        hideLoading();

        if (response && response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response ? response.error : '不明なエラーが発生しました'}`, 'error');
        }
    });
}

/**
 * ページ翻訳実行
 */
function translatePage() {
    const dialog = document.getElementById('ai-dialog-overlay');
    const data = dialog.dialogData;

    showLoading();

    // バックグラウンドスクリプトにメッセージを送信
    chrome.runtime.sendMessage({
        action: 'translatePage',
        data: {
            content: data.content || getPageContent(),
            pageUrl: data.url || data.pageUrl,
            pageTitle: data.title || data.pageTitle
        }
    }, (response) => {
        hideLoading();

        if (response && response.success) {
            showResult(response.result);
        } else {
            showResult(`エラー: ${response ? response.error : '不明なエラーが発生しました'}`, 'error');
        }
    });
}

/**
 * 結果をコピー
 */
function copyResult() {
    const resultContent = document.getElementById('ai-result-content');
    if (resultContent) {
        copyToClipboard(resultContent.textContent);
        showNotification('結果をクリップボードにコピーしました');
    }
}
