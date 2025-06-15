/*
 * Shima Edge拡張機能 - ポップアップスクリプト
 * Copyright (c) 2024 Shima Development Team
 */

// DOM読み込み完了時の初期化
document.addEventListener('DOMContentLoaded', function () {
    initializePopup();
});

/**
 * ポップアップの初期化
 */
function initializePopup() {
    // タブ切り替えイベント
    const tabButtons = document.querySelectorAll('.tab-button');
    tabButtons.forEach(button => {
        button.addEventListener('click', switchTab);
    });

    // クイック解析イベント
    document.getElementById('analyze-current-email').addEventListener('click', analyzeCurrentEmail);

    // クイック機能イベント
    document.getElementById('quick-reply').addEventListener('click', () => quickAction('reply'));
    document.getElementById('quick-summary').addEventListener('click', () => quickAction('summary'));
    document.getElementById('quick-action').addEventListener('click', () => quickAction('action'));

    // メール作成イベント
    document.getElementById('compose-email').addEventListener('click', composeEmail);

    // 履歴・設定イベント
    document.getElementById('clear-history').addEventListener('click', clearHistory);
    document.getElementById('open-settings').addEventListener('click', openSettings);
    document.getElementById('test-api').addEventListener('click', testAPI);

    // ページリンクコピーイベント
    document.getElementById('copy-page-link').addEventListener('click', copyPageLink);
    document.getElementById('copy-page-markdown').addEventListener('click', () => copyPageMarkdown());
    document.getElementById('copy-page-markdown-reply').addEventListener('click', () => copyPageMarkdown());
    document.getElementById('copy-page-markdown-summary').addEventListener('click', () => copyPageMarkdown());
    document.getElementById('copy-page-markdown-action').addEventListener('click', () => copyPageMarkdown());

    // 構造的コピーイベント
    document.getElementById('copy-structured-result').addEventListener('click', copyStructuredResult);

    // 初期データ読み込み
    loadPageInfo();
    loadHistory();
    checkAPISettings();
}

/**
 * タブ切り替え
 */
function switchTab(event) {
    const targetTab = event.target.getAttribute('data-tab');

    // すべてのタブボタンとコンテンツを非アクティブに
    document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    // 選択されたタブをアクティブに
    event.target.classList.add('active');
    document.getElementById(`${targetTab}-tab`).classList.add('active');

    // 履歴タブの場合は履歴を再読み込み
    if (targetTab === 'history') {
        loadHistory();
    }
}

/**
 * 現在のメールを解析
 */
async function analyzeCurrentEmail() {
    try {
        showLoading();

        // アクティブなタブから現在のメール情報を取得
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

        // コンテンツスクリプトからメールデータを取得
        const results = await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            function: getCurrentEmailData
        });

        const emailData = results[0].result;

        if (!emailData || !emailData.body) {
            throw new Error('メール内容を取得できませんでした。メールが表示されていることを確認してください。');
        }

        // バックグラウンドスクリプトに解析を依頼
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

    } catch (error) {
        hideLoading();
        showResult(`エラー: ${error.message}`, 'error');
    }
}

/**
 * クイック機能実行
 */
async function quickAction(actionType) {
    try {
        showLoading();

        // アクティブなタブから現在のメール情報を取得
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

        const results = await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            function: getCurrentEmailData
        });

        const emailData = results[0].result;

        if (!emailData || !emailData.body) {
            throw new Error('メール内容を取得できませんでした。');
        }

        let requestData;

        switch (actionType) {
            case 'reply':
                requestData = {
                    type: 'reply',
                    content: `件名「${emailData.subject}」への返信を作成してください。`,
                    originalEmail: emailData
                };
                break;

            case 'summary':
                requestData = {
                    type: 'summary',
                    content: `件名「${emailData.subject}」の要約を作成してください。本文: ${emailData.body.substring(0, 500)}`,
                    originalEmail: emailData
                };
                break;

            case 'action':
                requestData = {
                    type: 'action',
                    content: `件名「${emailData.subject}」に対して必要なアクションを提案してください。本文: ${emailData.body.substring(0, 500)}`,
                    originalEmail: emailData
                };
                break;
        }

        // バックグラウンドスクリプトに作成を依頼
        chrome.runtime.sendMessage({
            action: 'composeEmail',
            data: requestData
        }, (response) => {
            hideLoading();

            if (response.success) {
                showResult(response.result);
            } else {
                showResult(`エラー: ${response.error}`, 'error');
            }
        });

    } catch (error) {
        hideLoading();
        showResult(`エラー: ${error.message}`, 'error');
    }
}

/**
 * メール作成
 */
function composeEmail() {
    const type = document.getElementById('compose-type').value;
    const content = document.getElementById('compose-content').value.trim();

    if (!content) {
        showResult('内容を入力してください。', 'error');
        return;
    }

    showLoading();

    const requestData = {
        type: type,
        content: content
    };

    // バックグラウンドスクリプトに作成を依頼
    chrome.runtime.sendMessage({
        action: 'composeEmail',
        data: requestData
    }, (response) => {
        hideLoading();

        if (response.success) {
            showResult(response.result);
            // 成功したら入力フィールドをクリア
            document.getElementById('compose-content').value = '';
        } else {
            showResult(`エラー: ${response.error}`, 'error');
        }
    });
}

/**
 * 履歴読み込み
 */
function loadHistory() {
    chrome.storage.local.get(['ai_history'], (result) => {
        const history = result.ai_history || [];
        const historyList = document.getElementById('history-list');

        if (history.length === 0) {
            historyList.innerHTML = `
                <div class="empty-state">
                    <span class="empty-icon">📭</span>
                    <p>まだ履歴がありません</p>
                </div>
            `;
            return;
        }

        historyList.innerHTML = history.map(item => `
            <div class="history-item" onclick="showHistoryDetail(${item.id})">
                <div class="history-header">
                    <span class="history-type">${getHistoryTypeLabel(item.type)}</span>
                    <span class="history-date">${formatDate(item.timestamp)}</span>
                </div>
                <div class="history-content">
                    ${item.emailSubject || item.requestType || '内容'}
                </div>
            </div>
        `).join('');
    });
}

/**
 * 履歴詳細表示
 */
function showHistoryDetail(itemId) {
    chrome.storage.local.get(['ai_history'], (result) => {
        const history = result.ai_history || [];
        const item = history.find(h => h.id === itemId);

        if (item) {
            showResult(item.result);
        }
    });
}

/**
 * 履歴クリア
 */
function clearHistory() {
    if (confirm('履歴をすべて削除しますか？')) {
        chrome.storage.local.set({ 'ai_history': [] }, () => {
            loadHistory();
        });
    }
}

/**
 * 設定画面を開く
 */
function openSettings() {
    chrome.runtime.openOptionsPage();
}

/**
 * API接続テスト
 */
function testAPI() {
    showLoading();

    chrome.storage.local.get(['ai_settings'], (result) => {
        const settings = result.ai_settings || {};

        chrome.runtime.sendMessage({
            action: 'testApiConnection',
            data: settings
        }, (response) => {
            hideLoading();

            if (response.success) {
                showResult('API接続テストが成功しました！', 'success');
            } else {
                showResult(`API接続テストに失敗しました: ${response.error}`, 'error');
            }
        });
    });
}

/**
 * API設定チェック
 */
function checkAPISettings() {
    chrome.storage.local.get(['ai_settings'], (result) => {
        const settings = result.ai_settings || {};

        if (!settings.apiKey) {
            const warning = document.createElement('div');
            warning.className = 'api-warning';
            warning.innerHTML = `
                <p>⚠️ APIキーが設定されていません</p>
                <button onclick="openSettings()">設定画面を開く</button>
            `;

            document.querySelector('.popup-main').insertBefore(
                warning,
                document.querySelector('.tab-navigation')
            );
        }
    });
}

/**
 * 現在のページ情報を読み込み
 */
async function loadPageInfo() {
    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        const pageTitle = tab.title || 'ページタイトル不明';
        const pageUrl = tab.url || '';

        // ページタイトルを更新
        const titleElement = document.getElementById('current-page-title');
        titleElement.textContent = pageTitle;
        titleElement.title = `${pageTitle}\n${pageUrl}`;

        // グローバル変数として保存（他の関数で使用）
        window.currentPageInfo = {
            title: pageTitle,
            url: pageUrl
        };
    } catch (error) {
        console.error('ページ情報の取得に失敗:', error);
        document.getElementById('current-page-title').textContent = 'ページ情報取得エラー';
    }
}

/**
 * ページリンクをコピー
 */
async function copyPageLink() {
    try {
        if (!window.currentPageInfo) {
            await loadPageInfo();
        }

        const pageInfo = window.currentPageInfo;
        const linkText = `${pageInfo.title}\n${pageInfo.url}`;

        await navigator.clipboard.writeText(linkText);
        showNotification('ページリンクをコピーしました');
    } catch (error) {
        console.error('ページリンクのコピーに失敗:', error);
        showNotification('コピーに失敗しました', 'error');
    }
}

/**
 * ページ情報をMarkdown形式でコピー
 */
async function copyPageMarkdown() {
    try {
        if (!window.currentPageInfo) {
            await loadPageInfo();
        }

        const pageInfo = window.currentPageInfo;
        const markdownText = `[${pageInfo.title}](${pageInfo.url})`;

        await navigator.clipboard.writeText(markdownText);
        showNotification('Markdown形式でコピーしました');
    } catch (error) {
        console.error('Markdownコピーに失敗:', error);
        showNotification('コピーに失敗しました', 'error');
    }
}

/**
 * 構造的な結果をコピー
 */
async function copyStructuredResult() {
    try {
        const resultBody = document.getElementById('result-body');
        const resultText = resultBody.textContent || resultBody.innerText;

        if (!window.currentPageInfo) {
            await loadPageInfo();
        }

        const pageInfo = window.currentPageInfo;
        const structuredResult = `# AI解析結果\n\n## 対象ページ\n- タイトル: ${pageInfo.title}\n- URL: ${pageInfo.url}\n\n## 解析結果\n${resultText}\n\n---\n生成日時: ${new Date().toLocaleString('ja-JP')}`;

        await navigator.clipboard.writeText(structuredResult);
        showNotification('構造的な結果をコピーしました');
    } catch (error) {
        console.error('構造的コピーに失敗:', error);
        showNotification('コピーに失敗しました', 'error');
    }
}

/**
 * 通知を表示
 */
function showNotification(message, type = 'success') {
    // 既存の通知があれば削除
    const existingNotification = document.querySelector('.notification');
    if (existingNotification) {
        existingNotification.remove();
    }

    // 新しい通知を作成
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        top: 10px;
        right: 10px;
        background: ${type === 'error' ? '#f44336' : '#4caf50'};
        color: white;
        padding: 8px 12px;
        border-radius: 4px;
        font-size: 12px;
        z-index: 1000;
        animation: slideIn 0.3s ease;
    `;

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
    document.getElementById('loading-overlay').style.display = 'flex';
}

/**
 * ローディング非表示
 */
function hideLoading() {
    document.getElementById('loading-overlay').style.display = 'none';
}

/**
 * 結果表示
 */
function showResult(content, type = 'success') {
    const resultBody = document.getElementById('result-body');
    const resultOverlay = document.getElementById('result-overlay');

    resultBody.innerHTML = `<pre class="result-text ${type}">${content}</pre>`;
    resultOverlay.style.display = 'flex';

    // 結果を記録用のグローバル変数に保存
    window.currentResult = content;
}

/**
 * 結果画面を閉じる
 */
function closeResult() {
    document.getElementById('result-overlay').style.display = 'none';
}

/**
 * 結果をコピー
 */
function copyResult() {
    if (window.currentResult) {
        navigator.clipboard.writeText(window.currentResult).then(() => {
            // 一時的にボタンテキストを変更
            const copyBtn = document.querySelector('.result-actions .secondary');
            const originalText = copyBtn.innerHTML;
            copyBtn.innerHTML = '✅ コピー完了';

            setTimeout(() => {
                copyBtn.innerHTML = originalText;
            }, 1000);
        }).catch(() => {
            alert('コピーに失敗しました');
        });
    }
}

/**
 * 履歴タイプのラベル取得
 */
function getHistoryTypeLabel(type) {
    const labels = {
        'analysis': '📊 解析',
        'composition': '✍️ 作成',
        'reply': '💬 返信',
        'summary': '📝 要約',
        'action': '✅ アクション'
    };
    return labels[type] || '📋 記録';
}

/**
 * 日付フォーマット
 */
function formatDate(timestamp) {
    const date = new Date(timestamp);
    return date.toLocaleString('ja-JP', {
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

/**
 * 現在のメールデータを取得（コンテンツスクリプトで実行される関数）
 */
function getCurrentEmailData() {
    let currentService = 'unknown';
    if (window.location.hostname.includes('outlook.office.com') || window.location.hostname.includes('outlook.live.com')) {
        currentService = 'outlook';
    } else if (window.location.hostname.includes('mail.google.com')) {
        currentService = 'gmail';
    }

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

// グローバル関数として公開
window.closeResult = closeResult;
window.copyResult = copyResult;
window.showHistoryDetail = showHistoryDetail;