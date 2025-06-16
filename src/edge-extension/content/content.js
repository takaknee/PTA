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
        case 'forwardToTeams':
            handleForwardToTeamsFromContext(message.data);
            break;
        case 'addToCalendar':
            handleAddToCalendarFromContext(message.data);
            break;
        case 'analyzeVSCodeSettings':
            handleAnalyzeVSCodeSettingsFromContext(message.data);
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
            <span class="ai-drag-handle">⋮⋮</span>
        </div>
    `;    // 保存された位置を取得、なければデフォルト位置
    const savedPosition = getSavedButtonPositionSync();    // 基本スタイルを設定
    aiButton.style.cssText = `
        position: fixed !important;
        top: ${savedPosition.top}px !important;
        right: ${savedPosition.right}px !important;
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
        user-select: none !important;
    `;

    // ドラッグハンドルのスタイル
    const dragHandle = aiButton.querySelector('.ai-drag-handle');
    if (dragHandle) {
        dragHandle.style.cssText = `
            margin-left: 8px !important;
            opacity: 0.7 !important;
            font-size: 12px !important;
            cursor: move !important;
            padding: 2px !important;
            border-radius: 2px !important;
            display: inline-block !important;
        `;
    }

    // ホバー効果
    aiButton.addEventListener('mouseenter', () => {
        if (!aiButton.isDragging) {
            aiButton.style.transform = 'translateY(-2px)';
            aiButton.style.boxShadow = '0 6px 16px rgba(33, 150, 243, 0.4)';
        }
    });

    aiButton.addEventListener('mouseleave', () => {
        if (!aiButton.isDragging) {
            aiButton.style.transform = 'translateY(0)';
            aiButton.style.boxShadow = '0 4px 12px rgba(33, 150, 243, 0.3)';
        }
    });    // ドラッグ機能をハンドル部分のみに追加
    if (dragHandle) {
        makeDraggable(aiButton, dragHandle);
    }

    // クリックイベント（ドラッグ中でない場合のみ）
    aiButton.addEventListener('click', (e) => {
        // ドラッグハンドルがクリックされた場合は何もしない
        if (e.target.classList.contains('ai-drag-handle') || e.target.closest('.ai-drag-handle')) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }

        if (!aiButton.isDragging) {
            showAiDialog();
        }
    });

    // ページに追加
    document.body.appendChild(aiButton);

    console.log('AI支援ボタンを追加しました（移動可能版）');
}

/**
 * AIダイアログを表示
 */
function showAiDialog() {
    try {
        const selectedText = getSelectedText();
        let pageContent = '';

        // ページコンテンツを安全に取得
        try {
            pageContent = extractPageContent();
        } catch (contentError) {
            console.error('ページコンテンツ抽出でエラー:', contentError);
            pageContent = `コンテンツ抽出エラー: ${contentError.message}`;
        }

        const dialogData = {
            pageTitle: document.title || 'タイトル不明',
            pageUrl: window.location.href,
            pageContent: pageContent,
            selectedText: selectedText,
            currentService: currentService
        };

        // メール情報が利用可能な場合は追加
        if (currentService === 'outlook' || currentService === 'gmail') {
            try {
                const emailData = getCurrentEmailData();
                Object.assign(dialogData, emailData);
            } catch (emailError) {
                console.error('メール情報取得でエラー:', emailError);
                // メール情報の取得に失敗してもダイアログは表示
            }
        }

        // ダイアログを作成して表示
        createAiDialog(dialogData);

    } catch (error) {
        console.error('AIダイアログ表示でエラー:', error);
        console.error('エラー詳細:', error.stack);

        // エラーが発生した場合も最低限のダイアログを表示
        try {
            const fallbackData = {
                pageTitle: document.title || 'エラー',
                pageUrl: window.location.href,
                pageContent: `ダイアログ表示中にエラーが発生しました: ${error.message}`,
                selectedText: '',
                currentService: currentService
            };
            createAiDialog(fallbackData);
        } catch (fallbackError) {
            console.error('フォールバックダイアログでもエラー:', fallbackError);
            alert('AI支援ツールでエラーが発生しました。ページを再読み込みしてください。');
        }
    }
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
    content.innerHTML = `        <div class="ai-dialog-header" style="
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 16px 20px;
            border-bottom: 1px solid ${borderColor};
            background: ${headerBg};
            color: white;
            border-radius: 12px 12px 0 0;
            position: sticky;
            top: 0;
            z-index: 10;
        ">
            <div style="display: flex; align-items: center; gap: 8px; flex: 1; min-width: 0;">
                <span style="font-size: 18px;">🏫</span>
                <div style="display: flex; align-items: center; gap: 8px; min-width: 0; flex: 1;">
                    <span style="
                        font-size: 14px; 
                        font-weight: 500; 
                        color: ${headerTextColor}; 
                        white-space: nowrap; 
                        overflow: hidden; 
                        text-overflow: ellipsis;
                        max-width: 300px;
                        cursor: pointer;
                    " title="${dialogData.pageTitle || 'タイトル不明'}\n${dialogData.pageUrl || ''}">${dialogData.pageTitle || 'AI支援ツール'}</span>
                    <button class="ai-copy-page-link-btn" style="
                        background: rgba(255, 255, 255, 0.1);
                        border: 1px solid rgba(255, 255, 255, 0.2);
                        color: ${headerTextColor};
                        padding: 4px 6px;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 11px;
                        transition: all 0.2s ease;
                        flex-shrink: 0;
                    " title="ページリンクをMarkdown形式でコピー">📋</button>
                </div>
            </div>
            <div style="display: flex; align-items: center; gap: 8px;">
                <button class="ai-header-settings-btn" style="
                    background: rgba(255, 255, 255, 0.1);
                    border: 1px solid rgba(255, 255, 255, 0.2);
                    color: ${headerTextColor};
                    padding: 6px 8px;
                    border-radius: 4px;
                    cursor: pointer;
                    font-size: 12px;
                    transition: all 0.2s ease;
                " title="設定">⚙️</button>
                <button class="ai-close-btn" style="
                    background: none;
                    border: none;
                    color: ${headerTextColor};
                    font-size: 20px;
                    cursor: pointer;
                    width: 28px;
                    height: 28px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    border-radius: 50%;
                ">×</button>
            </div>
        </div>
        <div style="padding: 20px; overflow-y: auto; flex: 1;">
            ${dialogData.selectedText ? `<div style="background: ${infoBg}; padding: 12px; border-radius: 6px; border-left: 4px solid #2196F3; margin-bottom: 16px; color: ${dialogText};"><strong>選択テキスト:</strong> ${dialogData.selectedText.substring(0, 100)}...</div>` : ''}
            
            <div style="display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px;">
                ${currentService === 'outlook' || currentService === 'gmail' ? `
                    <div style="display: flex; align-items: center; gap: 4px;">
                        <button class="ai-analyze-email-btn" style="
                            background: linear-gradient(135deg, #2196F3, #1976D2);
                            color: white;
                            border: none;
                            border-radius: 6px;
                            padding: 12px 16px;
                            cursor: pointer;
                            font-size: 14px;
                            font-weight: 500;
                            flex: 1;
                        ">📧 メール解析</button>
                        <button class="ai-copy-markdown-email-btn" style="
                            background: #f5f5f5;
                            border: 1px solid #ddd;
                            color: #666;
                            padding: 12px 8px;
                            border-radius: 6px;
                            cursor: pointer;
                            font-size: 11px;
                            flex-shrink: 0;
                        " title="ページ情報をMarkdown形式でコピー">📝</button>
                    </div>
                    <div style="display: flex; align-items: center; gap: 4px;">
                        <button class="ai-compose-reply-btn" style="
                            background: linear-gradient(135deg, #4CAF50, #388E3C);
                            color: white;
                            border: none;
                            border-radius: 6px;
                            padding: 12px 16px;
                            cursor: pointer;
                            font-size: 14px;
                            font-weight: 500;
                            flex: 1;
                        ">📝 返信作成</button>
                        <button class="ai-copy-markdown-reply-btn" style="
                            background: #f5f5f5;
                            border: 1px solid #ddd;
                            color: #666;
                            padding: 12px 8px;
                            border-radius: 6px;
                            cursor: pointer;
                            font-size: 11px;
                            flex-shrink: 0;
                        " title="ページ情報をMarkdown形式でコピー">📝</button>
                    </div>                ` : `
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
                `}                ${dialogData.selectedText ? `
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
                ` : ''}
                
                <!-- M365統合機能ボタン -->
                <button class="ai-forward-teams-btn" style="
                    background: linear-gradient(135deg, #6264A7, #464775);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 12px 16px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 500;
                ">💬 Teams chatに転送</button>
                
                <button class="ai-add-calendar-btn" style="
                    background: linear-gradient(135deg, #0078D4, #106EBE);
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 12px 16px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 500;
                ">📅 予定表に追加</button>
                
                ${dialogData.pageUrl && (dialogData.pageUrl.includes('code.visualstudio.com') || dialogData.pageUrl.includes('vscode') || dialogData.pageUrl.includes('marketplace.visualstudio.com')) ? `
                    <button class="ai-analyze-vscode-btn" style="
                        background: linear-gradient(135deg, #007ACC, #005A9E);
                        color: white;
                        border: none;
                        border-radius: 6px;
                        padding: 12px 16px;
                        cursor: pointer;
                        font-size: 14px;
                        font-weight: 500;
                    ">⚙️ VSCode設定解析</button>
                ` : ''}
            </div>
            
            <div>
                <div id="ai-loading" style="display: none; text-align: center; padding: 20px; color: ${textMuted};">
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
                <div id="ai-result" style="
                    display: none; 
                    background: ${prefersDark ? '#404040' : '#f9f9f9'}; 
                    padding: 16px; 
                    border-radius: 8px; 
                    border-left: 4px solid #2196F3; 
                    color: ${dialogText};
                    position: relative;
                ">
                    <button class="ai-copy-structured-result-btn" style="
                        position: absolute;
                        top: 8px;
                        right: 8px;
                        background: #2196F3;
                        border: none;
                        color: white;
                        padding: 6px 8px;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 11px;
                        display: none;
                    " title="構造的にコピー">📋</button>
                    <div id="ai-result-content"></div>
                </div>
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
    }    // 設定ボタン（旧版）
    const openSettingsBtn = dialog.querySelector('.ai-open-settings-btn');
    if (openSettingsBtn) {
        openSettingsBtn.addEventListener('click', openSettings);
        console.log('設定ボタンのイベントリスナー設定完了');
    }

    // 設定ボタン（ヘッダー版）
    const headerSettingsBtn = dialog.querySelector('.ai-header-settings-btn');
    if (headerSettingsBtn) {
        headerSettingsBtn.addEventListener('click', openSettings);
        console.log('ヘッダー設定ボタンのイベントリスナー設定完了');
    }

    // ページリンクコピーボタン
    const copyPageLinkBtn = dialog.querySelector('.ai-copy-page-link-btn');
    if (copyPageLinkBtn) {
        copyPageLinkBtn.addEventListener('click', () => copyPageLink(dialog));
        console.log('ページリンクコピーボタンのイベントリスナー設定完了');
    }    // Markdownコピーボタン群
    const markdownBtns = [
        '.ai-copy-markdown-email-btn',
        '.ai-copy-markdown-reply-btn'
    ];

    markdownBtns.forEach(selector => {
        const btn = dialog.querySelector(selector);
        if (btn) {
            btn.addEventListener('click', () => copyPageMarkdown(dialog));
            console.log(`Markdownコピーボタン (${selector}) のイベントリスナー設定完了`);
        }
    });

    // 構造的結果コピーボタン
    const copyStructuredBtn = dialog.querySelector('.ai-copy-structured-result-btn');
    if (copyStructuredBtn) {
        copyStructuredBtn.addEventListener('click', () => copyStructuredResult(dialog));
        console.log('構造的結果コピーボタンのイベントリスナー設定完了');
    }

    // Teams転送ボタン
    const forwardTeamsBtn = dialog.querySelector('.ai-forward-teams-btn');
    if (forwardTeamsBtn) {
        forwardTeamsBtn.addEventListener('click', () => forwardToTeams(dialog));
        console.log('Teams転送ボタンのイベントリスナー設定完了');
    }

    // 予定表追加ボタン
    const addCalendarBtn = dialog.querySelector('.ai-add-calendar-btn');
    if (addCalendarBtn) {
        addCalendarBtn.addEventListener('click', () => addToCalendar(dialog));
        console.log('予定表追加ボタンのイベントリスナー設定完了');
    }

    // VSCode設定解析ボタン
    const analyzeVSCodeBtn = dialog.querySelector('.ai-analyze-vscode-btn');
    if (analyzeVSCodeBtn) {
        analyzeVSCodeBtn.addEventListener('click', () => analyzeVSCodeSettings(dialog));
        console.log('VSCode設定解析ボタンのイベントリスナー設定完了');
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
function showNotification(message, type = 'success') {
    // 既存の通知があれば削除
    const existingNotification = document.querySelector('.ai-notification');
    if (existingNotification) {
        existingNotification.remove();
    }

    // 新しい通知を作成
    const notification = document.createElement('div');
    notification.className = `ai-notification ${type}`;
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed !important;
        top: 20px !important;
        right: 20px !important;
        background: ${type === 'error' ? '#f44336' : '#4caf50'} !important;
        color: white !important;
        padding: 12px 16px !important;
        border-radius: 6px !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        z-index: 2147483647 !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3) !important;
        animation: slideInRight 0.3s ease !important;
    `;

    // アニメーション用のCSSを追加
    if (!document.getElementById('ai-notification-style')) {
        const style = document.createElement('style');
        style.id = 'ai-notification-style';
        style.textContent = `
            @keyframes slideInRight {
                from {
                    transform: translateX(100%);
                    opacity: 0;
                }
                to {
                    transform: translateX(0);
                    opacity: 1;
                }
            }
        `;
        document.head.appendChild(style);
    }

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
/**
 * AI応答をそのまま表示（HTML形式）
 */
function showResult(result) {
    const resultElement = document.getElementById('ai-result');
    const resultContentElement = document.getElementById('ai-result-content');
    const copyStructuredBtn = document.querySelector('.ai-copy-structured-result-btn');

    if (resultElement && resultContentElement) {
        // エラーメッセージの場合はそのまま表示
        if (typeof result === 'string' && result.includes('❌ エラー:')) {
            resultContentElement.innerHTML = result;
        } else {
            // AI応答をHTML形式でそのまま表示
            // 基本的なサニタイズを実行（セキュリティ対策）
            const sanitizedResult = sanitizeHtmlResponse(result);
            resultContentElement.innerHTML = sanitizedResult;
        }

        resultElement.style.display = 'block';

        // 構造的コピーボタンを表示
        if (copyStructuredBtn) {
            copyStructuredBtn.style.display = 'block';
        }
    }
}

/**
 * HTML応答の基本的なサニタイズ（セキュリティ対策）
 */
function sanitizeHtmlResponse(html) {
    // 危険なタグやJavaScriptコードを除去
    const sanitized = html
        .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '') // <script>タグを除去
        .replace(/on\w+="[^"]*"/gi, '') // onclick等のイベントハンドラーを除去
        .replace(/javascript:/gi, '') // javascript:を除去
        .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '') // <iframe>を除去
        .replace(/<object\b[^<]*(?:(?!<\/object>)<[^<]*)*<\/object>/gi, '') // <object>を除去
        .replace(/<embed\b[^<]*(?:(?!<\/embed>)<[^<]*)*<\/embed>/gi, ''); // <embed>を除去

    return sanitized;
}

/**
 * AI応答を構造化されたHTMLにフォーマット
 */
function formatAIResponse(response, colors) {
    // エラーメッセージの場合はそのまま返す
    if (typeof response === 'string' && response.includes('❌ エラー:')) {
        return response;
    }

    // HTMLタグやCSSコードを除去・サニタイズ
    let content = sanitizeAIResponse(response);

    // マークダウン風の構造を検出して変換
    content = content
        // 見出し（## または # で始まる行）
        .replace(/^(#{1,3})\s*(.+)$/gm, (match, hashes, title) => {
            const level = hashes.length;
            const fontSize = level === 1 ? '18px' : level === 2 ? '16px' : '14px';
            const marginTop = level === 1 ? '20px' : '16px';
            const fontWeight = level === 1 ? '700' : '600';
            return `<h${level + 2} style="
                color: ${colors.headingColor}; 
                font-size: ${fontSize}; 
                margin: ${marginTop} 0 8px 0; 
                font-weight: ${fontWeight}; 
                border-bottom: ${level <= 2 ? `2px solid ${colors.borderColor}` : 'none'}; 
                padding-bottom: ${level <= 2 ? '6px' : '0'};
                letter-spacing: 0.5px;
            ">${title}</h${level + 2}>`;
        })        // 箇条書き（- または * で始まる行、ただしCSSではない）
        .replace(/^[-*]\s+(.+)$/gm, (match, content) => {
            // CSSプロパティ形式（例：margin: 6px 0;）ではないことを確認
            if (content.includes(':') && content.includes(';')) {
                return match; // CSSの可能性があるのでそのまま残す
            }
            return `<li style="
                margin: 6px 0; 
                line-height: 1.6; 
                padding-left: 8px;
                position: relative;
            ">
                <span style="
                    position: absolute; 
                    left: -16px; 
                    color: ${colors.headingColor};
                    font-weight: bold;
                ">•</span>
                ${content}
            </li>`;
        })

        // 番号付きリスト（1. で始まる行、ただしCSSではない）
        .replace(/^(\d+)\.\s+(.+)$/gm, (match, number, content) => {
            // CSSプロパティ形式ではないことを確認
            if (content.includes(':') && content.includes(';')) {
                return match; // CSSの可能性があるのでそのまま残す
            }
            return `<li style="
                margin: 6px 0; 
                line-height: 1.6; 
                counter-increment: list-counter;
            " data-number="${number}">${content}</li>`;
        })

        // 太字（**text** または __text__）
        .replace(/\*\*(.*?)\*\*/g, `<strong style="
            color: ${colors.headingColor}; 
            font-weight: 600;
            background: ${colors.borderColor}20;
            padding: 2px 4px;
            border-radius: 3px;
        ">$1</strong>`)
        .replace(/__(.*?)__/g, `<strong style="
            color: ${colors.headingColor}; 
            font-weight: 600;
            background: ${colors.borderColor}20;
            padding: 2px 4px;
            border-radius: 3px;
        ">$1</strong>`)

        // コードブロック（```で囲まれた部分）
        .replace(/```([\s\S]*?)```/g, `<pre style="
            background: ${colors.borderColor}30;
            border: 1px solid ${colors.borderColor};
            border-radius: 4px;
            padding: 12px;
            margin: 8px 0;
            overflow-x: auto;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 13px;
            line-height: 1.4;
        "><code>$1</code></pre>`)

        // インラインコード（`code`）
        .replace(/`([^`]+)`/g, `<code style="
            background: ${colors.borderColor}30;
            border: 1px solid ${colors.borderColor};
            border-radius: 3px;
            padding: 2px 4px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            color: ${colors.headingColor};
        ">$1</code>`)

        // 改行を段落に変換
        .split('\n\n')
        .filter(para => para.trim())
        .map((para, index) => {
            // リストアイテムが含まれている場合
            if (para.includes('<li')) {
                const listItems = para.split('\n').filter(line => line.includes('<li'));
                const otherContent = para.split('\n').filter(line => !line.includes('<li') && line.trim());

                let result = '';
                if (otherContent.length > 0) {
                    result += `<p style="
                        margin: 12px 0; 
                        line-height: 1.6; 
                        color: ${colors.textColor};
                        text-align: justify;
                    ">${otherContent.join('<br>')}</p>`;
                }

                // 番号付きリストか通常のリストかを判定
                const isNumberedList = listItems.some(item => item.includes('data-number'));
                const listTag = isNumberedList ? 'ol' : 'ul';
                const listStyle = isNumberedList ?
                    `counter-reset: list-counter; list-style: none; padding-left: 24px;` :
                    `list-style: none; padding-left: 24px;`;

                result += `<${listTag} style="
                    margin: 12px 0; 
                    ${listStyle}
                    color: ${colors.textColor};
                ">${listItems.join('')}</${listTag}>`;
                return result;
            } else if (para.includes('<pre>')) {
                // コードブロックの場合はそのまま
                return para;
            } else {
                // 通常の段落
                return `<p style="
                    margin: 12px 0; 
                    line-height: 1.6; 
                    color: ${colors.textColor};
                    text-align: justify;
                    text-indent: ${index > 0 ? '1em' : '0'};
                ">${para.replace(/\n/g, '<br>')}</p>`;
            }
        })
        .join('');

    // ユニークIDを生成
    const containerId = `ai-result-container-${Date.now()}`;

    // 全体を囲むコンテナ
    const containerHTML = `
        <div id="${containerId}" style="
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            color: ${colors.textColor};
            background: ${colors.bgColor};
            padding: 20px;
            border-radius: 12px;
            border: 1px solid ${colors.borderColor};
            max-height: 500px;
            overflow-y: auto;
            box-shadow: 0 4px 12px ${colors.borderColor}40;
        ">
            <div style="
                margin-bottom: 16px;
                padding-bottom: 12px;
                border-bottom: 2px solid ${colors.borderColor};
                display: flex;
                align-items: center;
                gap: 8px;
            ">
                <span style="
                    font-size: 18px;
                    color: ${colors.headingColor};
                ">🤖</span>
                <h3 style="
                    margin: 0;
                    font-size: 16px;
                    font-weight: 600;
                    color: ${colors.headingColor};
                ">AI 解析結果</h3>
            </div>
            
            <div class="ai-content" style="margin-bottom: 16px;">
                ${content}
            </div>
            
            <div style="
                margin-top: 20px;
                padding-top: 16px;
                border-top: 1px solid ${colors.borderColor};
                display: flex;
                gap: 10px;
                flex-wrap: wrap;
                justify-content: flex-end;
            ">
                <button class="copy-btn" data-container="${containerId}" style="
                    background: #2196F3;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 13px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: all 0.2s ease;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                    box-shadow: 0 2px 4px rgba(33, 150, 243, 0.3);
                ">
                    <span>📋</span>
                    <span>結果をコピー</span>
                </button>
                <button class="save-btn" data-container="${containerId}" style="
                    background: #4CAF50;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 13px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: all 0.2s ease;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                    box-shadow: 0 2px 4px rgba(76, 175, 80, 0.3);
                ">
                    <span>💾</span>
                    <span>履歴に保存</span>
                </button>
                <button class="expand-btn" data-container="${containerId}" style="
                    background: #FF9800;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 6px;
                    font-size: 13px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: all 0.2s ease;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                    box-shadow: 0 2px 4px rgba(255, 152, 0, 0.3);
                ">
                    <span>�</span>
                    <span>拡大表示</span>
                </button>
            </div>
        </div>
    `;

    // イベントリスナーを後で追加するために、イベントハンドラーを設定
    setTimeout(() => {
        setupResultActionButtons(containerId, content, colors);
    }, 100);

    return containerHTML;
}

/**
 * 結果表示のアクションボタンのイベントハンドラーを設定
 */
function setupResultActionButtons(containerId, content, colors) {
    const container = document.getElementById(containerId);
    if (!container) return;

    // コピーボタン
    const copyBtn = container.querySelector('.copy-btn');
    if (copyBtn) {
        copyBtn.addEventListener('click', () => {
            copyResultToClipboard(container, content);
        });

        // ホバー効果
        copyBtn.addEventListener('mouseenter', () => {
            copyBtn.style.background = '#1976D2';
            copyBtn.style.transform = 'translateY(-1px)';
            copyBtn.style.boxShadow = '0 4px 8px rgba(33, 150, 243, 0.4)';
        });

        copyBtn.addEventListener('mouseleave', () => {
            copyBtn.style.background = '#2196F3';
            copyBtn.style.transform = 'translateY(0)';
            copyBtn.style.boxShadow = '0 2px 4px rgba(33, 150, 243, 0.3)';
        });
    }

    // 保存ボタン
    const saveBtn = container.querySelector('.save-btn');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            saveResultToHistory(content);
        });

        // ホバー効果
        saveBtn.addEventListener('mouseenter', () => {
            saveBtn.style.background = '#388E3C';
            saveBtn.style.transform = 'translateY(-1px)';
            saveBtn.style.boxShadow = '0 4px 8px rgba(76, 175, 80, 0.4)';
        });

        saveBtn.addEventListener('mouseleave', () => {
            saveBtn.style.background = '#4CAF50';
            saveBtn.style.transform = 'translateY(0)';
            saveBtn.style.boxShadow = '0 2px 4px rgba(76, 175, 80, 0.3)';
        });
    }

    // 拡大表示ボタン
    const expandBtn = container.querySelector('.expand-btn');
    if (expandBtn) {
        expandBtn.addEventListener('click', () => {
            expandResultView(content, colors);
        });

        // ホバー効果
        expandBtn.addEventListener('mouseenter', () => {
            expandBtn.style.background = '#F57C00';
            expandBtn.style.transform = 'translateY(-1px)';
            expandBtn.style.boxShadow = '0 4px 8px rgba(255, 152, 0, 0.4)';
        });

        expandBtn.addEventListener('mouseleave', () => {
            expandBtn.style.background = '#FF9800';
            expandBtn.style.transform = 'translateY(0)';
            expandBtn.style.boxShadow = '0 2px 4px rgba(255, 152, 0, 0.3)';
        });
    }
}

/**
 * 結果をクリップボードにコピー
 */
function copyResultToClipboard(container, content) {
    try {        // HTMLタグを削除してプレーンテキストに変換
        const textContent = content
            .replace(/<[^>]*>/g, '') // HTMLタグを削除
            .replace(/&nbsp;/g, ' ') // 非改行スペースを通常のスペースに
            .replace(/&lt;/g, '<')   // HTML エンティティをデコード
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/\s+/g, ' ')    // 複数の空白を単一のスペースに
            .trim();

        navigator.clipboard.writeText(textContent).then(() => {
            showNotification('結果をクリップボードにコピーしました', 'success');

            // ボタンのテキストを一時的に変更
            const copyBtn = container.querySelector('.copy-btn span:last-child');
            if (copyBtn) {
                const originalText = copyBtn.textContent;
                copyBtn.textContent = 'コピー完了!';
                setTimeout(() => {
                    copyBtn.textContent = originalText;
                }, 2000);
            }
        }).catch(err => {
            console.error('クリップボードコピーエラー:', err);
            showNotification('クリップボードコピーに失敗しました', 'error');
        });
    } catch (error) {
        console.error('クリップボードコピーエラー:', error);
        showNotification('クリップボードコピーに失敗しました', 'error');
    }
}

/**
 * 結果を履歴に保存
 */
function saveResultToHistory(content) {
    try {
        const timestamp = new Date().toISOString();
        const historyItem = {
            timestamp,
            content,
            url: window.location.href,
            title: document.title
        };

        // 履歴を取得（ローカルストレージから）
        let history = [];
        try {
            const savedHistory = localStorage.getItem('ptaAiAnalysisHistory');
            if (savedHistory) {
                history = JSON.parse(savedHistory);
            }
        } catch (e) {
            console.warn('履歴の読み込みに失敗:', e);
        }

        // 新しいアイテムを追加（最新を先頭に）
        history.unshift(historyItem);

        // 履歴の上限を設定（最大50件）
        if (history.length > 50) {
            history = history.slice(0, 50);
        }

        // 履歴を保存
        localStorage.setItem('ptaAiAnalysisHistory', JSON.stringify(history));

        showNotification('結果を履歴に保存しました', 'success');

    } catch (error) {
        console.error('履歴保存エラー:', error);
        showNotification('履歴保存に失敗しました', 'error');
    }
}

/**
 * 結果を拡大表示
 */
function expandResultView(content, colors) {
    // 拡大表示用のモーダルダイアログを作成
    const expandModal = document.createElement('div');
    expandModal.id = 'ai-expand-modal';
    expandModal.className = 'ai-modal-overlay';

    expandModal.innerHTML = `
        <div class="ai-expand-dialog" style="
            background: ${colors.bgColor};
            border: 2px solid ${colors.borderColor};
            border-radius: 12px;
            width: 90%;
            max-width: 900px;
            max-height: 90%;
            overflow: hidden;
            display: flex;
            flex-direction: column;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
        ">
            <div style="
                display: flex;
                justify-content: between;
                align-items: center;
                padding: 16px 24px;
                border-bottom: 2px solid ${colors.borderColor};
                background: ${colors.headingColor}10;
            ">
                <h2 style="
                    margin: 0;
                    color: ${colors.headingColor};
                    font-size: 18px;
                    font-weight: 600;
                    display: flex;
                    align-items: center;
                    gap: 10px;
                ">
                    <span>🔍</span>
                    <span>AI解析結果 - 拡大表示</span>
                </h2>
                <button class="expand-close-btn" style="
                    background: none;
                    border: none;
                    font-size: 24px;
                    cursor: pointer;
                    color: ${colors.textColor};
                    padding: 4px 8px;
                    border-radius: 4px;
                    transition: background 0.2s;
                " title="閉じる">
                    ×
                </button>
            </div>
            
            <div style="
                flex: 1;
                overflow-y: auto;
                padding: 24px;
                font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
                font-size: 15px;
                line-height: 1.7;
                color: ${colors.textColor};
            ">
                ${content}
            </div>
            
            <div style="
                padding: 16px 24px;
                border-top: 1px solid ${colors.borderColor};
                background: ${colors.headingColor}05;
                display: flex;
                gap: 12px;
                justify-content: flex-end;
            ">
                <button class="expand-copy-btn" style="
                    background: #2196F3;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 6px;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: all 0.2s ease;
                ">
                    📋 コピー
                </button>
                <button class="expand-close-action-btn" style="
                    background: #757575;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    border-radius: 6px;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    transition: all 0.2s ease;
                ">
                    閉じる
                </button>
            </div>
        </div>
    `;

    document.body.appendChild(expandModal);

    // 拡大表示モーダルのイベントハンドラーを設定
    setupExpandModalHandlers(expandModal, content, colors);

    // フォーカスを拡大表示モーダルに移動
    const dialog = expandModal.querySelector('.ai-expand-dialog');
    if (dialog) {
        dialog.focus();
    }
}

/**
 * 拡大表示モーダルのイベントハンドラーを設定
 */
function setupExpandModalHandlers(modal, content, colors) {
    // 閉じるボタンのイベント
    const closeBtn = modal.querySelector('.expand-close-btn');
    const closeActionBtn = modal.querySelector('.expand-close-action-btn');

    const closeModal = () => {
        modal.remove();
    };

    if (closeBtn) {
        closeBtn.addEventListener('click', closeModal);

        // ホバー効果
        closeBtn.addEventListener('mouseenter', () => {
            closeBtn.style.background = colors.borderColor + '40';
        });
        closeBtn.addEventListener('mouseleave', () => {
            closeBtn.style.background = 'none';
        });
    }

    if (closeActionBtn) {
        closeActionBtn.addEventListener('click', closeModal);

        // ホバー効果
        closeActionBtn.addEventListener('mouseenter', () => {
            closeActionBtn.style.background = '#616161';
        });
        closeActionBtn.addEventListener('mouseleave', () => {
            closeActionBtn.style.background = '#757575';
        });
    }

    // コピーボタンのイベント
    const copyBtn = modal.querySelector('.expand-copy-btn');
    if (copyBtn) {
        copyBtn.addEventListener('click', () => {
            copyResultToClipboard(modal, content);
        });

        // ホバー効果
        copyBtn.addEventListener('mouseenter', () => {
            copyBtn.style.background = '#1976D2';
        });
        copyBtn.addEventListener('mouseleave', () => {
            copyBtn.style.background = '#2196F3';
        });
    }

    // ESCキーで閉じる
    const escHandler = (event) => {
        if (event.key === 'Escape') {
            closeModal();
            document.removeEventListener('keydown', escHandler);
        }
    };
    document.addEventListener('keydown', escHandler);

    // オーバーレイクリックで閉じる
    modal.addEventListener('click', (event) => {
        if (event.target === modal) {
            closeModal();
        }
    });
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
        // 基本的な要素の存在チェック
        if (!document || !document.body) {
            console.warn('Document または body が存在しません');
            return 'ページが正常に読み込まれていません';
        }

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
            try {
                const element = document.querySelector(selector);
                if (element && element.innerText && element.innerText.trim().length > 50) {
                    pageContent = element.innerText.trim();
                    console.log(`コンテンツ抽出成功: ${selector}`);
                    break;
                }
            } catch (selectorError) {
                console.warn(`セレクタエラー: ${selector}`, selectorError);
                continue;
            }
        }

        // 方法2: body全体から抽出（スクリプトやスタイルを除外）
        if (!pageContent.trim()) {
            try {
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
                        elements.forEach(el => {
                            if (el && el.parentNode) {
                                el.parentNode.removeChild(el);
                            }
                        });
                    } catch (removeError) {
                        console.warn(`要素削除エラー: ${selector}`, removeError);
                    }
                });

                pageContent = bodyClone.innerText || bodyClone.textContent || '';
                console.log('body全体からコンテンツ抽出');
            } catch (cloneError) {
                console.warn('body cloneでエラー:', cloneError);
            }
        }

        // 方法3: 空の場合のフォールバック
        if (!pageContent.trim()) {
            try {
                pageContent = document.body.innerText || document.body.textContent || '';
                console.log('フォールバックでコンテンツ抽出');
            } catch (fallbackError) {
                console.warn('フォールバック抽出でエラー:', fallbackError);
            }
        }

        // 長すぎる場合は切り詰め（最初の8000文字）
        if (pageContent && pageContent.length > 8000) {
            pageContent = pageContent.substring(0, 8000).trim();
            console.log('コンテンツを8000文字に切り詰め');
        }

        // 最終チェック
        if (!pageContent || !pageContent.trim()) {
            pageContent = `（ページコンテンツを取得できませんでした）\nURL: ${window.location.href}\nタイトル: ${document.title || '不明'}`;
            console.warn('ページコンテンツが空です');
        }

    } catch (error) {
        console.error('ページコンテンツ抽出エラー:', error);
        console.error('エラースタック:', error.stack);

        // エラーの場合は最低限の情報を返す
        try {
            const fallbackContent = document.body?.innerText || document.body?.textContent;
            if (fallbackContent && fallbackContent.trim()) {
                pageContent = fallbackContent.trim();
            } else {
                pageContent = `コンテンツ取得中にエラーが発生しました\nURL: ${window.location.href}\nタイトル: ${document.title || '不明'}\nエラー: ${error.message}`;
            }
        } catch (finalError) {
            console.error('最終フォールバックでもエラー:', finalError);
            pageContent = `重大なエラー: コンテンツを取得できませんでした (${error.message})`;
        }
    }

    console.log(`最終的に抽出されたコンテンツ（最初の200文字）:`, pageContent.substring(0, 200));
    return pageContent || 'コンテンツが空です';
}

/**
 * AIの応答からCSSやHTMLコードを除去・サニタイズする
 * @param {string} response - AIの生の応答
 * @returns {string} - サニタイズされた応答
 */
function sanitizeAIResponse(response) {
    if (!response || typeof response !== 'string') {
        return response;
    }

    let sanitized = response;

    // CSSプロパティのパターンを除去（例：margin: 6px 0; や line-height: 1.6; など）
    sanitized = sanitized.replace(/[a-zA-Z-]+\s*:\s*[^;]+;/g, '');

    // CSS値のパターンを除去（例：counter-increment: list-counter;）
    sanitized = sanitized.replace(/counter-increment:\s*[^;]+;/g, '');

    // data-*属性を除去（例：data-number="1"）
    sanitized = sanitized.replace(/data-[a-zA-Z-]+\s*=\s*"[^"]*"/g, '');

    // CSS値の単体パターンを除去（例：margin: 6px 0;）
    sanitized = sanitized.replace(/\b(margin|padding|line-height|font-size|font-weight|color|background|border|display|position|width|height|top|left|right|bottom|float|clear|text-align|vertical-align|z-index|opacity|transform|transition|animation|box-shadow|border-radius|overflow|cursor|text-decoration|text-transform|letter-spacing|word-spacing|white-space|font-family|list-style|counter-increment|counter-reset)\s*:\s*[^;]+;?/gi, '');

    // CSSのプロパティ値だけが残ってしまった行を除去
    sanitized = sanitized.replace(/^\s*[0-9.]+px\s*$/gm, '');
    sanitized = sanitized.replace(/^\s*[0-9.]+\s*$/gm, '');
    sanitized = sanitized.replace(/^\s*(left|right|center|bold|normal|none|auto|inherit|initial|unset)\s*$/gm, '');

    // styleタグとその内容を除去
    sanitized = sanitized.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

    // scriptタグとその内容を除去
    sanitized = sanitized.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');

    // HTMLタグを除去（ただし、改行は保持）
    sanitized = sanitized.replace(/<[^>]+>/g, '');

    // CSSのコメントを除去
    sanitized = sanitized.replace(/\/\*[\s\S]*?\*\//g, '');

    // 単独で残ったCSS記号や値を除去
    sanitized = sanitized.replace(/^[\s]*[{}();,]+[\s]*$/gm, '');

    // 複数の連続する空白行を1つにまとめる
    sanitized = sanitized.replace(/\n\s*\n\s*\n/g, '\n\n');

    // 行頭の余分な空白を除去
    sanitized = sanitized.replace(/^\s+/gm, '');

    // 文字列の前後の空白を除去
    sanitized = sanitized.trim();
    console.log('🧼 AIレスポンスサニタイズ完了:', {
        originalLength: response.length,
        sanitizedLength: sanitized.length,
        originalPreview: response.substring(0, 300) + '...',
        sanitizedPreview: sanitized.substring(0, 300) + '...',
        removedCSSCount: (response.match(/[a-zA-Z-]+\s*:\s*[^;]+;/g) || []).length
    });

    return sanitized;
}

/**
 * 要素をドラッグ可能にする
 * @param {HTMLElement} element - ドラッグ可能にする要素
 * @param {HTMLElement} handle - ドラッグハンドル要素
 */
function makeDraggable(element, handle) {
    let isDragging = false;
    let startX, startY, startRight, startTop;

    // ハンドル要素にマウスダウンイベントを追加
    handle.addEventListener('mousedown', (e) => {
        // 左クリックのみ対応
        if (e.button !== 0) return;

        isDragging = true;
        element.isDragging = true;

        // ドラッグ開始時の位置を記録
        startX = e.clientX;
        startY = e.clientY;

        // 現在の位置を取得（right基準で保存しているため）
        const rect = element.getBoundingClientRect();
        startRight = window.innerWidth - rect.right;
        startTop = rect.top;

        // ドラッグ中のスタイル変更
        element.classList.add('dragging');
        handle.style.cursor = 'grabbing';
        element.style.opacity = '0.8';
        element.style.transform = 'scale(1.05)';
        element.style.transition = 'none';

        // ハンドルのホバー効果を強調
        handle.style.opacity = '1';
        handle.style.backgroundColor = 'rgba(255, 255, 255, 0.2)';

        // ドキュメント全体でマウスイベントを監視
        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);

        // テキスト選択を防止
        e.preventDefault();        // テキスト選択を防止
        e.preventDefault();
        e.stopPropagation();
    });

    function onMouseMove(e) {
        if (!isDragging) return;

        // マウスの移動量を計算
        const deltaX = e.clientX - startX;
        const deltaY = e.clientY - startY;

        // 新しい位置を計算（画面の境界を考慮）
        const newTop = Math.max(0, Math.min(window.innerHeight - element.offsetHeight, startTop + deltaY));
        const newRight = Math.max(0, Math.min(window.innerWidth - element.offsetWidth, startRight - deltaX));

        // 位置を更新
        element.style.top = newTop + 'px';
        element.style.right = newRight + 'px';
    } function onMouseUp() {
        if (!isDragging) return;

        isDragging = false;

        // ドラッグ終了後のスタイル復元
        element.classList.remove('dragging');
        handle.style.cursor = 'move';
        element.style.opacity = '1';
        element.style.transform = 'scale(1)';
        element.style.transition = 'all 0.3s ease';

        // ハンドルのスタイルを元に戻す
        handle.style.opacity = '0.7';
        handle.style.backgroundColor = '';

        // 現在の位置を保存
        const rect = element.getBoundingClientRect();
        const position = {
            top: rect.top,
            right: window.innerWidth - rect.right
        };
        saveButtonPosition(position);

        // イベントリスナーを削除
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);

        // 少し遅れてドラッグフラグをクリア（クリックイベントとの競合回避）
        setTimeout(() => {
            element.isDragging = false;
        }, 100);
    }
}

/**
 * ボタンの位置を保存
 * @param {Object} position - 位置情報 {top, right}
 */
function saveButtonPosition(position) {
    try {
        chrome.storage.local.set({
            'aiButtonPosition': position
        }, () => {
            console.log('ボタン位置を保存しました:', position);
        });
    } catch (error) {
        console.error('ボタン位置の保存に失敗:', error);
        // フォールバック: localStorageを使用
        try {
            localStorage.setItem('aiButtonPosition', JSON.stringify(position));
        } catch (e) {
            console.error('localStorageでの保存も失敗:', e);
        }
    }
}

/**
 * 保存されたボタンの位置を取得
 * @returns {Object} 位置情報 {top, right}
 */
function getSavedButtonPosition() {
    const defaultPosition = { top: 20, right: 20 };

    return new Promise((resolve) => {
        try {
            chrome.storage.local.get(['aiButtonPosition'], (result) => {
                if (result.aiButtonPosition) {
                    // 画面サイズが変わった場合の調整
                    const position = result.aiButtonPosition;
                    position.top = Math.max(0, Math.min(window.innerHeight - 60, position.top));
                    position.right = Math.max(0, Math.min(window.innerWidth - 100, position.right));
                    resolve(position);
                } else {
                    resolve(defaultPosition);
                }
            });
        } catch (error) {
            console.error('ボタン位置の取得に失敗:', error);
            // フォールバック: localStorageから取得
            try {
                const saved = localStorage.getItem('aiButtonPosition');
                if (saved) {
                    const position = JSON.parse(saved);
                    resolve(position);
                } else {
                    resolve(defaultPosition);
                }
            } catch (e) {
                console.error('localStorageからの取得も失敗:', e);
                resolve(defaultPosition);
            }
        }
    });
}

/**
 * 保存されたボタンの位置を同期的に取得（初期化用）
 * @returns {Object} 位置情報 {top, right}
 */
function getSavedButtonPositionSync() {
    const defaultPosition = { top: 20, right: 20 };

    try {
        // まずlocalStorageから取得を試行
        const saved = localStorage.getItem('aiButtonPosition');
        if (saved) {
            const position = JSON.parse(saved);
            // 画面サイズが変わった場合の調整
            position.top = Math.max(0, Math.min(window.innerHeight - 60, position.top));
            position.right = Math.max(0, Math.min(window.innerWidth - 100, position.right));
            return position;
        }
    } catch (error) {
        console.error('同期的な位置取得に失敗:', error);
    }

    return defaultPosition;
}

/**
 * ページリンクをコピー（Markdown形式）
 */
async function copyPageLink(dialog) {
    try {
        const dialogData = dialog.dialogData;
        const markdownLink = `[${dialogData.pageTitle || 'タイトル不明'}](${dialogData.pageUrl || ''})`;

        await navigator.clipboard.writeText(markdownLink);
        showNotification('ページリンクをMarkdown形式でコピーしました', 'success');
    } catch (error) {
        console.error('ページリンクコピーエラー:', error);
        showNotification('コピーに失敗しました', 'error');
    }
}

/**
 * ページ情報をMarkdown形式でコピー
 */
async function copyPageMarkdown(dialog) {
    try {
        const dialogData = dialog.dialogData;
        const markdownText = `[${dialogData.pageTitle || 'タイトル不明'}](${dialogData.pageUrl || ''})`;

        await navigator.clipboard.writeText(markdownText);
        showNotification('Markdown形式でコピーしました', 'success');
    } catch (error) {
        console.error('Markdownコピーエラー:', error);
        showNotification('コピーに失敗しました', 'error');
    }
}

/**
 * 構造的な結果をコピー
 */
async function copyStructuredResult(dialog) {
    try {
        const dialogData = dialog.dialogData;
        const resultContent = document.getElementById('ai-result-content');
        const resultText = resultContent ? (resultContent.textContent || resultContent.innerText) : '';

        const structuredResult = `# AI解析結果\n\n## 対象ページ\n- タイトル: ${dialogData.pageTitle || 'タイトル不明'}\n- URL: ${dialogData.pageUrl || ''}\n\n## 解析結果\n${resultText}\n\n---\n生成日時: ${new Date().toLocaleString('ja-JP')}`;

        await navigator.clipboard.writeText(structuredResult);
        showNotification('構造的な結果をコピーしました', 'success');
    } catch (error) {
        console.error('構造的コピーエラー:', error);
        showNotification('コピーに失敗しました', 'error');
    }
}

/**
 * Teams chatへの転送処理
 */
async function forwardToTeams(dialog) {
    try {
        const dialogData = dialog.dialogData;
        
        // ローディング表示
        showLoading();
        
        // バックグラウンドスクリプトに転送リクエストを送信
        const response = await chrome.runtime.sendMessage({
            action: 'forwardToTeams',
            data: {
                pageTitle: dialogData.pageTitle,
                pageUrl: dialogData.pageUrl,
                content: dialogData.pageContent || dialogData.selectedText || ''
            }
        });

        hideLoading();
        
        if (response.success) {
            showResult(`<div style="color: #4CAF50; padding: 16px; background: #f1f8e9; border-radius: 8px; border-left: 4px solid #4CAF50;">
                <h3>✅ Teams転送完了</h3>
                <p>${response.message}</p>
                ${response.method === 'web' ? '<p><small>💡 Teams Web版が開きます。チャット画面で内容を確認して送信してください。</small></p>' : ''}
            </div>`);
        } else {
            showResult(`<div style="color: #f44336; padding: 16px; background: #ffebee; border-radius: 8px; border-left: 4px solid #f44336;">
                <h3>❌ Teams転送エラー</h3>
                <p>${response.error}</p>
                <p><small>💡 Microsoft 365へのログインとTeamsへのアクセス権限が必要です。</small></p>
            </div>`);
        }
        
    } catch (error) {
        hideLoading();
        console.error('Teams転送エラー:', error);
        showResult(`<div style="color: #f44336; padding: 16px; background: #ffebee; border-radius: 8px; border-left: 4px solid #f44336;">
            <h3>❌ 転送処理エラー</h3>
            <p>Teams転送中にエラーが発生しました: ${error.message}</p>
        </div>`);
    }
}

/**
 * 予定表への追加処理
 */
async function addToCalendar(dialog) {
    try {
        const dialogData = dialog.dialogData;
        
        // ローディング表示
        showLoading();
        
        // バックグラウンドスクリプトに予定表追加リクエストを送信
        const response = await chrome.runtime.sendMessage({
            action: 'addToCalendar',
            data: {
                pageTitle: dialogData.pageTitle,
                pageUrl: dialogData.pageUrl,
                content: dialogData.pageContent || dialogData.selectedText || ''
            }
        });

        hideLoading();
        
        if (response.success) {
            const eventInfo = response.event;
            showResult(`<div style="color: #2196F3; padding: 16px; background: #e3f2fd; border-radius: 8px; border-left: 4px solid #2196F3;">
                <h3>📅 予定表追加完了</h3>
                <p>${response.message}</p>
                ${eventInfo ? `
                    <div style="margin-top: 12px; padding: 8px; background: rgba(255,255,255,0.5); border-radius: 4px;">
                        <p><strong>件名:</strong> ${eventInfo.subject}</p>
                        <p><strong>開始時刻:</strong> ${new Date(eventInfo.startTime).toLocaleString('ja-JP')}</p>
                    </div>
                ` : ''}
                ${response.method === 'web' ? '<p><small>💡 Outlook Web版が開きます。予定の詳細を確認して保存してください。</small></p>' : ''}
            </div>`);
        } else {
            showResult(`<div style="color: #f44336; padding: 16px; background: #ffebee; border-radius: 8px; border-left: 4px solid #f44336;">
                <h3>❌ 予定表追加エラー</h3>
                <p>${response.error}</p>
                <p><small>💡 Microsoft 365へのログインとOutlookへのアクセス権限が必要です。</small></p>
            </div>`);
        }
        
    } catch (error) {
        hideLoading();
        console.error('予定表追加エラー:', error);
        showResult(`<div style="color: #f44336; padding: 16px; background: #ffebee; border-radius: 8px; border-left: 4px solid #f44336;">
            <h3>❌ 予定表処理エラー</h3>
            <p>予定表追加中にエラーが発生しました: ${error.message}</p>
        </div>`);
    }
}

/**
 * VSCode設定解析処理
 */
async function analyzeVSCodeSettings(dialog) {
    try {
        const dialogData = dialog.dialogData;
        
        // VSCodeドキュメントページかチェック
        const isVSCodeDoc = dialogData.pageUrl && (
            dialogData.pageUrl.includes('code.visualstudio.com') ||
            dialogData.pageUrl.includes('vscode') ||
            dialogData.pageUrl.includes('marketplace.visualstudio.com')
        );
        
        if (!isVSCodeDoc) {
            showResult(`<div style="color: #ff9800; padding: 16px; background: #fff3e0; border-radius: 8px; border-left: 4px solid #ff9800;">
                <h3>⚠️ VSCodeドキュメントページではありません</h3>
                <p>この機能はVSCode関連のドキュメントページでのみ利用できます。</p>
                <p>対象サイト: code.visualstudio.com, marketplace.visualstudio.com など</p>
            </div>`);
            return;
        }
        
        // ローディング表示
        showLoading();
        
        // バックグラウンドスクリプトに解析リクエストを送信
        const response = await chrome.runtime.sendMessage({
            action: 'analyzeVSCodeSettings',
            data: {
                pageTitle: dialogData.pageTitle,
                pageUrl: dialogData.pageUrl,
                content: dialogData.pageContent || ''
            }
        });

        hideLoading();
        
        if (response.success) {
            // 解析結果を表示（コピーボタン付き）
            const resultHtml = `<div style="padding: 16px; background: #f0f8ff; border-radius: 8px; border-left: 4px solid #007ACC;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                    <h3 style="margin: 0; color: #007ACC;">⚙️ VSCode設定解析結果</h3>
                    <button onclick="copyVSCodeAnalysis()" style="
                        background: #007ACC;
                        color: white;
                        border: none;
                        padding: 6px 12px;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 12px;
                    ">📋 全体をコピー</button>
                </div>
                <div id="vscode-analysis-content">${response.analysis}</div>
                <div style="margin-top: 12px; padding: 8px; background: rgba(255,255,255,0.5); border-radius: 4px; font-size: 12px; color: #666;">
                    <strong>対象ページ:</strong> <a href="${response.pageInfo.url}" target="_blank">${response.pageInfo.title}</a>
                </div>
            </div>`;
            
            showResult(resultHtml);
            
            // コピー機能をグローバルに追加
            window.copyVSCodeAnalysis = async () => {
                try {
                    const content = document.getElementById('vscode-analysis-content');
                    if (content) {
                        const textContent = content.innerText || content.textContent;
                        await navigator.clipboard.writeText(textContent);
                        showNotification('VSCode設定解析結果をコピーしました', 'success');
                    }
                } catch (error) {
                    console.error('コピーエラー:', error);
                    showNotification('コピーに失敗しました', 'error');
                }
            };
            
        } else {
            showResult(`<div style="color: #f44336; padding: 16px; background: #ffebee; border-radius: 8px; border-left: 4px solid #f44336;">
                <h3>❌ VSCode設定解析エラー</h3>
                <p>${response.error}</p>
                ${response.suggestion ? `<p><small>💡 ${response.suggestion}</small></p>` : ''}
            </div>`);
        }
        
    } catch (error) {
        hideLoading();
        console.error('VSCode設定解析エラー:', error);
        showResult(`<div style="color: #f44336; padding: 16px; background: #ffebee; border-radius: 8px; border-left: 4px solid #f44336;">
            <h3>❌ 解析処理エラー</h3>
            <p>VSCode設定解析中にエラーが発生しました: ${error.message}</p>
        </div>`);
    }
}

/**
 * コンテキストメニューからのTeams転送ハンドラー
 */
function handleForwardToTeamsFromContext(data) {
    const pageData = {
        pageTitle: data.pageTitle,
        pageUrl: data.pageUrl,
        pageContent: document.body.textContent || document.body.innerText || '',
        currentService: currentService
    };

    createAiDialog(pageData);
    
    // ダイアログが作成された後にTeams転送を実行
    setTimeout(() => {
        const dialog = document.getElementById('ai-dialog');
        if (dialog) {
            forwardToTeams(dialog);
        }
    }, 100);
}

/**
 * コンテキストメニューからの予定表追加ハンドラー
 */
function handleAddToCalendarFromContext(data) {
    const pageData = {
        pageTitle: data.pageTitle,
        pageUrl: data.pageUrl,
        pageContent: document.body.textContent || document.body.innerText || '',
        currentService: currentService
    };

    createAiDialog(pageData);
    
    // ダイアログが作成された後に予定表追加を実行
    setTimeout(() => {
        const dialog = document.getElementById('ai-dialog');
        if (dialog) {
            addToCalendar(dialog);
        }
    }, 100);
}

/**
 * コンテキストメニューからのVSCode設定解析ハンドラー
 */
function handleAnalyzeVSCodeSettingsFromContext(data) {
    const pageData = {
        pageTitle: data.pageTitle,
        pageUrl: data.pageUrl,
        pageContent: document.body.textContent || document.body.innerText || '',
        currentService: currentService
    };

    createAiDialog(pageData);
    
    // ダイアログが作成された後にVSCode設定解析を実行
    setTimeout(() => {
        const dialog = document.getElementById('ai-dialog');
        if (dialog) {
            analyzeVSCodeSettings(dialog);
        }
    }, 100);
}
