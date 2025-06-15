/*
 * AI業務支援ツール Edge拡張機能 - コンテンツスクリプト（リファクタリング版）
 * Copyright (c) 2024 AI Business Support Team
 */

import {
  getMessagingService,
  getSettingsService
} from '../services/index.js';

import {
  MESSAGE_ACTIONS,
  MESSAGE_TARGETS,
  ELEMENT_IDS,
  CSS_CLASSES
} from '../lib/constants.js';

import {
  Logger,
  debounce,
  sanitizeHtml,
  formatDateTime
} from '../lib/utils.js';

/**
 * コンテンツスクリプトのメインクラス
 */
class ContentScript {
  constructor() {
    this.isInitialized = false;
    this.messagingService = null;
    this.settingsService = null;
    this.currentService = 'unknown';
    this.aiButton = null;
    this.currentDialog = null;
    this.isProcessing = false;
  }

  /**
   * 初期化処理
   */
  async initialize() {
    if (this.isInitialized) {
      return;
    }

    try {
      Logger.info('コンテンツスクリプト初期化開始');

      // Services の取得
      this.messagingService = await getMessagingService();
      this.settingsService = await getSettingsService();

      // 現在のサービスを判定
      this.detectCurrentService();

      // メッセージハンドラーの登録
      this.registerMessageHandlers();

      // UI要素の初期化
      await this.initializeUI();

      // URL変更監視の開始
      this.observeUrlChanges();

      this.isInitialized = true;
      Logger.info('コンテンツスクリプト初期化完了', { service: this.currentService });

    } catch (error) {
      Logger.error('コンテンツスクリプト初期化エラー', error);
    }
  }

  /**
   * 現在のサービスを判定
   */
  detectCurrentService() {
    const hostname = window.location.hostname;

    if (hostname.includes('outlook.office.com') || hostname.includes('outlook.live.com')) {
      this.currentService = 'outlook';
    } else if (hostname.includes('mail.google.com')) {
      this.currentService = 'gmail';
    } else {
      this.currentService = 'general';
    }

    Logger.debug('サービス判定完了', { service: this.currentService, hostname });
  }

  /**
   * メッセージハンドラーを登録
   */
  registerMessageHandlers() {
    this.messagingService.registerHandler(MESSAGE_ACTIONS.ANALYZE_EMAIL, this.handleAnalyzeEmail.bind(this));
    this.messagingService.registerHandler(MESSAGE_ACTIONS.ANALYZE_PAGE, this.handleAnalyzePage.bind(this));
    this.messagingService.registerHandler(MESSAGE_ACTIONS.ANALYZE_SELECTION, this.handleAnalyzeSelection.bind(this));
    this.messagingService.registerHandler(MESSAGE_ACTIONS.COMPOSE_EMAIL, this.handleComposeEmail.bind(this));
    this.messagingService.registerHandler(MESSAGE_ACTIONS.TRANSLATE_SELECTION, this.handleTranslateSelection.bind(this));
    this.messagingService.registerHandler(MESSAGE_ACTIONS.TRANSLATE_PAGE, this.handleTranslatePage.bind(this));
    this.messagingService.registerHandler(MESSAGE_ACTIONS.EXTRACT_URLS, this.handleExtractUrls.bind(this));
    this.messagingService.registerHandler(MESSAGE_ACTIONS.COPY_PAGE_INFO, this.handleCopyPageInfo.bind(this));

    Logger.debug('メッセージハンドラー登録完了');
  }

  /**
   * UI要素を初期化
   */
  async initializeUI() {
    try {
      // AI支援ボタンを追加
      await this.addAISupportButton();

      // テーマを適用
      this.applyTheme();

      Logger.debug('UI初期化完了');
    } catch (error) {
      Logger.error('UI初期化エラー', error);
    }
  }

  /**
   * AI支援ボタンを追加
   */
  async addAISupportButton() {
    // 既存のボタンが存在する場合は削除
    const existingButton = document.getElementById(ELEMENT_IDS.AI_BUTTON);
    if (existingButton) {
      existingButton.remove();
    }

    // ボタン要素を作成
    this.aiButton = document.createElement('div');
    this.aiButton.id = ELEMENT_IDS.AI_BUTTON;
    this.aiButton.className = CSS_CLASSES.AI_BUTTON;
    this.aiButton.innerHTML = `
      <div class="ai-button-icon">🤖</div>
      <div class="ai-button-text">AI支援</div>
    `;

    // スタイルを設定
    this.aiButton.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      width: 80px;
      height: 80px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      border-radius: 50%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      z-index: 10000;
      box-shadow: 0 4px 20px rgba(0,0,0,0.15);
      transition: all 0.3s ease;
      color: white;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      font-size: 10px;
      font-weight: 600;
      user-select: none;
    `;

    // ホバー効果
    this.aiButton.addEventListener('mouseenter', () => {
      this.aiButton.style.transform = 'scale(1.1)';
      this.aiButton.style.boxShadow = '0 6px 25px rgba(0,0,0,0.2)';
    });

    this.aiButton.addEventListener('mouseleave', () => {
      this.aiButton.style.transform = 'scale(1)';
      this.aiButton.style.boxShadow = '0 4px 20px rgba(0,0,0,0.15)';
    });

    // クリックイベント
    this.aiButton.addEventListener('click', this.handleButtonClick.bind(this));

    // ドラッグ機能
    this.makeDraggable(this.aiButton);

    // ページに追加
    document.body.appendChild(this.aiButton);

    // 保存された位置を復元
    await this.restoreButtonPosition();

    Logger.debug('AI支援ボタン追加完了');
  }

  /**
   * ボタンクリック処理
   */
  async handleButtonClick(event) {
    event.stopPropagation();

    if (this.isProcessing) {
      Logger.debug('処理中のためクリックを無視');
      return;
    }

    try {
      await this.showAiDialog();
    } catch (error) {
      Logger.error('ボタンクリック処理エラー', error);
      this.showNotification('エラーが発生しました', 'error');
    }
  }

  /**
   * AIダイアログを表示
   */
  async showAiDialog() {
    try {
      // 既存のダイアログがある場合は閉じる
      if (this.currentDialog) {
        this.closeAiDialog();
      }

      // ダイアログ要素を作成
      this.currentDialog = this.createAiDialog();

      // ページに追加
      document.body.appendChild(this.currentDialog);

      // アニメーション
      setTimeout(() => {
        this.currentDialog.style.opacity = '1';
        this.currentDialog.querySelector('.ai-dialog-content').style.transform = 'scale(1)';
      }, 10);

      // ESCキーでの閉じる機能
      this.addEscapeKeyHandler();

      Logger.debug('AIダイアログ表示完了');

    } catch (error) {
      Logger.error('AIダイアログ表示エラー', error);
      throw error;
    }
  }

  /**
   * AIダイアログを作成
   */
  createAiDialog() {
    const dialog = document.createElement('div');
    dialog.id = ELEMENT_IDS.AI_DIALOG;
    dialog.className = CSS_CLASSES.AI_DIALOG;

    dialog.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0, 0, 0, 0.5);
      z-index: 10001;
      display: flex;
      align-items: center;
      justify-content: center;
      opacity: 0;
      transition: opacity 0.3s ease;
    `;

    const content = document.createElement('div');
    content.className = 'ai-dialog-content';
    content.style.cssText = `
      background: white;
      border-radius: 12px;
      padding: 24px;
      max-width: 600px;
      width: 90%;
      max-height: 80vh;
      overflow-y: auto;
      box-shadow: 0 20px 60px rgba(0,0,0,0.2);
      transform: scale(0.9);
      transition: transform 0.3s ease;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    `;

    content.innerHTML = this.getDialogHTML();
    dialog.appendChild(content);

    // イベントリスナーを設定
    this.setupDialogEventListeners(dialog);

    return dialog;
  }

  /**
   * ダイアログHTMLを取得
   */
  getDialogHTML() {
    const selectedText = this.getSelectedText();
    const hasSelection = selectedText && selectedText.length > 0;

    return `
      <div class="ai-dialog-header">
        <h2>🤖 AI業務支援ツール</h2>
        <button class="close-button" onclick="window.closeAiDialog()">✕</button>
      </div>
      
      <div class="ai-dialog-body">
        <div class="action-buttons">
          ${this.currentService === 'outlook' || this.currentService === 'gmail' ? `
            <button class="action-button primary" onclick="window.analyzeEmail()">
              📧 メール解析
            </button>
          ` : ''}
          
          <button class="action-button" onclick="window.analyzePage()">
            🌐 ページ解析
          </button>
          
          ${hasSelection ? `
            <button class="action-button" onclick="window.analyzeSelection()">
              📝 選択テキスト解析
            </button>
            <button class="action-button" onclick="window.translateSelection()">
              🌍 選択テキスト翻訳
            </button>
          ` : ''}
          
          <button class="action-button" onclick="window.extractUrls()">
            🔗 URL抽出
          </button>
          
          <button class="action-button" onclick="window.copyPageInfo()">
            📋 ページ情報コピー
          </button>
          
          ${this.currentService === 'outlook' || this.currentService === 'gmail' ? `
            <button class="action-button" onclick="window.composeEmail()">
              ✉️ メール作成支援
            </button>
          ` : ''}
        </div>
        
        <div class="quick-info">
          <div class="info-item">
            <strong>ページ:</strong> ${sanitizeHtml(document.title || '不明')}
          </div>
          <div class="info-item">
            <strong>URL:</strong> ${sanitizeHtml(window.location.href)}
          </div>
          ${hasSelection ? `
            <div class="info-item">
              <strong>選択テキスト:</strong> ${sanitizeHtml(selectedText.substring(0, 100))}${selectedText.length > 100 ? '...' : ''}
            </div>
          ` : ''}
        </div>
        
        <div class="result-area" id="ai-result-area" style="display: none;">
          <div class="result-header">
            <h3>📊 処理結果</h3>
            <div class="result-actions">
              <button class="action-button small" onclick="window.copyResult()">📋 コピー</button>
              <button class="action-button small" onclick="window.closeResult()">✕ 閉じる</button>
            </div>
          </div>
          <div class="result-content" id="ai-result-content"></div>
        </div>
      </div>
      
      <div class="ai-dialog-footer">
        <button class="action-button secondary" onclick="window.openSettings()">
          ⚙️ 設定
        </button>
      </div>
    `;
  }

  /**
   * ダイアログのイベントリスナーを設定
   */
  setupDialogEventListeners(dialog) {
    // 背景クリックで閉じる
    dialog.addEventListener('click', (event) => {
      if (event.target === dialog) {
        this.closeAiDialog();
      }
    });

    // グローバル関数を設定
    this.setupGlobalFunctions();
  }

  /**
   * グローバル関数を設定
   */
  setupGlobalFunctions() {
    window.analyzeEmail = this.analyzeEmail.bind(this);
    window.analyzePage = this.analyzePage.bind(this);
    window.analyzeSelection = this.analyzeSelection.bind(this);
    window.composeEmail = this.composeEmail.bind(this);
    window.translateSelection = this.translateSelection.bind(this);
    window.extractUrls = this.extractUrls.bind(this);
    window.copyPageInfo = this.copyPageInfo.bind(this);
    window.openSettings = this.openSettings.bind(this);
    window.closeAiDialog = this.closeAiDialog.bind(this);
    window.copyResult = this.copyResult.bind(this);
    window.closeResult = this.closeResult.bind(this);
  }

  /**
   * メール解析を実行
   */
  async analyzeEmail() {
    if (this.isProcessing) return;

    try {
      this.isProcessing = true;
      this.showLoading('メールを解析中...');

      const emailData = this.getCurrentEmailData();

      if (!emailData.body) {
        throw new Error('メールの内容が取得できませんでした');
      }

      const result = await this.messagingService.sendToBackground(
        MESSAGE_ACTIONS.ANALYZE_EMAIL,
        emailData
      );

      this.showResult(result.data || result, 'success');

    } catch (error) {
      Logger.error('メール解析エラー', error);
      this.showResult(`エラー: ${error.message}`, 'error');
    } finally {
      this.isProcessing = false;
      this.hideLoading();
    }
  }

  /**
   * ページ解析を実行
   */
  async analyzePage() {
    if (this.isProcessing) return;

    try {
      this.isProcessing = true;
      this.showLoading('ページを解析中...');

      const pageData = this.extractPageContent();

      const result = await this.messagingService.sendToBackground(
        MESSAGE_ACTIONS.ANALYZE_PAGE,
        pageData
      );

      this.showResult(result.data || result, 'success');

    } catch (error) {
      Logger.error('ページ解析エラー', error);
      this.showResult(`エラー: ${error.message}`, 'error');
    } finally {
      this.isProcessing = false;
      this.hideLoading();
    }
  }

  /**
   * 選択テキスト解析を実行
   */
  async analyzeSelection() {
    if (this.isProcessing) return;

    try {
      this.isProcessing = true;
      this.showLoading('選択テキストを解析中...');

      const selectedText = this.getSelectedText();

      if (!selectedText) {
        throw new Error('テキストが選択されていません');
      }

      const selectionData = {
        selectedText: selectedText,
        url: window.location.href,
        pageTitle: document.title
      };

      const result = await this.messagingService.sendToBackground(
        MESSAGE_ACTIONS.ANALYZE_SELECTION,
        selectionData
      );

      this.showResult(result.data || result, 'success');

    } catch (error) {
      Logger.error('選択テキスト解析エラー', error);
      this.showResult(`エラー: ${error.message}`, 'error');
    } finally {
      this.isProcessing = false;
      this.hideLoading();
    }
  }

  /**
   * メール作成支援を実行
   */
  async composeEmail() {
    if (this.isProcessing) return;

    try {
      // 簡単なプロンプトダイアログ
      const purpose = prompt('メールの目的を入力してください:');
      if (!purpose) return;

      const recipient = prompt('宛先を入力してください（省略可）:') || '相手先';
      const details = prompt('詳細な指示を入力してください:') || '';

      this.isProcessing = true;
      this.showLoading('メールを作成中...');

      const requestData = {
        type: '一般',
        purpose: purpose,
        recipient: recipient,
        details: details
      };

      const result = await this.messagingService.sendToBackground(
        MESSAGE_ACTIONS.COMPOSE_EMAIL,
        requestData
      );

      this.showResult(result.data || result, 'success');

    } catch (error) {
      Logger.error('メール作成エラー', error);
      this.showResult(`エラー: ${error.message}`, 'error');
    } finally {
      this.isProcessing = false;
      this.hideLoading();
    }
  }

  /**
   * 選択テキスト翻訳を実行
   */
  async translateSelection() {
    if (this.isProcessing) return;

    try {
      this.isProcessing = true;
      this.showLoading('翻訳中...');

      const selectedText = this.getSelectedText();

      if (!selectedText) {
        throw new Error('テキストが選択されていません');
      }

      const targetLanguage = prompt('翻訳先言語を入力してください（例: 英語、中国語）:', '英語') || '英語';

      const translationData = {
        selectedText: selectedText,
        targetLanguage: targetLanguage,
        url: window.location.href
      };

      const result = await this.messagingService.sendToBackground(
        MESSAGE_ACTIONS.TRANSLATE_SELECTION,
        translationData
      );

      this.showResult(result.data || result, 'success');

    } catch (error) {
      Logger.error('翻訳エラー', error);
      this.showResult(`エラー: ${error.message}`, 'error');
    } finally {
      this.isProcessing = false;
      this.hideLoading();
    }
  }

  /**
   * URL抽出を実行
   */
  async extractUrls() {
    if (this.isProcessing) return;

    try {
      this.isProcessing = true;
      this.showLoading('URLを抽出中...');

      const pageData = this.extractPageContent();

      const result = await this.messagingService.sendToBackground(
        MESSAGE_ACTIONS.EXTRACT_URLS,
        pageData
      );

      this.showResult(result.data || result, 'success');

    } catch (error) {
      Logger.error('URL抽出エラー', error);
      this.showResult(`エラー: ${error.message}`, 'error');
    } finally {
      this.isProcessing = false;
      this.hideLoading();
    }
  }

  /**
   * ページ情報コピーを実行
   */
  async copyPageInfo() {
    try {
      const pageData = this.extractPageContent();

      const result = await this.messagingService.sendToBackground(
        MESSAGE_ACTIONS.COPY_PAGE_INFO,
        pageData
      );

      // クリップボードにコピー
      await navigator.clipboard.writeText(result.data || result);
      this.showNotification('ページ情報をクリップボードにコピーしました', 'success');

    } catch (error) {
      Logger.error('ページ情報コピーエラー', error);
      this.showNotification('コピーに失敗しました', 'error');
    }
  }

  /**
   * 設定画面を開く
   */
  openSettings() {
    chrome.runtime.openOptionsPage();
  }

  /**
   * AIダイアログを閉じる
   */
  closeAiDialog() {
    if (this.currentDialog) {
      this.currentDialog.style.opacity = '0';
      setTimeout(() => {
        if (this.currentDialog && this.currentDialog.parentNode) {
          this.currentDialog.parentNode.removeChild(this.currentDialog);
        }
        this.currentDialog = null;
      }, 300);
    }

    this.removeEscapeKeyHandler();
  }

  /**
   * 結果をコピー
   */
  async copyResult() {
    try {
      const resultContent = document.getElementById('ai-result-content');
      if (resultContent) {
        await navigator.clipboard.writeText(resultContent.textContent);
        this.showNotification('結果をクリップボードにコピーしました', 'success');
      }
    } catch (error) {
      Logger.error('結果コピーエラー', error);
      this.showNotification('コピーに失敗しました', 'error');
    }
  }

  /**
   * 結果表示を閉じる
   */
  closeResult() {
    const resultArea = document.getElementById('ai-result-area');
    if (resultArea) {
      resultArea.style.display = 'none';
    }
  }

  /**
   * メッセージハンドラー: メール解析
   */
  async handleAnalyzeEmail(data) {
    return await this.analyzeEmail();
  }

  /**
   * メッセージハンドラー: ページ解析
   */
  async handleAnalyzePage(data) {
    return await this.analyzePage();
  }

  /**
   * メッセージハンドラー: 選択テキスト解析
   */
  async handleAnalyzeSelection(data) {
    if (data.selectedText) {
      // 外部から指定されたテキストがある場合は、それを選択状態にする
      // ここでは簡易実装として、グローバル変数に保存
      window._externalSelectedText = data.selectedText;
    }
    return await this.analyzeSelection();
  }

  /**
   * メッセージハンドラー: メール作成
   */
  async handleComposeEmail(data) {
    return await this.composeEmail();
  }

  /**
   * メッセージハンドラー: 翻訳
   */
  async handleTranslateSelection(data) {
    if (data.selectedText) {
      window._externalSelectedText = data.selectedText;
    }
    return await this.translateSelection();
  }

  /**
   * メッセージハンドラー: ページ翻訳
   */
  async handleTranslatePage(data) {
    // ページ全体の翻訳（選択テキスト翻訳と同様の処理）
    return await this.translateSelection();
  }

  /**
   * メッセージハンドラー: URL抽出
   */
  async handleExtractUrls(data) {
    return await this.extractUrls();
  }

  /**
   * メッセージハンドラー: ページ情報コピー
   */
  async handleCopyPageInfo(data) {
    return await this.copyPageInfo();
  }

  /**
   * 現在のメールデータを取得
   */
  getCurrentEmailData() {
    try {
      if (this.currentService === 'outlook') {
        return this.getOutlookEmailData();
      } else if (this.currentService === 'gmail') {
        return this.getGmailEmailData();
      } else {
        throw new Error('メールサービスが検出されませんでした');
      }
    } catch (error) {
      Logger.error('メールデータ取得エラー', error);
      return {
        from: '不明',
        subject: '件名不明',
        body: 'メール本文が取得できませんでした',
        date: new Date().toISOString()
      };
    }
  }

  /**
   * Outlookのメールデータを取得
   */
  getOutlookEmailData() {
    // Outlook Web App の構造に基づく取得
    const subjectElement = document.querySelector('[data-testid="message-subject"]') ||
      document.querySelector('.allowTextSelection[role="heading"]');

    const fromElement = document.querySelector('[data-testid="message-from"]') ||
      document.querySelector('button[data-testid="persona-button"] span');

    const bodyElement = document.querySelector('[data-testid="message-body"]') ||
      document.querySelector('.elementToProof') ||
      document.querySelector('.rps_7cb5');

    return {
      from: fromElement?.textContent?.trim() || '不明',
      subject: subjectElement?.textContent?.trim() || '件名不明',
      body: bodyElement?.textContent?.trim() || 'メール本文が取得できませんでした',
      date: new Date().toISOString()
    };
  }

  /**
   * Gmailのメールデータを取得
   */
  getGmailEmailData() {
    // Gmail の構造に基づく取得
    const subjectElement = document.querySelector('h2[data-legacy-thread-id]') ||
      document.querySelector('.hP');

    const fromElement = document.querySelector('.go span[email]') ||
      document.querySelector('.gD');

    const bodyElement = document.querySelector('.ii.gt .a3s.aiL') ||
      document.querySelector('.ii.gt div[dir="ltr"]');

    return {
      from: fromElement?.getAttribute('email') || fromElement?.textContent?.trim() || '不明',
      subject: subjectElement?.textContent?.trim() || '件名不明',
      body: bodyElement?.textContent?.trim() || 'メール本文が取得できませんでした',
      date: new Date().toISOString()
    };
  }

  /**
   * ページコンテンツを抽出
   */
  extractPageContent() {
    try {
      // 基本情報
      const title = document.title;
      const url = window.location.href;

      // メインコンテンツを抽出
      let content = '';

      // 主要なコンテンツ要素を優先的に抽出
      const contentSelectors = [
        'main',
        'article',
        '[role="main"]',
        '.main-content',
        '.content',
        '#content',
        '.post-content',
        '.entry-content'
      ];

      for (const selector of contentSelectors) {
        const element = document.querySelector(selector);
        if (element) {
          content = element.textContent?.trim();
          break;
        }
      }

      // フォールバック: body全体から取得
      if (!content) {
        content = document.body.textContent?.trim() || '';
      }

      // 長すぎる場合は制限
      if (content.length > 10000) {
        content = content.substring(0, 10000) + '...';
      }

      return {
        title: title,
        url: url,
        content: content
      };

    } catch (error) {
      Logger.error('ページコンテンツ抽出エラー', error);
      return {
        title: document.title || 'タイトル不明',
        url: window.location.href,
        content: 'コンテンツが取得できませんでした'
      };
    }
  }

  /**
   * 選択されたテキストを取得
   */
  getSelectedText() {
    // 外部から指定されたテキストがある場合はそれを使用
    if (window._externalSelectedText) {
      const text = window._externalSelectedText;
      delete window._externalSelectedText;
      return text;
    }

    const selection = window.getSelection();
    return selection.toString().trim();
  }

  /**
   * ローディング表示
   */
  showLoading(message = '処理中...') {
    // 既存のローディングを削除
    this.hideLoading();

    const loading = document.createElement('div');
    loading.id = ELEMENT_IDS.LOADING;
    loading.className = CSS_CLASSES.LOADING;
    loading.style.cssText = `
      position: fixed;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      background: rgba(0, 0, 0, 0.8);
      color: white;
      padding: 20px 30px;
      border-radius: 8px;
      z-index: 10002;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      display: flex;
      align-items: center;
      gap: 12px;
    `;

    loading.innerHTML = `
      <div style="
        width: 20px;
        height: 20px;
        border: 2px solid transparent;
        border-top: 2px solid #fff;
        border-radius: 50%;
        animation: spin 1s linear infinite;
      "></div>
      <span>${sanitizeHtml(message)}</span>
    `;

    // CSS アニメーションを追加
    const style = document.createElement('style');
    style.textContent = `
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    `;
    document.head.appendChild(style);

    document.body.appendChild(loading);
  }

  /**
   * ローディング非表示
   */
  hideLoading() {
    const loading = document.getElementById(ELEMENT_IDS.LOADING);
    if (loading) {
      loading.remove();
    }
  }

  /**
   * 結果を表示
   */
  showResult(result, type = 'success') {
    const resultArea = document.getElementById('ai-result-area');
    const resultContent = document.getElementById('ai-result-content');

    if (resultArea && resultContent) {
      resultContent.innerHTML = `
        <div class="result-message ${type}">
          <pre style="white-space: pre-wrap; font-family: inherit;">${sanitizeHtml(result)}</pre>
        </div>
      `;

      resultArea.style.display = 'block';

      // 結果エリアまでスクロール
      resultArea.scrollIntoView({ behavior: 'smooth' });
    } else {
      // フォールバック通知
      this.showNotification(result, type);
    }
  }

  /**
   * 通知を表示
   */
  showNotification(message, type = 'info') {
    // 既存の通知を削除
    const existingNotification = document.getElementById(ELEMENT_IDS.NOTIFICATION);
    if (existingNotification) {
      existingNotification.remove();
    }

    const notification = document.createElement('div');
    notification.id = ELEMENT_IDS.NOTIFICATION;
    notification.className = CSS_CLASSES.NOTIFICATION;

    const colors = {
      success: '#4CAF50',
      error: '#f44336',
      warning: '#ff9800',
      info: '#2196F3'
    };

    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: ${colors[type] || colors.info};
      color: white;
      padding: 12px 20px;
      border-radius: 6px;
      z-index: 10003;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      font-size: 14px;
      max-width: 400px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.15);
      animation: slideIn 0.3s ease;
    `;

    notification.textContent = message;
    document.body.appendChild(notification);

    // 3秒後に自動削除
    setTimeout(() => {
      if (notification.parentNode) {
        notification.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => {
          if (notification.parentNode) {
            notification.remove();
          }
        }, 300);
      }
    }, 3000);
  }

  /**
   * テーマを適用
   */
  applyTheme() {
    // ダークモード検出
    const isDarkMode = window.matchMedia('(prefers-color-scheme: dark)').matches ||
      document.documentElement.classList.contains('dark') ||
      document.body.classList.contains('dark-theme');

    if (isDarkMode) {
      document.body.classList.add(CSS_CLASSES.DARK_THEME);
    } else {
      document.body.classList.add(CSS_CLASSES.LIGHT_THEME);
    }
  }

  /**
   * 要素をドラッグ可能にする
   */
  makeDraggable(element) {
    let isDragging = false;
    let startX, startY, startLeft, startTop;

    element.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return; // 左クリックのみ

      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      startLeft = element.offsetLeft;
      startTop = element.offsetTop;

      e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;

      const deltaX = e.clientX - startX;
      const deltaY = e.clientY - startY;

      element.style.left = (startLeft + deltaX) + 'px';
      element.style.top = (startTop + deltaY) + 'px';
      element.style.right = 'auto';
    });

    document.addEventListener('mouseup', async () => {
      if (isDragging) {
        isDragging = false;

        // 位置を保存
        await this.saveButtonPosition({
          left: element.offsetLeft,
          top: element.offsetTop
        });
      }
    });
  }

  /**
   * ボタン位置を保存
   */
  async saveButtonPosition(position) {
    try {
      await chrome.storage.local.set({
        aiButtonPosition: position
      });
    } catch (error) {
      Logger.error('ボタン位置保存エラー', error);
    }
  }

  /**
   * ボタン位置を復元
   */
  async restoreButtonPosition() {
    try {
      const result = await chrome.storage.local.get('aiButtonPosition');
      const position = result.aiButtonPosition;

      if (position && this.aiButton) {
        this.aiButton.style.left = position.left + 'px';
        this.aiButton.style.top = position.top + 'px';
        this.aiButton.style.right = 'auto';
      }
    } catch (error) {
      Logger.error('ボタン位置復元エラー', error);
    }
  }

  /**
   * ESCキーハンドラーを追加
   */
  addEscapeKeyHandler() {
    this.escapeKeyHandler = (event) => {
      if (event.key === 'Escape') {
        this.closeAiDialog();
      }
    };
    document.addEventListener('keydown', this.escapeKeyHandler);
  }

  /**
   * ESCキーハンドラーを削除
   */
  removeEscapeKeyHandler() {
    if (this.escapeKeyHandler) {
      document.removeEventListener('keydown', this.escapeKeyHandler);
      this.escapeKeyHandler = null;
    }
  }

  /**
   * URL変更を監視（SPA対応）
   */
  observeUrlChanges() {
    let currentUrl = window.location.href;

    const checkUrlChange = () => {
      if (window.location.href !== currentUrl) {
        currentUrl = window.location.href;
        Logger.debug('URL変更検出', { newUrl: currentUrl });

        // サービス再判定
        this.detectCurrentService();

        // UIを再初期化（必要に応じて）
        debounce(() => {
          this.initializeUI();
        }, 1000)();
      }
    };

    // MutationObserver でDOM変更を監視
    const observer = new MutationObserver(checkUrlChange);
    observer.observe(document.body, {
      childList: true,
      subtree: true
    });

    // popstate イベントも監視
    window.addEventListener('popstate', checkUrlChange);
  }
}

// グローバルインスタンス
const contentScript = new ContentScript();

// DOM読み込み完了を待機
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    contentScript.initialize();
  });
} else {
  contentScript.initialize();
}

// グローバルエクスポート（デバッグ用）
globalThis.contentScript = contentScript;

Logger.info('AI支援ツール - コンテンツスクリプト読み込み完了');
