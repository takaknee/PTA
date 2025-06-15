/*
 * AI業務支援ツール Edge拡張機能 - バックグラウンドサービスワーカー（リファクタリング版）
 * Copyright (c) 2024 AI Business Support Team
 */

import {
  getServices,
  getApiService,
  getSettingsService,
  getHistoryService,
  getStatisticsService,
  getMessagingService,
  getPromptService
} from '../services/index.js';

import {
  MESSAGE_ACTIONS,
  CONTEXT_MENU_IDS,
  HISTORY_TYPES,
  ERROR_MESSAGES
} from '../lib/constants.js';

import { Logger, performanceMonitor } from '../lib/utils.js';
import { HistoryItem, DiagnosticsInfo } from '../lib/models.js';

/**
 * バックグラウンドサービスワーカーのメインクラス
 */
class BackgroundService {
  constructor() {
    this.isInitialized = false;
    this.services = null;
  }

  /**
   * サービスの初期化
   */
  async initialize() {
    try {
      Logger.info('AI業務支援ツール - バックグラウンドサービス初期化開始');

      // Services の初期化
      this.services = await getServices();

      // メッセージハンドラーの登録
      this.registerMessageHandlers();

      // Chrome拡張機能のイベントリスナー設定
      this.setupChromeEventListeners();

      this.isInitialized = true;
      Logger.info('バックグラウンドサービス初期化完了');

    } catch (error) {
      Logger.error('バックグラウンドサービス初期化エラー', error);
      throw error;
    }
  }

  /**
   * メッセージハンドラーを登録
   */
  registerMessageHandlers() {
    const messagingService = this.services.messaging;

    // AI解析系ハンドラー
    messagingService.registerHandler(MESSAGE_ACTIONS.ANALYZE_EMAIL, this.handleAnalyzeEmail.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.ANALYZE_PAGE, this.handleAnalyzePage.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.ANALYZE_SELECTION, this.handleAnalyzeSelection.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.COMPOSE_EMAIL, this.handleComposeEmail.bind(this));

    // 翻訳・抽出系ハンドラー
    messagingService.registerHandler(MESSAGE_ACTIONS.TRANSLATE_SELECTION, this.handleTranslateSelection.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.TRANSLATE_PAGE, this.handleTranslatePage.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.EXTRACT_URLS, this.handleExtractUrls.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.COPY_PAGE_INFO, this.handleCopyPageInfo.bind(this));

    // システム系ハンドラー
    messagingService.registerHandler(MESSAGE_ACTIONS.TEST_API, this.handleApiTest.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.CONNECTION_TEST, this.handleConnectionTest.bind(this));
    messagingService.registerHandler(MESSAGE_ACTIONS.SYSTEM_DIAGNOSTICS, this.handleSystemDiagnostics.bind(this));

    Logger.info('メッセージハンドラー登録完了');
  }

  /**
   * Chrome拡張機能のイベントリスナーを設定
   */
  setupChromeEventListeners() {
    // 拡張機能インストール時
    chrome.runtime.onInstalled.addListener(this.handleInstalled.bind(this));

    // コンテキストメニュークリック
    chrome.contextMenus.onClicked.addListener(this.handleContextMenuClick.bind(this));

    Logger.info('Chromeイベントリスナー設定完了');
  }

  /**
   * 拡張機能インストール時の処理
   */
  async handleInstalled(details) {
    Logger.info('AI業務支援ツールがインストールされました', { reason: details.reason });

    try {
      if (details.reason === 'install') {
        // 初回インストール時の初期設定
        await this.performInitialSetup();
      }

      // コンテキストメニューを作成
      await this.createContextMenus();

    } catch (error) {
      Logger.error('インストール時処理エラー', error);
    }
  }

  /**
   * 初回インストール時の設定
   */
  async performInitialSetup() {
    try {
      // デフォルト設定の確認・保存
      const settings = await this.services.settings.getSettings();
      await this.services.settings.saveSettings(settings);

      Logger.info('初期設定完了');
    } catch (error) {
      Logger.error('初期設定エラー', error);
    }
  }

  /**
   * コンテキストメニューを作成
   */
  async createContextMenus() {
    try {
      // 既存のメニューをクリア
      await chrome.contextMenus.removeAll();

      // メインメニュー
      chrome.contextMenus.create({
        id: 'ai-tool-main',
        title: 'AI業務支援ツール',
        contexts: ['page', 'selection']
      });

      // ページ解析メニュー
      chrome.contextMenus.create({
        id: CONTEXT_MENU_IDS.ANALYZE_PAGE,
        parentId: 'ai-tool-main',
        title: 'ページを解析',
        contexts: ['page']
      });

      // 選択テキスト解析メニュー
      chrome.contextMenus.create({
        id: CONTEXT_MENU_IDS.ANALYZE_SELECTION,
        parentId: 'ai-tool-main',
        title: '選択テキストを解析',
        contexts: ['selection']
      });

      // 翻訳メニュー
      chrome.contextMenus.create({
        id: CONTEXT_MENU_IDS.TRANSLATE_SELECTION,
        parentId: 'ai-tool-main',
        title: '選択テキストを翻訳',
        contexts: ['selection']
      });

      // URL抽出メニュー
      chrome.contextMenus.create({
        id: CONTEXT_MENU_IDS.EXTRACT_URLS,
        parentId: 'ai-tool-main',
        title: 'URLを抽出',
        contexts: ['page']
      });

      // ページ情報コピーメニュー
      chrome.contextMenus.create({
        id: CONTEXT_MENU_IDS.COPY_PAGE_INFO,
        parentId: 'ai-tool-main',
        title: 'ページ情報をコピー',
        contexts: ['page']
      });

      Logger.info('コンテキストメニュー作成完了');

    } catch (error) {
      Logger.error('コンテキストメニュー作成エラー', error);
    }
  }

  /**
   * コンテキストメニュークリック処理
   */
  async handleContextMenuClick(info, tab) {
    try {
      Logger.info('コンテキストメニュークリック', { menuItemId: info.menuItemId });

      const messagingService = this.services.messaging;

      switch (info.menuItemId) {
        case CONTEXT_MENU_IDS.ANALYZE_PAGE:
          await messagingService.sendToActiveTab(MESSAGE_ACTIONS.ANALYZE_PAGE);
          break;

        case CONTEXT_MENU_IDS.ANALYZE_SELECTION:
          await messagingService.sendToActiveTab(MESSAGE_ACTIONS.ANALYZE_SELECTION, {
            selectedText: info.selectionText
          });
          break;

        case CONTEXT_MENU_IDS.TRANSLATE_SELECTION:
          await messagingService.sendToActiveTab(MESSAGE_ACTIONS.TRANSLATE_SELECTION, {
            selectedText: info.selectionText
          });
          break;

        case CONTEXT_MENU_IDS.EXTRACT_URLS:
          await messagingService.sendToActiveTab(MESSAGE_ACTIONS.EXTRACT_URLS);
          break;

        case CONTEXT_MENU_IDS.COPY_PAGE_INFO:
          await messagingService.sendToActiveTab(MESSAGE_ACTIONS.COPY_PAGE_INFO);
          break;
      }

    } catch (error) {
      Logger.error('コンテキストメニュー処理エラー', error);
    }
  }

  /**
   * メール解析処理
   */
  async handleAnalyzeEmail(data) {
    const startTime = Date.now();
    performanceMonitor.start('analyzeEmail');

    try {
      Logger.info('メール解析処理開始');

      // プロンプトを生成
      const promptService = this.services.prompt;
      const prompt = promptService.createEmailAnalysisPrompt(data);

      // AI API を呼び出し
      const apiService = this.services.api;
      const response = await apiService.callAI(prompt.userPrompt, {
        systemPrompt: prompt.systemPrompt
      });

      if (!response.isSuccess()) {
        throw new Error(response.error);
      }

      // 履歴に保存
      await this.saveToHistory({
        type: HISTORY_TYPES.EMAIL_ANALYSIS,
        title: `メール解析: ${data.subject || '件名なし'}`,
        content: data.body || '',
        result: response.data,
        metadata: {
          from: data.from,
          subject: data.subject,
          date: data.date
        }
      });

      // 統計を記録
      const duration = performanceMonitor.end('analyzeEmail');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.EMAIL_ANALYSIS,
        true,
        duration,
        { provider: (await this.services.settings.getSettings()).provider }
      );

      Logger.info('メール解析処理完了', { duration: `${duration.toFixed(2)}ms` });
      return response.data;

    } catch (error) {
      const duration = performanceMonitor.end('analyzeEmail');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.EMAIL_ANALYSIS,
        false,
        duration
      );

      Logger.error('メール解析処理エラー', error);
      throw error;
    }
  }

  /**
   * ページ解析処理
   */
  async handleAnalyzePage(data) {
    const startTime = Date.now();
    performanceMonitor.start('analyzePage');

    try {
      Logger.info('ページ解析処理開始');

      // プロンプトを生成
      const promptService = this.services.prompt;
      const prompt = promptService.createPageAnalysisPrompt(data);

      // AI API を呼び出し
      const apiService = this.services.api;
      const response = await apiService.callAI(prompt.userPrompt, {
        systemPrompt: prompt.systemPrompt
      });

      if (!response.isSuccess()) {
        throw new Error(response.error);
      }

      // 履歴に保存
      await this.saveToHistory({
        type: HISTORY_TYPES.PAGE_ANALYSIS,
        title: `ページ解析: ${data.title || 'タイトル不明'}`,
        content: data.content || '',
        result: response.data,
        metadata: {
          url: data.url,
          title: data.title
        }
      });

      // 統計を記録
      const duration = performanceMonitor.end('analyzePage');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.PAGE_ANALYSIS,
        true,
        duration
      );

      Logger.info('ページ解析処理完了');
      return response.data;

    } catch (error) {
      const duration = performanceMonitor.end('analyzePage');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.PAGE_ANALYSIS,
        false,
        duration
      );

      Logger.error('ページ解析処理エラー', error);
      throw error;
    }
  }

  /**
   * 選択テキスト解析処理
   */
  async handleAnalyzeSelection(data) {
    const startTime = Date.now();
    performanceMonitor.start('analyzeSelection');

    try {
      Logger.info('選択テキスト解析処理開始');

      // プロンプトを生成
      const promptService = this.services.prompt;
      const prompt = promptService.createSelectionAnalysisPrompt(data);

      // AI API を呼び出し
      const apiService = this.services.api;
      const response = await apiService.callAI(prompt.userPrompt, {
        systemPrompt: prompt.systemPrompt
      });

      if (!response.isSuccess()) {
        throw new Error(response.error);
      }

      // 履歴に保存
      await this.saveToHistory({
        type: HISTORY_TYPES.SELECTION_ANALYSIS,
        title: `選択テキスト解析: ${data.selectedText?.substring(0, 50) || ''}...`,
        content: data.selectedText || '',
        result: response.data,
        metadata: {
          url: data.url,
          pageTitle: data.pageTitle
        }
      });

      // 統計を記録
      const duration = performanceMonitor.end('analyzeSelection');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.SELECTION_ANALYSIS,
        true,
        duration
      );

      Logger.info('選択テキスト解析処理完了');
      return response.data;

    } catch (error) {
      const duration = performanceMonitor.end('analyzeSelection');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.SELECTION_ANALYSIS,
        false,
        duration
      );

      Logger.error('選択テキスト解析処理エラー', error);
      throw error;
    }
  }

  /**
   * メール作成処理
   */
  async handleComposeEmail(data) {
    const startTime = Date.now();
    performanceMonitor.start('composeEmail');

    try {
      Logger.info('メール作成処理開始');

      // プロンプトを生成
      const promptService = this.services.prompt;
      const prompt = promptService.createEmailCompositionPrompt(data);

      // AI API を呼び出し
      const apiService = this.services.api;
      const response = await apiService.callAI(prompt.userPrompt, {
        systemPrompt: prompt.systemPrompt
      });

      if (!response.isSuccess()) {
        throw new Error(response.error);
      }

      // 履歴に保存
      await this.saveToHistory({
        type: HISTORY_TYPES.EMAIL_COMPOSITION,
        title: `メール作成: ${data.purpose || '目的不明'}`,
        content: data.details || '',
        result: response.data,
        metadata: {
          type: data.type,
          recipient: data.recipient,
          purpose: data.purpose
        }
      });

      // 統計を記録
      const duration = performanceMonitor.end('composeEmail');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.EMAIL_COMPOSITION,
        true,
        duration
      );

      Logger.info('メール作成処理完了');
      return response.data;

    } catch (error) {
      const duration = performanceMonitor.end('composeEmail');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.EMAIL_COMPOSITION,
        false,
        duration
      );

      Logger.error('メール作成処理エラー', error);
      throw error;
    }
  }

  /**
   * 翻訳処理
   */
  async handleTranslateSelection(data) {
    const startTime = Date.now();
    performanceMonitor.start('translateSelection');

    try {
      Logger.info('翻訳処理開始');

      // プロンプトを生成
      const promptService = this.services.prompt;
      const prompt = promptService.createTranslationPrompt({
        text: data.selectedText,
        targetLanguage: data.targetLanguage || '日本語'
      });

      // AI API を呼び出し
      const apiService = this.services.api;
      const response = await apiService.callAI(prompt.userPrompt, {
        systemPrompt: prompt.systemPrompt
      });

      if (!response.isSuccess()) {
        throw new Error(response.error);
      }

      // 履歴に保存
      await this.saveToHistory({
        type: HISTORY_TYPES.TRANSLATION,
        title: `翻訳: ${data.selectedText?.substring(0, 50) || ''}...`,
        content: data.selectedText || '',
        result: response.data,
        metadata: {
          targetLanguage: data.targetLanguage || '日本語',
          url: data.url
        }
      });

      // 統計を記録
      const duration = performanceMonitor.end('translateSelection');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.TRANSLATION,
        true,
        duration
      );

      Logger.info('翻訳処理完了');
      return response.data;

    } catch (error) {
      const duration = performanceMonitor.end('translateSelection');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.TRANSLATION,
        false,
        duration
      );

      Logger.error('翻訳処理エラー', error);
      throw error;
    }
  }

  /**
   * ページ翻訳処理
   */
  async handleTranslatePage(data) {
    return await this.handleTranslateSelection({
      selectedText: data.content,
      targetLanguage: data.targetLanguage,
      url: data.url
    });
  }

  /**
   * URL抽出処理
   */
  async handleExtractUrls(data) {
    const startTime = Date.now();
    performanceMonitor.start('extractUrls');

    try {
      Logger.info('URL抽出処理開始');

      // プロンプトを生成
      const promptService = this.services.prompt;
      const prompt = promptService.createUrlExtractionPrompt({
        text: data.content || data.selectedText
      });

      // AI API を呼び出し
      const apiService = this.services.api;
      const response = await apiService.callAI(prompt.userPrompt, {
        systemPrompt: prompt.systemPrompt
      });

      if (!response.isSuccess()) {
        throw new Error(response.error);
      }

      // 履歴に保存
      await this.saveToHistory({
        type: HISTORY_TYPES.URL_EXTRACTION,
        title: `URL抽出: ${data.url || 'ページ'}`,
        content: data.content || data.selectedText || '',
        result: response.data,
        metadata: {
          url: data.url,
          pageTitle: data.pageTitle
        }
      });

      // 統計を記録
      const duration = performanceMonitor.end('extractUrls');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.URL_EXTRACTION,
        true,
        duration
      );

      Logger.info('URL抽出処理完了');
      return response.data;

    } catch (error) {
      const duration = performanceMonitor.end('extractUrls');
      await this.services.statistics.recordRequest(
        HISTORY_TYPES.URL_EXTRACTION,
        false,
        duration
      );

      Logger.error('URL抽出処理エラー', error);
      throw error;
    }
  }

  /**
   * ページ情報コピー処理
   */
  async handleCopyPageInfo(data) {
    try {
      Logger.info('ページ情報コピー処理開始');

      const pageInfo = `# ${data.title || 'タイトル不明'}

**URL**: ${data.url || '不明'}
**取得日時**: ${new Date().toLocaleString('ja-JP')}

## 内容
${data.content || 'コンテンツが取得できませんでした'}
`;

      // クリップボードにコピー（Content scriptで実行されることを想定）
      return pageInfo;

    } catch (error) {
      Logger.error('ページ情報コピー処理エラー', error);
      throw error;
    }
  }

  /**
   * API接続テスト
   */
  async handleApiTest(data) {
    try {
      Logger.info('API接続テスト開始');

      const apiService = this.services.api;
      const response = await apiService.testConnection();

      return {
        success: response.isSuccess(),
        message: response.isSuccess() ? '接続成功' : response.error,
        details: response.toJSON()
      };

    } catch (error) {
      Logger.error('API接続テストエラー', error);
      return {
        success: false,
        message: error.message,
        details: null
      };
    }
  }

  /**
   * 接続テスト処理
   */
  async handleConnectionTest(data) {
    try {
      Logger.info('接続テスト開始');

      // 基本的な接続テスト
      const testResult = await this.handleApiTest(data);

      // システム診断情報も含める
      const diagnostics = await this.handleSystemDiagnostics();

      return {
        connectionTest: testResult,
        diagnostics: diagnostics
      };

    } catch (error) {
      Logger.error('接続テストエラー', error);
      throw error;
    }
  }

  /**
   * システム診断処理
   */
  async handleSystemDiagnostics() {
    try {
      Logger.info('システム診断開始');

      const diagnostics = new DiagnosticsInfo({
        chrome: {
          version: navigator.userAgent,
          offscreenSupport: typeof chrome.offscreen !== 'undefined',
          runtimeSupport: typeof chrome.runtime !== 'undefined'
        },
        settings: await this.services.settings.validateCurrentSettings(),
        statistics: await this.services.statistics.getStatistics()
      });

      // 権限チェック
      try {
        diagnostics.permissions = {
          activeTab: await chrome.permissions.contains({ permissions: ['activeTab'] }),
          storage: await chrome.permissions.contains({ permissions: ['storage'] }),
          scripting: await chrome.permissions.contains({ permissions: ['scripting'] }),
          contextMenus: await chrome.permissions.contains({ permissions: ['contextMenus'] })
        };
      } catch (error) {
        diagnostics.permissions = { error: error.message };
      }

      // ネットワーク接続チェック
      try {
        const testResponse = await fetch('https://www.google.com', { method: 'HEAD' });
        diagnostics.network.basicConnectivity = testResponse.ok;
      } catch {
        diagnostics.network.basicConnectivity = false;
      }

      Logger.info('システム診断完了');
      return diagnostics.toJSON();

    } catch (error) {
      Logger.error('システム診断エラー', error);
      throw error;
    }
  }

  /**
   * 履歴に保存
   */
  async saveToHistory(entry) {
    try {
      const historyItem = new HistoryItem(entry);
      await this.services.history.addHistoryItem(historyItem);
      Logger.debug('履歴保存完了', { id: historyItem.id });
    } catch (error) {
      Logger.error('履歴保存エラー', error);
      // 履歴保存エラーは処理継続
    }
  }
}

// グローバルインスタンス
const backgroundService = new BackgroundService();

// 拡張機能の開始
backgroundService.initialize().catch(error => {
  Logger.error('バックグラウンドサービス初期化失敗', error);
});

// グローバルエクスポート（デバッグ用）
globalThis.backgroundService = backgroundService;
