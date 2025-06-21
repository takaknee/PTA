/*
 * PTA AI業務支援ツール - リファクタリング版バックグラウンドサービスワーカー
 * Copyright (c) 2024 PTA Development Team
 */

// 新しいコアモジュールをインポート
import { APP_CONSTANTS } from '../core/constants.js';
import { logger, logInfo, logError } from '../core/logger.js';
import { handleError, PTAError } from '../core/error-handler.js';
import { eventManager, EVENTS } from '../core/event-manager.js';
import { settingsManager } from '../core/settings-manager.js';
import { apiClient } from '../core/api-client.js';

/**
 * サービスワーカーのライフサイクル管理クラス
 */
class PTABackgroundService {
    constructor() {
        this.isInitialized = false;
        this.contextMenus = new Map();
        this.notifications = new Map();
        this.activeAnalyses = new Map();
    }

    /**
     * サービスワーカーを初期化
     */
    async initialize() {
        try {
            logInfo('PTA AI業務支援ツール バックグラウンドサービス開始');

            // 設定管理システムの初期化
            await settingsManager.initialize();

            // イベントリスナーの設定
            this._setupEventListeners();

            // コンテキストメニューの設定
            await this._setupContextMenus();

            // 通知システムの初期化
            this._setupNotificationSystem();

            // 拡張機能の状態監視
            this._startHealthCheck();

            this.isInitialized = true;
            logInfo('バックグラウンドサービスの初期化が完了しました');

            // 初期化完了イベント
            eventManager.emit(EVENTS.APP_READY);

        } catch (error) {
            logError('バックグラウンドサービスの初期化中にエラーが発生しました', error);
            throw handleError(error, 'バックグラウンドサービス初期化');
        }
    }

    /**
     * イベントリスナーを設定
     */
    _setupEventListeners() {
        // Chrome拡張機能のメッセージリスナー
        chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
            this._handleMessage(message, sender, sendResponse);
            return true; // 非同期レスポンスを有効化
        });

        // 拡張機能インストール・更新時の処理
        chrome.runtime.onInstalled.addListener((details) => {
            this._handleInstallUpdate(details);
        });

        // アラーム処理
        chrome.alarms.onAlarm.addListener((alarm) => {
            this._handleAlarm(alarm);
        });

        // 設定変更時の処理
        eventManager.on(EVENTS.SETTINGS_CHANGED, (data) => {
            this._handleSettingsChange(data);
        });

        // API エラー時の処理
        eventManager.on(EVENTS.API_REQUEST_ERROR, (data) => {
            this._handleAPIError(data);
        });

        logInfo('イベントリスナーを設定しました');
    }

    /**
     * メッセージハンドラー
     */
    async _handleMessage(message, sender, sendResponse) {
        try {
            const { action, data } = message;
            logInfo(`メッセージ受信: ${action}`, { tabId: sender.tab?.id });

            let response;

            switch (action) {
                case 'analyzeEmail':
                    response = await this._analyzeEmail(data, sender.tab);
                    break;

                case 'analyzePage':
                    response = await this._analyzePage(data, sender.tab);
                    break;

                case 'analyzeSelection':
                    response = await this._analyzeSelection(data, sender.tab);
                    break;

                case 'forwardToTeams':
                    response = await this._forwardToTeams(data);
                    break;

                case 'addToCalendar':
                    response = await this._addToCalendar(data);
                    break;

                case 'testAPIConnection':
                    response = await this._testAPIConnection();
                    break;

                case 'getSettings':
                    response = this._getSettings(data.category);
                    break;

                case 'updateSettings':
                    response = await this._updateSettings(data.category, data.settings);
                    break;

                case 'showNotification':
                    response = await this._showNotification(data);
                    break;

                case 'reportCriticalError':
                    response = await this._reportCriticalError(data);
                    break;

                default:
                    throw new PTAError(`未知のアクション: ${action}`, 'UNKNOWN_ACTION');
            }

            sendResponse({ success: true, data: response });

        } catch (error) {
            const errorInfo = handleError(error, `メッセージ処理: ${message.action}`);
            sendResponse({
                success: false,
                error: errorInfo.userMessage,
                details: errorInfo
            });
        }
    }

    /**
     * メール分析処理
     */
    async _analyzeEmail(emailData, tab) {
        const analysisId = this._generateAnalysisId();

        try {
            this.activeAnalyses.set(analysisId, {
                type: 'email',
                startTime: Date.now(),
                tabId: tab?.id,
            });

            logInfo('メール分析を開始します', { analysisId, emailLength: emailData.body?.length });

            // 入力検証
            this._validateEmailData(emailData);

            // AI分析実行
            const analysisResult = await apiClient.analyzeContent(
                emailData.body,
                'email',
                {
                    subject: emailData.subject,
                    from: emailData.from,
                    date: emailData.date,
                    tone: settingsManager.get('userPreferences', 'defaultTone', 'polite'),
                }
            );

            // 結果を処理
            const processedResult = this._processAnalysisResult(analysisResult, 'email');

            // 履歴に保存
            await this._saveAnalysisHistory('email', emailData, processedResult);

            logInfo('メール分析が完了しました', { analysisId });

            eventManager.emit(EVENTS.CONTENT_ANALYZED, {
                type: 'email',
                analysisId,
                result: processedResult,
            });

            return processedResult;

        } catch (error) {
            logError('メール分析中にエラーが発生しました', { analysisId, error });
            throw error;
        } finally {
            this.activeAnalyses.delete(analysisId);
        }
    }

    /**
     * ページ分析処理
     */
    async _analyzePage(pageData, tab) {
        const analysisId = this._generateAnalysisId();

        try {
            this.activeAnalyses.set(analysisId, {
                type: 'page',
                startTime: Date.now(),
                tabId: tab?.id,
            });

            logInfo('ページ分析を開始します', { analysisId, url: tab?.url });

            // 入力検証
            this._validatePageData(pageData);

            // AI分析実行
            const analysisResult = await apiClient.analyzeContent(
                pageData.content,
                'summary',
                {
                    title: pageData.title,
                    url: pageData.url,
                }
            );

            // 結果を処理
            const processedResult = this._processAnalysisResult(analysisResult, 'page');

            // 履歴に保存
            await this._saveAnalysisHistory('page', pageData, processedResult);

            logInfo('ページ分析が完了しました', { analysisId });

            return processedResult;

        } catch (error) {
            logError('ページ分析中にエラーが発生しました', { analysisId, error });
            throw error;
        } finally {
            this.activeAnalyses.delete(analysisId);
        }
    }

    /**
     * 選択テキスト分析処理
     */
    async _analyzeSelection(selectionData, tab) {
        const analysisId = this._generateAnalysisId();

        try {
            this.activeAnalyses.set(analysisId, {
                type: 'selection',
                startTime: Date.now(),
                tabId: tab?.id,
            });

            logInfo('選択テキスト分析を開始します', { analysisId });

            // 入力検証
            this._validateSelectionData(selectionData);

            // AI分析実行
            const analysisResult = await apiClient.analyzeContent(
                selectionData.text,
                'general',
                {
                    context: selectionData.context,
                }
            );

            // 結果を処理
            const processedResult = this._processAnalysisResult(analysisResult, 'selection');

            logInfo('選択テキスト分析が完了しました', { analysisId });

            return processedResult;

        } catch (error) {
            logError('選択テキスト分析中にエラーが発生しました', { analysisId, error });
            throw error;
        } finally {
            this.activeAnalyses.delete(analysisId);
        }
    }

    /**
     * API接続テスト
     */
    async _testAPIConnection() {
        try {
            logInfo('API接続テストを開始します');
            const result = await apiClient.testConnection();
            logInfo('API接続テストが完了しました', result);
            return result;
        } catch (error) {
            logError('API接続テスト中にエラーが発生しました', error);
            throw error;
        }
    }

    /**
     * コンテキストメニューを設定
     */
    async _setupContextMenus() {
        try {
            // 既存のメニューをクリア
            await chrome.contextMenus.removeAll();

            const menuItems = [
                {
                    id: 'pta-analyze-page',
                    title: 'このページを分析',
                    contexts: ['page'],
                },
                {
                    id: 'pta-analyze-selection',
                    title: '選択テキストを分析',
                    contexts: ['selection'],
                },
                {
                    id: 'pta-separator-1',
                    type: 'separator',
                    contexts: ['page', 'selection'],
                },
                {
                    id: 'pta-quick-reply',
                    title: '返信を作成',
                    contexts: ['selection'],
                },
                {
                    id: 'pta-summarize',
                    title: '要約を作成',
                    contexts: ['selection'],
                },
                {
                    id: 'pta-separator-2',
                    type: 'separator',
                    contexts: ['page', 'selection'],
                },
                {
                    id: 'pta-open-settings',
                    title: '設定を開く',
                    contexts: ['page'],
                },
            ];

            for (const item of menuItems) {
                await chrome.contextMenus.create(item);
                this.contextMenus.set(item.id, item);
            }

            // コンテキストメニューのクリックリスナー
            chrome.contextMenus.onClicked.addListener((info, tab) => {
                this._handleContextMenuClick(info, tab);
            });

            logInfo('コンテキストメニューを設定しました', { count: menuItems.length });

        } catch (error) {
            logError('コンテキストメニューの設定中にエラーが発生しました', error);
        }
    }

    /**
     * コンテキストメニューのクリック処理
     */
    async _handleContextMenuClick(info, tab) {
        try {
            logInfo('コンテキストメニューがクリックされました', { menuItemId: info.menuItemId });

            switch (info.menuItemId) {
                case 'pta-analyze-page':
                    await this._sendToContentScript(tab.id, {
                        action: 'analyzePage',
                        data: { url: tab.url, title: tab.title }
                    });
                    break;

                case 'pta-analyze-selection':
                    await this._sendToContentScript(tab.id, {
                        action: 'analyzeSelection',
                        data: { text: info.selectionText }
                    });
                    break;

                case 'pta-quick-reply':
                    await this._sendToContentScript(tab.id, {
                        action: 'createQuickReply',
                        data: { text: info.selectionText }
                    });
                    break;

                case 'pta-summarize':
                    await this._sendToContentScript(tab.id, {
                        action: 'createSummary',
                        data: { text: info.selectionText }
                    });
                    break;

                case 'pta-open-settings':
                    await chrome.runtime.openOptionsPage();
                    break;
            }

        } catch (error) {
            handleError(error, 'コンテキストメニュー処理');
        }
    }

    /**
     * コンテンツスクリプトにメッセージを送信
     */
    async _sendToContentScript(tabId, message) {
        try {
            await chrome.tabs.sendMessage(tabId, message);
        } catch (error) {
            logError('コンテンツスクリプトへのメッセージ送信に失敗しました', { tabId, error });
        }
    }

    /**
     * 通知システムを設定
     */
    _setupNotificationSystem() {
        // 通知クリック処理
        chrome.notifications.onClicked.addListener((notificationId) => {
            this._handleNotificationClick(notificationId);
        });

        // 通知ボタンクリック処理
        chrome.notifications.onButtonClicked.addListener((notificationId, buttonIndex) => {
            this._handleNotificationButtonClick(notificationId, buttonIndex);
        });

        logInfo('通知システムを設定しました');
    }

    /**
     * 通知を表示
     */
    async _showNotification(notificationData) {
        try {
            const notificationId = `pta-${Date.now()}`;

            const options = {
                type: 'basic',
                iconUrl: '../assets/icons/icon128.png',
                title: notificationData.title || 'PTA AI業務支援ツール',
                message: notificationData.message,
                buttons: notificationData.buttons || [],
                ...notificationData.options,
            };

            await chrome.notifications.create(notificationId, options);

            this.notifications.set(notificationId, {
                ...notificationData,
                createdAt: Date.now(),
            });

            // 自動削除タイマー
            if (notificationData.timeout) {
                setTimeout(() => {
                    chrome.notifications.clear(notificationId);
                    this.notifications.delete(notificationId);
                }, notificationData.timeout);
            }

            logInfo('通知を表示しました', { notificationId });
            return notificationId;

        } catch (error) {
            logError('通知の表示中にエラーが発生しました', error);
            throw error;
        }
    }

    /**
     * ヘルスチェックを開始
     */
    _startHealthCheck() {
        const healthCheckInterval = 60000; // 1分間隔

        chrome.alarms.create('pta-health-check', {
            delayInMinutes: 1,
            periodInMinutes: 1,
        });

        logInfo('ヘルスチェックを開始しました', { interval: healthCheckInterval });
    }

    /**
     * アラーム処理
     */
    _handleAlarm(alarm) {
        switch (alarm.name) {
            case 'pta-health-check':
                this._performHealthCheck();
                break;
            default:
                logInfo('未知のアラーム', { name: alarm.name });
        }
    }

    /**
     * ヘルスチェック実行
     */
    _performHealthCheck() {
        const stats = {
            activeAnalyses: this.activeAnalyses.size,
            notifications: this.notifications.size,
            apiStats: apiClient.getStats(),
            memoryUsage: this._getMemoryUsage(),
        };

        logInfo('ヘルスチェック実行', stats);

        // 長時間実行中の分析をクリーンアップ
        const maxAnalysisTime = 300000; // 5分
        const now = Date.now();

        this.activeAnalyses.forEach((analysis, id) => {
            if (now - analysis.startTime > maxAnalysisTime) {
                logError('長時間実行中の分析をキャンセルします', { analysisId: id, duration: now - analysis.startTime });
                this.activeAnalyses.delete(id);
            }
        });
    }

    /**
     * メモリ使用量を取得
     */
    _getMemoryUsage() {
        if (typeof performance !== 'undefined' && performance.memory) {
            return {
                used: Math.round(performance.memory.usedJSHeapSize / 1024 / 1024),
                total: Math.round(performance.memory.totalJSHeapSize / 1024 / 1024),
                limit: Math.round(performance.memory.jsHeapSizeLimit / 1024 / 1024),
            };
        }
        return null;
    }

    /**
     * 分析IDを生成
     */
    _generateAnalysisId() {
        return `analysis_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }

    /**
     * 入力データを検証
     */
    _validateEmailData(emailData) {
        if (!emailData || !emailData.body) {
            throw new PTAError('メールデータが不正です', 'VALIDATION_ERROR');
        }
        if (emailData.body.length > APP_CONSTANTS.CONTENT_LIMITS.MAX_TEXT_LENGTH) {
            throw new PTAError(APP_CONSTANTS.ERROR_MESSAGES.CONTENT_TOO_LARGE, 'VALIDATION_ERROR');
        }
    }

    _validatePageData(pageData) {
        if (!pageData || !pageData.content) {
            throw new PTAError('ページデータが不正です', 'VALIDATION_ERROR');
        }
        if (pageData.content.length > APP_CONSTANTS.CONTENT_LIMITS.MAX_HTML_LENGTH) {
            throw new PTAError(APP_CONSTANTS.ERROR_MESSAGES.CONTENT_TOO_LARGE, 'VALIDATION_ERROR');
        }
    }

    _validateSelectionData(selectionData) {
        if (!selectionData || !selectionData.text) {
            throw new PTAError('選択データが不正です', 'VALIDATION_ERROR');
        }
        if (selectionData.text.length > APP_CONSTANTS.CONTENT_LIMITS.MAX_SELECTION_LENGTH) {
            throw new PTAError(APP_CONSTANTS.ERROR_MESSAGES.CONTENT_TOO_LARGE, 'VALIDATION_ERROR');
        }
    }

    /**
     * 分析結果を処理
     */
    _processAnalysisResult(result, type) {
        try {
            let content;

            if (result.data?.choices?.[0]?.message?.content) {
                content = result.data.choices[0].message.content;
            } else if (typeof result.data === 'string') {
                content = result.data;
            } else {
                throw new PTAError('AIレスポンスの形式が不正です', 'API_ERROR');
            }

            return {
                type,
                content,
                timestamp: new Date().toISOString(),
                model: result.data?.model || 'unknown',
                usage: result.data?.usage || null,
            };

        } catch (error) {
            logError('分析結果の処理中にエラーが発生しました', error);
            throw error;
        }
    }

    /**
     * 分析履歴を保存
     */
    async _saveAnalysisHistory(type, inputData, result) {
        try {
            if (!settingsManager.get('analysisHistory', 'autoSave', true)) {
                return;
            }

            const historyEntry = {
                id: this._generateAnalysisId(),
                type,
                timestamp: new Date().toISOString(),
                input: {
                    // 機密情報を除外した入力データのサマリー
                    length: typeof inputData === 'string' ? inputData.length : JSON.stringify(inputData).length,
                    type: typeof inputData,
                },
                result: {
                    type: result.type,
                    length: result.content.length,
                    model: result.model,
                },
            };

            // 履歴をストレージに保存（実装は簡略化）
            logInfo('分析履歴を保存しました', { entryId: historyEntry.id });

        } catch (error) {
            logError('分析履歴の保存中にエラーが発生しました', error);
        }
    }

    /**
     * 設定を取得
     */
    _getSettings(category) {
        return settingsManager.get(category);
    }

    /**
     * 設定を更新
     */
    async _updateSettings(category, settings) {
        await settingsManager.set(category, settings);
        return { success: true };
    }

    /**
     * インストール・更新処理
     */
    _handleInstallUpdate(details) {
        switch (details.reason) {
            case 'install':
                logInfo('拡張機能がインストールされました');
                this._showWelcomeNotification();
                break;
            case 'update':
                logInfo('拡張機能が更新されました', {
                    previousVersion: details.previousVersion,
                    currentVersion: chrome.runtime.getManifest().version,
                });
                break;
        }
    }

    /**
     * ウェルカム通知を表示
     */
    async _showWelcomeNotification() {
        await this._showNotification({
            title: 'PTA AI業務支援ツール',
            message: 'インストールありがとうございます。設定画面でAPIキーを設定してください。',
            buttons: [
                { title: '設定を開く' },
                { title: '後で設定' },
            ],
            timeout: 10000,
        });
    }

    /**
     * 設定変更時の処理
     */
    _handleSettingsChange(data) {
        logInfo('設定が変更されました', { category: data.category, key: data.key });

        // 特定の設定変更に対する処理
        if (data.category === 'userPreferences' && data.key === 'debugMode') {
            eventManager.setDebugMode(data.value);
        }
    }

    /**
     * APIエラー処理
     */
    _handleAPIError(data) {
        logError('API呼び出しでエラーが発生しました', data);

        // 重要なエラーの場合は通知
        if (data.error.includes('認証') || data.error.includes('401')) {
            this._showNotification({
                title: 'API認証エラー',
                message: 'APIキーを確認してください。設定画面から再設定できます。',
                buttons: [{ title: '設定を開く' }],
            });
        }
    }

    /**
     * 重要エラーの報告
     */
    async _reportCriticalError(errorData) {
        logError('重要エラーが報告されました', errorData);

        // 必要に応じて外部サービスに報告
        // この実装では ログ記録のみ

        return { reported: true, timestamp: new Date().toISOString() };
    }

    /**
     * 通知クリック処理
     */
    _handleNotificationClick(notificationId) {
        const notification = this.notifications.get(notificationId);
        if (notification?.action === 'openSettings') {
            chrome.runtime.openOptionsPage();
        }
        this.notifications.delete(notificationId);
    }

    /**
     * 通知ボタンクリック処理
     */
    _handleNotificationButtonClick(notificationId, buttonIndex) {
        const notification = this.notifications.get(notificationId);
        if (notification?.buttons?.[buttonIndex]?.title === '設定を開く') {
            chrome.runtime.openOptionsPage();
        }
        this.notifications.delete(notificationId);
    }
}

// グローバルサービスインスタンス
const backgroundService = new PTABackgroundService();

// サービスワーカーの開始
backgroundService.initialize().catch(error => {
    console.error('バックグラウンドサービスの初期化に失敗しました:', error);
});

// サービスワーカーのアクティブ保持
chrome.runtime.onStartup.addListener(() => {
    logInfo('サービスワーカーが起動しました');
});

// エクスポート（テスト用）
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { PTABackgroundService, backgroundService };
}
