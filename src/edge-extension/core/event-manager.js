/*
 * PTA AI業務支援ツール - イベント管理モジュール
 * Copyright (c) 2024 PTA Development Team
 */

import { logger } from './logger.js';
import { handleError } from './error-handler.js';

/**
 * 統一イベント管理クラス
 * アプリケーション全体でのイベント通信を管理
 */
export class PTAEventManager {
    constructor() {
        this.listeners = new Map();
        this.eventHistory = [];
        this.maxHistorySize = 1000;
        this.debugMode = false;
    }

    /**
     * イベントリスナーを登録
     */
    on(eventName, callback, options = {}) {
        if (!this.listeners.has(eventName)) {
            this.listeners.set(eventName, []);
        }

        const listener = {
            callback,
            once: options.once || false,
            priority: options.priority || 0,
            context: options.context || '不明',
            id: this._generateListenerId(),
        };

        this.listeners.get(eventName).push(listener);

        // 優先度でソート（高い優先度が先に実行）
        this.listeners.get(eventName).sort((a, b) => b.priority - a.priority);

        if (this.debugMode) {
            logger.debug(`イベントリスナー登録: ${eventName}`, {
                context: listener.context,
                priority: listener.priority,
                once: listener.once,
            });
        }

        return listener.id;
    }

    /**
     * 一度だけ実行されるイベントリスナーを登録
     */
    once(eventName, callback, options = {}) {
        return this.on(eventName, callback, { ...options, once: true });
    }

    /**
     * イベントリスナーを削除
     */
    off(eventName, listenerId) {
        const listeners = this.listeners.get(eventName);
        if (!listeners) return false;

        const index = listeners.findIndex(listener => listener.id === listenerId);
        if (index === -1) return false;

        listeners.splice(index, 1);

        if (listeners.length === 0) {
            this.listeners.delete(eventName);
        }

        if (this.debugMode) {
            logger.debug(`イベントリスナー削除: ${eventName}`, { listenerId });
        }

        return true;
    }

    /**
     * 指定したイベントのすべてのリスナーを削除
     */
    removeAllListeners(eventName) {
        if (eventName) {
            const count = this.listeners.get(eventName)?.length || 0;
            this.listeners.delete(eventName);

            if (this.debugMode) {
                logger.debug(`イベントリスナー一括削除: ${eventName}`, { count });
            }

            return count;
        } else {
            const totalCount = Array.from(this.listeners.values())
                .reduce((sum, listeners) => sum + listeners.length, 0);
            this.listeners.clear();

            if (this.debugMode) {
                logger.debug('全イベントリスナー削除', { totalCount });
            }

            return totalCount;
        }
    }

    /**
     * イベントを発火
     */
    async emit(eventName, data = null, options = {}) {
        const eventInfo = {
            name: eventName,
            data,
            timestamp: new Date().toISOString(),
            async: options.async !== false,
            source: options.source || 'unknown',
        };

        // 履歴に追加
        this._addToHistory(eventInfo);

        const listeners = this.listeners.get(eventName);
        if (!listeners || listeners.length === 0) {
            if (this.debugMode) {
                logger.debug(`イベント発火（リスナーなし）: ${eventName}`, eventInfo);
            }
            return [];
        }

        if (this.debugMode) {
            logger.debug(`イベント発火: ${eventName}`, {
                ...eventInfo,
                listenerCount: listeners.length,
            });
        }

        const results = [];
        const toRemove = [];

        for (const listener of listeners) {
            try {
                let result;

                if (eventInfo.async) {
                    // 非同期実行
                    result = await this._executeListener(listener, eventInfo);
                } else {
                    // 同期実行
                    result = this._executeListener(listener, eventInfo);
                }

                results.push({
                    listenerId: listener.id,
                    context: listener.context,
                    result,
                    success: true,
                });

                // once オプションが設定されている場合は削除対象に追加
                if (listener.once) {
                    toRemove.push(listener.id);
                }

            } catch (error) {
                const errorResult = {
                    listenerId: listener.id,
                    context: listener.context,
                    error: error.message,
                    success: false,
                };

                results.push(errorResult);

                // エラーログ
                handleError(error, `イベントリスナー実行: ${eventName}`, false);

                // once オプションが設定されている場合はエラーでも削除
                if (listener.once) {
                    toRemove.push(listener.id);
                }
            }
        }

        // once リスナーを削除
        toRemove.forEach(listenerId => {
            this.off(eventName, listenerId);
        });

        return results;
    }

    /**
     * リスナーを実行
     */
    _executeListener(listener, eventInfo) {
        const startTime = performance.now();

        try {
            const result = listener.callback(eventInfo.data, eventInfo);

            if (this.debugMode) {
                const duration = performance.now() - startTime;
                logger.debug(`リスナー実行完了: ${listener.context}`, {
                    duration: `${duration.toFixed(2)}ms`,
                    eventName: eventInfo.name,
                });
            }

            return result;
        } catch (error) {
            const duration = performance.now() - startTime;
            logger.error(`リスナー実行エラー: ${listener.context}`, {
                duration: `${duration.toFixed(2)}ms`,
                eventName: eventInfo.name,
                error: error.message,
            });
            throw error;
        }
    }

    /**
     * 条件に一致するまでイベントを待機
     */
    waitFor(eventName, condition = null, timeout = 30000) {
        return new Promise((resolve, reject) => {
            const timeoutId = setTimeout(() => {
                this.off(eventName, listenerId);
                reject(new Error(`イベント待機タイムアウト: ${eventName} (${timeout}ms)`));
            }, timeout);

            const listenerId = this.once(eventName, (data, eventInfo) => {
                try {
                    if (!condition || condition(data, eventInfo)) {
                        clearTimeout(timeoutId);
                        resolve({ data, eventInfo });
                    } else {
                        // 条件に一致しない場合は再度待機
                        const newListenerId = this.waitFor(eventName, condition, timeout - (Date.now() - new Date(eventInfo.timestamp).getTime()));
                        newListenerId.then(resolve).catch(reject);
                    }
                } catch (error) {
                    clearTimeout(timeoutId);
                    reject(error);
                }
            });
        });
    }

    /**
     * リスナーIDを生成
     */
    _generateListenerId() {
        return `listener_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }

    /**
     * イベント履歴に追加
     */
    _addToHistory(eventInfo) {
        this.eventHistory.push(eventInfo);

        if (this.eventHistory.length > this.maxHistorySize) {
            this.eventHistory.shift();
        }
    }

    /**
     * イベント統計を取得
     */
    getStats() {
        const stats = {
            totalEvents: this.eventHistory.length,
            totalListeners: Array.from(this.listeners.values())
                .reduce((sum, listeners) => sum + listeners.length, 0),
            eventsByName: {},
            recentEvents: this.eventHistory.slice(-20),
            listenersByEvent: {},
        };

        // イベント名別の統計
        this.eventHistory.forEach(event => {
            stats.eventsByName[event.name] = (stats.eventsByName[event.name] || 0) + 1;
        });

        // リスナー数の統計
        this.listeners.forEach((listeners, eventName) => {
            stats.listenersByEvent[eventName] = listeners.length;
        });

        return stats;
    }

    /**
     * デバッグモードの切り替え
     */
    setDebugMode(enabled) {
        this.debugMode = enabled;
        logger.info(`イベントマネージャーデバッグモード: ${enabled ? '有効' : '無効'}`);
    }

    /**
     * 履歴をクリア
     */
    clearHistory() {
        this.eventHistory = [];
        logger.info('イベント履歴をクリアしました');
    }
}

/**
 * グローバルイベントマネージャーインスタンス
 */
export const eventManager = new PTAEventManager();

/**
 * 便利な関数群
 */
export const on = (eventName, callback, options) => eventManager.on(eventName, callback, options);
export const once = (eventName, callback, options) => eventManager.once(eventName, callback, options);
export const off = (eventName, listenerId) => eventManager.off(eventName, listenerId);
export const emit = (eventName, data, options) => eventManager.emit(eventName, data, options);
export const waitFor = (eventName, condition, timeout) => eventManager.waitFor(eventName, condition, timeout);

// よく使用されるイベント名の定数
export const EVENTS = {
    // アプリケーションライフサイクル
    APP_READY: 'app:ready',
    APP_ERROR: 'app:error',

    // 設定関連
    SETTINGS_CHANGED: 'settings:changed',
    SETTINGS_SAVED: 'settings:saved',
    SETTINGS_LOADED: 'settings:loaded',

    // API関連
    API_REQUEST_START: 'api:request:start',
    API_REQUEST_SUCCESS: 'api:request:success',
    API_REQUEST_ERROR: 'api:request:error',

    // UI関連
    UI_THEME_CHANGED: 'ui:theme:changed',
    UI_NOTIFICATION_SHOW: 'ui:notification:show',
    UI_MODAL_OPEN: 'ui:modal:open',
    UI_MODAL_CLOSE: 'ui:modal:close',

    // コンテンツ関連
    CONTENT_ANALYZED: 'content:analyzed',
    CONTENT_SELECTED: 'content:selected',
    CONTENT_COPIED: 'content:copied',

    // セキュリティ関連
    SECURITY_VIOLATION: 'security:violation',
    SECURITY_WARNING: 'security:warning',
};
