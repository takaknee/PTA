/*
 * PTA AI業務支援ツール - エラーハンドリングモジュール
 * Copyright (c) 2024 PTA Development Team
 */

import { APP_CONSTANTS } from './constants.js';
import { logger } from './logger.js';

/**
 * PTA固有のエラークラス
 */
export class PTAError extends Error {
    constructor(message, type = 'GENERAL', details = null) {
        super(message);
        this.name = 'PTAError';
        this.type = type;
        this.details = details;
        this.timestamp = new Date().toISOString();
    }
}

/**
 * API関連のエラークラス
 */
export class PTAAPIError extends PTAError {
    constructor(message, statusCode = null, response = null) {
        super(message, 'API_ERROR');
        this.statusCode = statusCode;
        this.response = response;
    }
}

/**
 * セキュリティ関連のエラークラス
 */
export class PTASecurityError extends PTAError {
    constructor(message, securityType = 'GENERAL', details = null) {
        super(message, 'SECURITY_ERROR');
        this.securityType = securityType;
        this.details = details;
    }
}

/**
 * バリデーション関連のエラークラス
 */
export class PTAValidationError extends PTAError {
    constructor(message, field = null, value = null) {
        super(message, 'VALIDATION_ERROR');
        this.field = field;
        this.value = value;
    }
}

/**
 * 統一エラーハンドラークラス
 */
export class PTAErrorHandler {
    constructor() {
        this.errorHistory = [];
        this.maxHistorySize = 100;
        this.retryAttempts = new Map();
    }

    /**
     * エラーを処理
     */
    handle(error, context = '不明', showToUser = true) {
        const errorInfo = this._analyzeError(error, context);

        // ログに記録
        logger.error(errorInfo.message, {
            type: errorInfo.type,
            context: errorInfo.context,
            details: errorInfo.details,
            stack: errorInfo.stack,
        });

        // 履歴に追加
        this._addToHistory(errorInfo);

        // ユーザーに表示（必要に応じて）
        if (showToUser) {
            this._showUserMessage(errorInfo);
        }

        // 重要なエラーの場合は通知
        if (this._isCriticalError(errorInfo)) {
            this._notifyCriticalError(errorInfo);
        }

        return errorInfo;
    }

    /**
     * エラーを分析
     */
    _analyzeError(error, context) {
        const timestamp = new Date().toISOString();
        let type, message, details, userMessage, stack;

        if (error instanceof PTAError) {
            type = error.type;
            message = error.message;
            details = error.details;
            stack = error.stack;
        } else if (error instanceof Error) {
            type = 'JAVASCRIPT_ERROR';
            message = error.message;
            stack = error.stack;
        } else if (typeof error === 'string') {
            type = 'STRING_ERROR';
            message = error;
        } else {
            type = 'UNKNOWN_ERROR';
            message = 'Unknown error occurred';
            details = error;
        }

        // ユーザー向けメッセージを生成
        userMessage = this._generateUserMessage(type, message);

        return {
            timestamp,
            type,
            message,
            details,
            userMessage,
            context,
            stack,
        };
    }

    /**
     * ユーザー向けメッセージを生成
     */
    _generateUserMessage(type, originalMessage) {
        switch (type) {
            case 'API_ERROR':
                if (originalMessage.includes('401') || originalMessage.includes('認証')) {
                    return APP_CONSTANTS.ERROR_MESSAGES.API_KEY_MISSING;
                }
                if (originalMessage.includes('timeout') || originalMessage.includes('タイムアウト')) {
                    return APP_CONSTANTS.ERROR_MESSAGES.TIMEOUT_ERROR;
                }
                return APP_CONSTANTS.ERROR_MESSAGES.API_CONNECTION_FAILED;

            case 'SECURITY_ERROR':
                return 'セキュリティ上の問題が検出されました。管理者にお問い合わせください。';

            case 'VALIDATION_ERROR':
                if (originalMessage.includes('large') || originalMessage.includes('大きすぎ')) {
                    return APP_CONSTANTS.ERROR_MESSAGES.CONTENT_TOO_LARGE;
                }
                return APP_CONSTANTS.ERROR_MESSAGES.INVALID_FORMAT;

            case 'PERMISSION_ERROR':
                return APP_CONSTANTS.ERROR_MESSAGES.PERMISSION_DENIED;

            default:
                return originalMessage.length > 100 ?
                    APP_CONSTANTS.ERROR_MESSAGES.UNKNOWN_ERROR :
                    originalMessage;
        }
    }

    /**
     * ユーザーにメッセージを表示
     */
    _showUserMessage(errorInfo) {
        if (typeof chrome !== 'undefined' && chrome.runtime) {
            // バックグラウンドスクリプトに通知依頼
            chrome.runtime.sendMessage({
                action: 'showNotification',
                data: {
                    type: 'error',
                    title: 'エラーが発生しました',
                    message: errorInfo.userMessage,
                    timeout: APP_CONSTANTS.UI.NOTIFICATION_TIMEOUT,
                }
            }).catch(() => {
                // フォールバック: コンソールに出力
                console.error('エラー:', errorInfo.userMessage);
            });
        } else {
            // Webページの場合はアラート表示
            console.error('エラー:', errorInfo.userMessage);
            if (typeof alert !== 'undefined') {
                alert(`エラー: ${errorInfo.userMessage}`);
            }
        }
    }

    /**
     * 重要なエラーかどうかを判定
     */
    _isCriticalError(errorInfo) {
        const criticalTypes = ['SECURITY_ERROR', 'API_ERROR'];
        const criticalKeywords = ['認証', 'セキュリティ', '権限', 'データ破損'];

        return criticalTypes.includes(errorInfo.type) ||
            criticalKeywords.some(keyword => errorInfo.message.includes(keyword));
    }

    /**
     * 重要なエラーの通知
     */
    _notifyCriticalError(errorInfo) {
        try {
            if (typeof chrome !== 'undefined' && chrome.runtime) {
                chrome.runtime.sendMessage({
                    action: 'reportCriticalError',
                    data: errorInfo
                });
            }
        } catch (error) {
            console.error('重要エラーの通知に失敗:', error);
        }
    }

    /**
     * エラー履歴に追加
     */
    _addToHistory(errorInfo) {
        this.errorHistory.push(errorInfo);

        // 履歴サイズ制限
        if (this.errorHistory.length > this.maxHistorySize) {
            this.errorHistory.shift();
        }
    }

    /**
     * リトライ処理
     */
    async retry(operation, maxAttempts = APP_CONSTANTS.API.MAX_RETRIES, delay = APP_CONSTANTS.API.RETRY_DELAY) {
        const operationId = operation.name || 'anonymous';
        let attempts = this.retryAttempts.get(operationId) || 0;

        while (attempts < maxAttempts) {
            try {
                const result = await operation();
                this.retryAttempts.delete(operationId);
                return result;
            } catch (error) {
                attempts++;
                this.retryAttempts.set(operationId, attempts);

                if (attempts >= maxAttempts) {
                    throw new PTAError(
                        `${maxAttempts}回の試行後も処理に失敗しました: ${error.message}`,
                        'RETRY_EXHAUSTED',
                        { originalError: error, attempts }
                    );
                }

                logger.warn(`処理失敗、リトライします (${attempts}/${maxAttempts})`, {
                    operation: operationId,
                    error: error.message,
                    delay
                });

                await this._delay(delay * attempts); // 指数バックオフ
            }
        }
    }

    /**
     * 遅延処理
     */
    _delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * エラー履歴を取得
     */
    getErrorHistory(limit = 50) {
        return this.errorHistory.slice(-limit);
    }

    /**
     * エラー統計を取得
     */
    getErrorStats() {
        const stats = {
            total: this.errorHistory.length,
            byType: {},
            byContext: {},
            recent: this.errorHistory.slice(-10),
        };

        this.errorHistory.forEach(error => {
            stats.byType[error.type] = (stats.byType[error.type] || 0) + 1;
            stats.byContext[error.context] = (stats.byContext[error.context] || 0) + 1;
        });

        return stats;
    }

    /**
     * エラー履歴をクリア
     */
    clearHistory() {
        this.errorHistory = [];
        this.retryAttempts.clear();
        logger.info('エラー履歴をクリアしました');
    }
}

/**
 * グローバルエラーハンドラーインスタンス
 */
export const errorHandler = new PTAErrorHandler();

/**
 * 便利な関数群
 */
export const handleError = (error, context, showToUser = true) =>
    errorHandler.handle(error, context, showToUser);

export const retryOperation = (operation, maxAttempts, delay) =>
    errorHandler.retry(operation, maxAttempts, delay);

/**
 * 非同期処理のラッパー
 */
export function safeAsync(asyncFn, context = '非同期処理') {
    return async (...args) => {
        try {
            return await asyncFn(...args);
        } catch (error) {
            throw handleError(error, context, false);
        }
    };
}

/**
 * 同期処理のラッパー
 */
export function safeSync(syncFn, context = '同期処理') {
    return (...args) => {
        try {
            return syncFn(...args);
        } catch (error) {
            throw handleError(error, context, false);
        }
    };
}
