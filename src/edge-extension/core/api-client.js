/*
 * PTA AI業務支援ツール - APIクライアントモジュール
 * Copyright (c) 2024 PTA Development Team
 */

import { APP_CONSTANTS } from './constants.js';
import { logger } from './logger.js';
import { handleError, retryOperation, PTAAPIError } from './error-handler.js';
import { settingsManager } from './settings-manager.js';
import { eventManager, EVENTS } from './event-manager.js';

/**
 * APIレスポンス形式
 */
class APIResponse {
    constructor(data, status, headers = {}) {
        this.data = data;
        this.status = status;
        this.headers = headers;
        this.timestamp = new Date().toISOString();
        this.success = status >= 200 && status < 300;
    }
}

/**
 * APIリクエスト設定
 */
class APIRequestConfig {
    constructor(options = {}) {
        this.method = options.method || 'POST';
        this.headers = options.headers || {};
        this.timeout = options.timeout || settingsManager.get('apiSettings', 'timeout', APP_CONSTANTS.API.TIMEOUT);
        this.retries = options.retries || APP_CONSTANTS.API.MAX_RETRIES;
        this.validateResponse = options.validateResponse !== false;
        this.cacheEnabled = options.cacheEnabled !== false;
        this.rateLimitEnabled = options.rateLimitEnabled !== false;
    }
}

/**
 * レート制限管理クラス
 */
class RateLimiter {
    constructor() {
        this.requests = new Map();
        this.limits = {
            perMinute: 60,
            perHour: 3600,
            perDay: 86400,
        };
    }

    /**
     * リクエスト可能かチェック
     */
    canMakeRequest(endpoint = 'default') {
        const now = Date.now();
        const requestHistory = this.requests.get(endpoint) || [];

        // 古いリクエスト履歴を削除
        const oneMinuteAgo = now - 60000;
        const validRequests = requestHistory.filter(timestamp => timestamp > oneMinuteAgo);

        if (validRequests.length >= this.limits.perMinute) {
            return false;
        }

        this.requests.set(endpoint, validRequests);
        return true;
    }

    /**
     * リクエストを記録
     */
    recordRequest(endpoint = 'default') {
        const requestHistory = this.requests.get(endpoint) || [];
        requestHistory.push(Date.now());
        this.requests.set(endpoint, requestHistory);
    }

    /**
     * 次のリクエスト可能時刻を取得
     */
    getNextAvailableTime(endpoint = 'default') {
        const requestHistory = this.requests.get(endpoint) || [];
        if (requestHistory.length === 0) return 0;

        const oldestRequest = Math.min(...requestHistory);
        return Math.max(0, oldestRequest + 60000 - Date.now());
    }
}

/**
 * キャッシュ管理クラス
 */
class APICache {
    constructor() {
        this.cache = new Map();
        this.maxSize = settingsManager.get('cacheSettings', 'maxSize', 10485760); // 10MB
        this.ttl = settingsManager.get('cacheSettings', 'ttl', 3600000); // 1時間
        this.currentSize = 0;
    }

    /**
     * キャッシュキーを生成
     */
    _generateKey(url, body, headers) {
        const keyData = {
            url,
            body: typeof body === 'string' ? body : JSON.stringify(body),
            headers: JSON.stringify(headers),
        };

        // 簡易ハッシュ関数
        const str = JSON.stringify(keyData);
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // 32bit整数に変換
        }
        return `api_cache_${hash}`;
    }

    /**
     * データサイズを計算
     */
    _calculateSize(data) {
        return new Blob([JSON.stringify(data)]).size;
    }

    /**
     * キャッシュから取得
     */
    get(url, body, headers) {
        const key = this._generateKey(url, body, headers);
        const cached = this.cache.get(key);

        if (!cached) return null;

        // TTLチェック
        if (Date.now() - cached.timestamp > this.ttl) {
            this.cache.delete(key);
            this.currentSize -= cached.size;
            return null;
        }

        logger.debug('APIキャッシュヒット', { key, url });
        return cached.data;
    }

    /**
     * キャッシュに保存
     */
    set(url, body, headers, data) {
        if (!settingsManager.get('cacheSettings', 'enabled', true)) return;

        const key = this._generateKey(url, body, headers);
        const size = this._calculateSize(data);

        // サイズ制限チェック
        if (size > this.maxSize * 0.1) { // 全体の10%を超える場合はキャッシュしない
            return;
        }

        // 容量確保
        while (this.currentSize + size > this.maxSize && this.cache.size > 0) {
            const oldestKey = this.cache.keys().next().value;
            const oldest = this.cache.get(oldestKey);
            this.cache.delete(oldestKey);
            this.currentSize -= oldest.size;
        }

        this.cache.set(key, {
            data,
            timestamp: Date.now(),
            size,
        });

        this.currentSize += size;

        logger.debug('APIレスポンスをキャッシュしました', {
            key,
            size,
            totalSize: this.currentSize,
            cacheEntries: this.cache.size,
        });
    }

    /**
     * キャッシュをクリア
     */
    clear() {
        this.cache.clear();
        this.currentSize = 0;
        logger.info('APIキャッシュをクリアしました');
    }

    /**
     * キャッシュ統計を取得
     */
    getStats() {
        return {
            entries: this.cache.size,
            size: this.currentSize,
            maxSize: this.maxSize,
            utilization: (this.currentSize / this.maxSize) * 100,
        };
    }
}

/**
 * 統一APIクライアントクラス
 */
export class PTAAPIClient {
    constructor() {
        this.rateLimiter = new RateLimiter();
        this.cache = new APICache();
        this.activeRequests = new Map();
        this.requestId = 0;
    }

    /**
     * OpenAI APIに分析リクエストを送信
     */
    async analyzeContent(content, analysisType = 'general', options = {}) {
        const prompt = this._buildPrompt(content, analysisType, options);

        return await this.sendChatRequest([
            { role: 'system', content: prompt.system },
            { role: 'user', content: prompt.user },
        ], options);
    }

    /**
     * チャットリクエストを送信
     */
    async sendChatRequest(messages, options = {}) {
        const config = new APIRequestConfig(options);
        const apiSettings = settingsManager.get('apiSettings');

        if (!apiSettings.openaiApiKey || apiSettings.openaiApiKey === '') {
            throw new PTAAPIError(APP_CONSTANTS.ERROR_MESSAGES.API_KEY_MISSING);
        }

        const requestBody = {
            model: apiSettings.openaiModel,
            messages,
            max_tokens: apiSettings.maxTokens,
            temperature: apiSettings.temperature,
            ...options.additionalParams,
        };

        const requestConfig = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiSettings.openaiApiKey}`,
                ...config.headers,
            },
            body: JSON.stringify(requestBody),
            timeout: config.timeout,
        };

        return await this._makeRequest(apiSettings.openaiEndpoint, requestConfig);
    }

    /**
     * HTTPリクエストを実行
     */
    async _makeRequest(url, config) {
        const requestId = ++this.requestId;
        const startTime = performance.now();

        try {
            // イベント発火
            eventManager.emit(EVENTS.API_REQUEST_START, {
                requestId,
                url,
                method: config.method,
            });

            // キャッシュチェック
            if (config.method === 'GET' || config.cacheEnabled) {
                const cached = this.cache.get(url, config.body, config.headers);
                if (cached) {
                    return new APIResponse(cached, 200);
                }
            }

            // レート制限チェック
            if (config.rateLimitEnabled && !this.rateLimiter.canMakeRequest(url)) {
                const waitTime = this.rateLimiter.getNextAvailableTime(url);
                logger.warn(`APIレート制限、${waitTime}ms待機します`, { url, waitTime });
                await this._delay(waitTime);
            }

            // リクエスト実行
            const response = await this._executeRequest(url, config, requestId);

            // レート制限記録
            if (config.rateLimitEnabled) {
                this.rateLimiter.recordRequest(url);
            }

            // キャッシュ保存
            if ((config.method === 'GET' || config.cacheEnabled) && response.success) {
                this.cache.set(url, config.body, config.headers, response.data);
            }

            // イベント発火
            const duration = performance.now() - startTime;
            eventManager.emit(EVENTS.API_REQUEST_SUCCESS, {
                requestId,
                url,
                method: config.method,
                duration,
                status: response.status,
            });

            logger.apiCall(config.method, url, duration, true, response.status);

            return response;

        } catch (error) {
            const duration = performance.now() - startTime;

            // イベント発火
            eventManager.emit(EVENTS.API_REQUEST_ERROR, {
                requestId,
                url,
                method: config.method,
                duration,
                error: error.message,
            });

            logger.apiCall(config.method, url, duration, false, error.message);

            throw error;
        } finally {
            this.activeRequests.delete(requestId);
        }
    }

    /**
     * リクエストを実行（リトライ付き）
     */
    async _executeRequest(url, config, requestId) {
        return await retryOperation(async () => {
            this.activeRequests.set(requestId, { url, config, startTime: Date.now() });

            // AbortController でタイムアウト制御
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), config.timeout);

            try {
                const response = await fetch(url, {
                    ...config,
                    signal: controller.signal,
                });

                clearTimeout(timeoutId);

                // レスポンスを解析
                const responseData = await this._parseResponse(response);

                if (!response.ok) {
                    throw new PTAAPIError(
                        `HTTPエラー: ${response.status} ${response.statusText}`,
                        response.status,
                        responseData
                    );
                }

                return new APIResponse(responseData, response.status, Object.fromEntries(response.headers));

            } catch (error) {
                clearTimeout(timeoutId);

                if (error.name === 'AbortError') {
                    throw new PTAAPIError('リクエストがタイムアウトしました', 408);
                }

                if (error instanceof PTAAPIError) {
                    throw error;
                }

                throw new PTAAPIError(`ネットワークエラー: ${error.message}`, null, error);
            }
        }, config.retries);
    }

    /**
     * レスポンスを解析
     */
    async _parseResponse(response) {
        const contentType = response.headers.get('content-type');

        if (contentType?.includes('application/json')) {
            return await response.json();
        } else if (contentType?.includes('text/')) {
            return await response.text();
        } else {
            return await response.blob();
        }
    }

    /**
     * プロンプトを構築
     */
    _buildPrompt(content, analysisType, options) {
        const prompts = APP_CONSTANTS.PROMPTS;
        let template;

        switch (analysisType) {
            case 'email':
                template = prompts.EMAIL_ANALYSIS;
                break;
            case 'reply':
                template = prompts.QUICK_REPLY;
                break;
            case 'summary':
                template = prompts.PAGE_SUMMARY;
                break;
            default:
                template = prompts.EMAIL_ANALYSIS;
        }

        // テンプレート変数を置換
        let systemPrompt = template.SYSTEM;
        let userPrompt = template.USER_TEMPLATE;

        Object.entries(options).forEach(([key, value]) => {
            const placeholder = `{{${key}}}`;
            userPrompt = userPrompt.replace(new RegExp(placeholder, 'g'), value);
        });

        // contentプレースホルダーを置換
        userPrompt = userPrompt.replace(/{{content}}/g, content);

        return {
            system: systemPrompt,
            user: userPrompt,
        };
    }

    /**
     * 遅延処理
     */
    _delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * アクティブなリクエストを取得
     */
    getActiveRequests() {
        return Array.from(this.activeRequests.entries()).map(([id, request]) => ({
            id,
            url: request.url,
            method: request.config.method,
            duration: Date.now() - request.startTime,
        }));
    }

    /**
     * リクエストをキャンセル
     */
    cancelRequest(requestId) {
        const request = this.activeRequests.get(requestId);
        if (request && request.controller) {
            request.controller.abort();
            this.activeRequests.delete(requestId);
            logger.info(`リクエストをキャンセルしました: ${requestId}`);
        }
    }

    /**
     * すべてのリクエストをキャンセル
     */
    cancelAllRequests() {
        this.activeRequests.forEach((request, id) => {
            if (request.controller) {
                request.controller.abort();
            }
        });
        this.activeRequests.clear();
        logger.info('すべてのアクティブリクエストをキャンセルしました');
    }

    /**
     * API統計を取得
     */
    getStats() {
        return {
            activeRequests: this.activeRequests.size,
            cache: this.cache.getStats(),
            rateLimiter: {
                endpoints: Array.from(this.rateLimiter.requests.keys()),
                totalRequests: Array.from(this.rateLimiter.requests.values())
                    .reduce((sum, requests) => sum + requests.length, 0),
            },
        };
    }

    /**
     * キャッシュをクリア
     */
    clearCache() {
        this.cache.clear();
    }

    /**
     * 設定をテスト
     */
    async testConnection() {
        try {
            const testMessage = 'このメッセージに「テスト成功」と返答してください。';
            const response = await this.sendChatRequest([
                { role: 'user', content: testMessage }
            ], {
                cacheEnabled: false,
                rateLimitEnabled: false,
            });

            return {
                success: true,
                message: 'API接続テスト成功',
                response: response.data,
            };
        } catch (error) {
            return {
                success: false,
                message: `API接続テスト失敗: ${error.message}`,
                error,
            };
        }
    }
}

/**
 * グローバルAPIクライアントインスタンス
 */
export const apiClient = new PTAAPIClient();
