/*
 * PTA AI業務支援ツール - テストスイート
 * Copyright (c) 2024 PTA Development Team
 */

import { describe, test, expect, beforeEach, afterEach, jest } from '@jest/globals';

// テスト対象のモジュール
import { logger, createLogger } from '../core/logger.js';
import { errorHandler, PTAError, PTAAPIError } from '../core/error-handler.js';
import { settingsManager } from '../core/settings-manager.js';
import { apiClient } from '../core/api-client.js';
import { eventManager, EVENTS } from '../core/event-manager.js';

/**
 * ログシステムのテスト
 */
describe('Logger System', () => {
    beforeEach(() => {
        logger.clear();
    });

    test('ログエントリの作成と記録', () => {
        logger.info('テストメッセージ', { key: 'value' });

        const history = logger.getHistory();
        expect(history).toHaveLength(1);
        expect(history[0].message).toBe('テストメッセージ');
        expect(history[0].level).toBe('INFO');
        expect(history[0].data).toEqual({ key: 'value' });
    });

    test('ログレベルによるフィルタリング', () => {
        logger.debug('デバッグメッセージ');
        logger.info('情報メッセージ');
        logger.error('エラーメッセージ');

        const errorLogs = logger.getHistory('ERROR');
        expect(errorLogs).toHaveLength(1);
        expect(errorLogs[0].message).toBe('エラーメッセージ');
    });

    test('パフォーマンス測定', () => {
        const perf = logger.startPerformance('テスト処理');
        // 短い処理をシミュレート
        const duration = perf.end();

        expect(duration).toBeGreaterThan(0);
    });

    test('コンテキスト別ロガーの作成', () => {
        const contextLogger = createLogger('テストコンテキスト');
        contextLogger.info('コンテキストメッセージ');

        const history = contextLogger.getHistory();
        expect(history[0].context).toBe('テストコンテキスト');
    });
});

/**
 * エラーハンドリングのテスト
 */
describe('Error Handler', () => {
    beforeEach(() => {
        errorHandler.clearHistory();
    });

    test('PTAErrorの作成と処理', () => {
        const error = new PTAError('テストエラー', 'TEST_ERROR', { detail: 'test' });

        expect(error.message).toBe('テストエラー');
        expect(error.type).toBe('TEST_ERROR');
        expect(error.details).toEqual({ detail: 'test' });
    });

    test('APIエラーの処理', () => {
        const apiError = new PTAAPIError('API呼び出しエラー', 500, { error: 'Server Error' });

        expect(apiError.statusCode).toBe(500);
        expect(apiError.response).toEqual({ error: 'Server Error' });
    });

    test('エラーハンドリングと履歴記録', () => {
        const error = new Error('通常のJavaScriptエラー');

        errorHandler.handle(error, 'テストコンテキスト', false);

        const history = errorHandler.getErrorHistory();
        expect(history).toHaveLength(1);
        expect(history[0].context).toBe('テストコンテキスト');
    });

    test('リトライ機能', async () => {
        let attempts = 0;
        const operation = jest.fn().mockImplementation(() => {
            attempts++;
            if (attempts < 3) {
                throw new Error('一時的なエラー');
            }
            return 'success';
        });

        const result = await errorHandler.retry(operation, 3, 10);

        expect(result).toBe('success');
        expect(attempts).toBe(3);
    });
});

/**
 * 設定管理のテスト
 */
describe('Settings Manager', () => {
    beforeEach(async () => {
        // テスト用のクリーンな設定で初期化
        await settingsManager.reset();
    });

    test('デフォルト設定の取得', () => {
        const language = settingsManager.get('userPreferences', 'language');
        expect(language).toBe('ja');
    });

    test('設定の更新と取得', async () => {
        await settingsManager.set('userPreferences', 'theme', 'dark');

        const theme = settingsManager.get('userPreferences', 'theme');
        expect(theme).toBe('dark');
    });

    test('設定バリデーション - 不正な値', async () => {
        await expect(
            settingsManager.set('apiSettings', 'maxTokens', 'invalid')
        ).rejects.toThrow('設定「apiSettings.maxTokens」の型が正しくありません');
    });

    test('設定バリデーション - 範囲外の値', async () => {
        await expect(
            settingsManager.set('apiSettings', 'maxTokens', 10000)
        ).rejects.toThrow('設定「apiSettings.maxTokens」の値が大きすぎます');
    });

    test('設定の監視', (done) => {
        const unwatch = settingsManager.watch('userPreferences', 'theme', (value) => {
            expect(value).toBe('light');
            unwatch();
            done();
        });

        settingsManager.set('userPreferences', 'theme', 'light');
    });
});

/**
 * イベントマネージャーのテスト
 */
describe('Event Manager', () => {
    beforeEach(() => {
        eventManager.removeAllListeners();
        eventManager.clearHistory();
    });

    test('イベントの登録と発火', async () => {
        const callback = jest.fn();

        eventManager.on('test:event', callback);
        await eventManager.emit('test:event', { data: 'test' });

        expect(callback).toHaveBeenCalledWith({ data: 'test' }, expect.any(Object));
    });

    test('一度だけ実行されるイベントリスナー', async () => {
        const callback = jest.fn();

        eventManager.once('test:once', callback);
        await eventManager.emit('test:once');
        await eventManager.emit('test:once');

        expect(callback).toHaveBeenCalledTimes(1);
    });

    test('優先度付きイベントリスナー', async () => {
        const results = [];

        eventManager.on('test:priority', () => results.push('low'), { priority: 1 });
        eventManager.on('test:priority', () => results.push('high'), { priority: 10 });

        await eventManager.emit('test:priority');

        expect(results).toEqual(['high', 'low']);
    });

    test('イベント待機機能', async () => {
        setTimeout(() => {
            eventManager.emit('test:wait', { result: 'success' });
        }, 100);

        const { data } = await eventManager.waitFor('test:wait', null, 1000);
        expect(data.result).toBe('success');
    });

    test('条件付きイベント待機', async () => {
        setTimeout(() => {
            eventManager.emit('test:condition', { value: 10 });
        }, 50);

        setTimeout(() => {
            eventManager.emit('test:condition', { value: 20 });
        }, 100);

        const { data } = await eventManager.waitFor(
            'test:condition',
            (data) => data.value >= 20,
            1000
        );

        expect(data.value).toBe(20);
    });
});

/**
 * APIクライアントのテスト
 */
describe('API Client', () => {
    beforeEach(() => {
        // モックの設定
        global.fetch = jest.fn();
        apiClient.clearCache();
    });

    afterEach(() => {
        jest.restoreAllMocks();
    });

    test('成功時のAPIレスポンス処理', async () => {
        const mockResponse = {
            ok: true,
            status: 200,
            headers: new Map([['content-type', 'application/json']]),
            json: () => Promise.resolve({ choices: [{ message: { content: 'テスト応答' } }] }),
        };

        global.fetch.mockResolvedValueOnce(mockResponse);

        // settingsManagerにテスト用の設定を追加
        settingsManager.settings = {
            apiSettings: {
                openaiApiKey: 'test-api-key',
                openaiEndpoint: 'https://api.openai.com/v1/chat/completions',
                openaiModel: 'gpt-4',
                maxTokens: 1000,
                temperature: 0.7,
            }
        };

        const result = await apiClient.sendChatRequest([
            { role: 'user', content: 'テストメッセージ' }
        ]);

        expect(result.success).toBe(true);
        expect(result.data.choices[0].message.content).toBe('テスト応答');
    });

    test('APIエラー時の処理', async () => {
        const mockResponse = {
            ok: false,
            status: 401,
            statusText: 'Unauthorized',
            headers: new Map([['content-type', 'application/json']]),
            json: () => Promise.resolve({ error: 'Invalid API key' }),
        };

        global.fetch.mockResolvedValueOnce(mockResponse);

        settingsManager.settings = {
            apiSettings: {
                openaiApiKey: 'invalid-key',
                openaiEndpoint: 'https://api.openai.com/v1/chat/completions',
                openaiModel: 'gpt-4',
                maxTokens: 1000,
                temperature: 0.7,
            }
        };

        await expect(
            apiClient.sendChatRequest([{ role: 'user', content: 'テスト' }])
        ).rejects.toThrow('HTTPエラー: 401 Unauthorized');
    });

    test('キャッシュ機能', async () => {
        const mockResponse = {
            ok: true,
            status: 200,
            headers: new Map([['content-type', 'application/json']]),
            json: () => Promise.resolve({ cached: 'response' }),
        };

        global.fetch.mockResolvedValueOnce(mockResponse);

        settingsManager.settings = {
            apiSettings: {
                openaiApiKey: 'test-key',
                openaiEndpoint: 'https://api.openai.com/v1/chat/completions',
                openaiModel: 'gpt-4',
                maxTokens: 1000,
                temperature: 0.7,
            },
            cacheSettings: {
                enabled: true,
                maxSize: 10485760,
                ttl: 3600000,
            }
        };

        // 最初のリクエスト
        const result1 = await apiClient.sendChatRequest([
            { role: 'user', content: 'キャッシュテスト' }
        ], { cacheEnabled: true });

        // 2回目のリクエスト（キャッシュから取得）
        const result2 = await apiClient.sendChatRequest([
            { role: 'user', content: 'キャッシュテスト' }
        ], { cacheEnabled: true });

        // fetchは1回だけ呼ばれる（2回目はキャッシュ）
        expect(global.fetch).toHaveBeenCalledTimes(1);
        expect(result1.data).toEqual(result2.data);
    });
});

/**
 * 統合テスト
 */
describe('Integration Tests', () => {
    test('設定変更時のイベント発火', async () => {
        const settingsChangedCallback = jest.fn();

        eventManager.on(EVENTS.SETTINGS_CHANGED, settingsChangedCallback);

        await settingsManager.set('userPreferences', 'language', 'en');

        expect(settingsChangedCallback).toHaveBeenCalledWith(
            expect.objectContaining({
                category: 'userPreferences',
                key: 'language',
                value: 'en',
            })
        );
    });

    test('API呼び出し時のイベント発火', async () => {
        const apiStartCallback = jest.fn();
        const apiSuccessCallback = jest.fn();

        eventManager.on(EVENTS.API_REQUEST_START, apiStartCallback);
        eventManager.on(EVENTS.API_REQUEST_SUCCESS, apiSuccessCallback);

        const mockResponse = {
            ok: true,
            status: 200,
            headers: new Map([['content-type', 'application/json']]),
            json: () => Promise.resolve({ test: 'response' }),
        };

        global.fetch.mockResolvedValueOnce(mockResponse);

        settingsManager.settings = {
            apiSettings: {
                openaiApiKey: 'test-key',
                openaiEndpoint: 'https://api.openai.com/v1/chat/completions',
                openaiModel: 'gpt-4',
                maxTokens: 1000,
                temperature: 0.7,
            }
        };

        await apiClient.sendChatRequest([
            { role: 'user', content: 'テスト' }
        ]);

        expect(apiStartCallback).toHaveBeenCalled();
        expect(apiSuccessCallback).toHaveBeenCalled();
    });

    test('エラー発生時の統合処理', async () => {
        const appErrorCallback = jest.fn();

        eventManager.on(EVENTS.APP_ERROR, appErrorCallback);

        const error = new PTAError('統合テストエラー', 'INTEGRATION_TEST');

        errorHandler.handle(error, '統合テスト');

        // エラーが適切に処理され、統計に記録されることを確認
        const stats = errorHandler.getErrorStats();
        expect(stats.total).toBe(1);
        expect(stats.byType['INTEGRATION_TEST']).toBe(1);
    });
});

/**
 * ユーティリティ関数
 */
function mockChromeAPI() {
    global.chrome = {
        storage: {
            local: {
                get: jest.fn().mockResolvedValue({}),
                set: jest.fn().mockResolvedValue(),
            }
        },
        runtime: {
            sendMessage: jest.fn(),
            getManifest: jest.fn().mockReturnValue({ version: '2.0.0' }),
        }
    };
}

// Chrome API のモック設定
beforeEach(() => {
    mockChromeAPI();
});

export {
    mockChromeAPI,
};
