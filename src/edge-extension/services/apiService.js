/*
 * AI業務支援ツール Edge拡張機能 - APIサービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { ApiRequest, ApiResponse } from '../lib/models.js';
import { API_PROVIDERS, API_ENDPOINTS, ERROR_MESSAGES } from '../lib/constants.js';
import { Logger, retryAsync, performanceMonitor } from '../lib/utils.js';
import { SettingsService } from './settingsService.js';

/**
 * API呼び出しサービス
 */
export class ApiService {
  constructor() {
    this.settingsService = new SettingsService();
  }

  /**
   * AI APIを呼び出し（統一インターフェース）
   * @param {string} prompt プロンプト
   * @param {Object} options オプション
   * @returns {Promise<ApiResponse>} API応答
   */
  async callAI(prompt, options = {}) {
    const startTime = Date.now();
    performanceMonitor.start('aiApiCall');

    try {
      Logger.info('AI API呼び出し開始', { prompt: prompt.substring(0, 100) + '...' });

      // 設定を取得
      const settings = await this.settingsService.getSettings();
      const validation = settings.validate();

      if (!validation.isValid) {
        throw new Error(`設定エラー: ${validation.getErrorMessage()}`);
      }

      // API リクエストを構築
      const apiRequest = this.buildApiRequest(prompt, settings, options);
      const requestValidation = apiRequest.validate();

      if (!requestValidation.isValid) {
        throw new Error(`リクエストエラー: ${requestValidation.getErrorMessage()}`);
      }

      // APIを呼び出し
      const result = await this.executeApiRequest(apiRequest, settings);

      const duration = performanceMonitor.end('aiApiCall');
      Logger.info('AI API呼び出し成功', { duration: `${duration.toFixed(2)}ms` });

      return new ApiResponse({
        success: true,
        data: result,
        statusCode: 200,
        timestamp: Date.now(),
        duration: Date.now() - startTime
      });

    } catch (error) {
      const duration = performanceMonitor.end('aiApiCall');
      Logger.error('AI API呼び出しエラー', error);

      return new ApiResponse({
        success: false,
        error: error.message,
        statusCode: this.extractStatusCode(error),
        timestamp: Date.now(),
        duration: Date.now() - startTime
      });
    }
  }

  /**
   * APIリクエストを構築
   * @param {string} prompt プロンプト
   * @param {ApiSettings} settings 設定
   * @param {Object} options オプション
   * @returns {ApiRequest} APIリクエスト
   */
  buildApiRequest(prompt, settings, options) {
    const { endpoint, headers } = this.getApiConfiguration(settings);
    const body = this.buildRequestBody(prompt, settings, options);

    return new ApiRequest({
      endpoint,
      headers,
      body: JSON.stringify(body),
      method: 'POST',
      timeout: settings.requestTimeout,
      retryAttempts: options.retryAttempts || 3,
      retryDelay: options.retryDelay || 1000
    });
  }

  /**
   * API設定を取得
   * @param {ApiSettings} settings 設定
   * @returns {Object} エンドポイントとヘッダー
   */
  getApiConfiguration(settings) {
    switch (settings.provider) {
      case API_PROVIDERS.OPENAI:
        return {
          endpoint: API_ENDPOINTS.OPENAI,
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${settings.apiKey}`
          }
        };

      case API_PROVIDERS.AZURE:
        const azureEndpoint = API_ENDPOINTS.AZURE_TEMPLATE
          .replace('{resource}', this.extractResourceName(settings.azureEndpoint))
          .replace('{deployment}', settings.azureDeploymentName);

        return {
          endpoint: azureEndpoint,
          headers: {
            'Content-Type': 'application/json',
            'api-key': settings.apiKey
          }
        };

      default:
        throw new Error(`サポートされていないプロバイダー: ${settings.provider}`);
    }
  }

  /**
   * リクエストボディを構築
   * @param {string} prompt プロンプト
   * @param {ApiSettings} settings 設定
   * @param {Object} options オプション
   * @returns {Object} リクエストボディ
   */
  buildRequestBody(prompt, settings, options) {
    const systemPrompt = options.systemPrompt || 'あなたは優秀なAIアシスタントです。日本語で回答してください。';

    const body = {
      model: settings.model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: prompt }
      ],
      max_tokens: options.maxTokens || settings.maxTokens,
      temperature: options.temperature || settings.temperature
    };

    // Azure OpenAI の場合はモデル名を除去
    if (settings.provider === API_PROVIDERS.AZURE) {
      delete body.model;
    }

    return body;
  }

  /**
   * APIリクエストを実行
   * @param {ApiRequest} request APIリクエスト
   * @param {ApiSettings} settings 設定
   * @returns {Promise<string>} AI応答テキスト
   */
  async executeApiRequest(request, settings) {
    return await retryAsync(async () => {
      // Chrome拡張のOffscreen documentを使用してCORS制限を回避
      const response = await this.fetchViaOffscreen(request, settings);

      if (!response.success) {
        throw new Error(response.error);
      }

      return response.data;
    }, request.retryAttempts, request.retryDelay);
  }

  /**
   * Offscreen documentを使用してfetchを実行
   * @param {ApiRequest} request APIリクエスト
   * @param {ApiSettings} settings 設定
   * @returns {Promise<Object>} 応答データ
   */
  async fetchViaOffscreen(request, settings) {
    try {
      // Offscreen documentが利用可能かチェック
      await this.ensureOffscreenDocument();

      // Offscreen documentにメッセージを送信
      const response = await chrome.runtime.sendMessage({
        target: 'offscreen',
        action: 'fetchAPI',
        data: {
          endpoint: request.endpoint,
          headers: request.headers,
          body: request.body,
          provider: settings.provider
        }
      });

      return response;

    } catch (error) {
      Logger.warn('Offscreen document利用不可、フォールバックを試行', error.message);

      // フォールバック: 直接fetchを試行（制限あり）
      return await this.fallbackDirectFetch(request, settings);
    }
  }

  /**
   * Offscreen documentの存在確認・作成
   * @returns {Promise<void>}
   */
  async ensureOffscreenDocument() {
    if (typeof chrome.offscreen === 'undefined') {
      throw new Error('Offscreen API is not available');
    }

    try {
      // 既存のoffscreen documentをチェック
      const existingClients = await chrome.runtime.getContexts({
        contextTypes: ['OFFSCREEN_DOCUMENT']
      });

      if (existingClients.length === 0) {
        // Offscreen documentを作成
        await chrome.offscreen.createDocument({
          url: 'offscreen/offscreen.html',
          reasons: ['CORS'],
          justification: 'CORS制限回避のためのAPI呼び出し'
        });

        Logger.info('Offscreen document作成完了');
      }
    } catch (error) {
      Logger.error('Offscreen document作成エラー', error);
      throw error;
    }
  }

  /**
   * フォールバック: 直接fetch（制限あり）
   * @param {ApiRequest} request APIリクエスト
   * @param {ApiSettings} settings 設定
   * @returns {Promise<Object>} 応答データ
   */
  async fallbackDirectFetch(request, settings) {
    try {
      Logger.warn('直接fetch実行（CORS制限の可能性あり）');

      const response = await fetch(request.endpoint, {
        method: request.method,
        headers: request.headers,
        body: request.body,
        mode: 'cors',
        credentials: 'omit'
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      const content = this.extractContentFromResponse(data, settings.provider);

      return { success: true, data: content };

    } catch (error) {
      Logger.error('直接fetch失敗', error);
      return { success: false, error: error.message };
    }
  }

  /**
   * API応答からコンテンツを抽出
   * @param {Object} data 応答データ
   * @param {string} provider プロバイダー
   * @returns {string} 抽出されたコンテンツ
   */
  extractContentFromResponse(data, provider) {
    try {
      if (provider === API_PROVIDERS.OPENAI || provider === API_PROVIDERS.AZURE) {
        if (data.choices && data.choices.length > 0) {
          return data.choices[0].message.content.trim();
        }
      }

      throw new Error('無効なAPI応答形式');
    } catch (error) {
      Logger.error('応答解析エラー', error);
      throw new Error(ERROR_MESSAGES.INVALID_RESPONSE);
    }
  }

  /**
   * Azure エンドポイントからリソース名を抽出
   * @param {string} endpoint エンドポイントURL
   * @returns {string} リソース名
   */
  extractResourceName(endpoint) {
    try {
      const url = new URL(endpoint);
      return url.hostname.split('.')[0];
    } catch (error) {
      throw new Error('無効なAzure エンドポイントURL');
    }
  }

  /**
   * エラーからHTTPステータスコードを抽出
   * @param {Error} error エラーオブジェクト
   * @returns {number} ステータスコード
   */
  extractStatusCode(error) {
    const message = error.message.toLowerCase();

    if (message.includes('401') || message.includes('unauthorized')) return 401;
    if (message.includes('403') || message.includes('forbidden')) return 403;
    if (message.includes('404') || message.includes('not found')) return 404;
    if (message.includes('429') || message.includes('rate limit')) return 429;
    if (message.includes('500') || message.includes('internal server')) return 500;
    if (message.includes('502') || message.includes('bad gateway')) return 502;
    if (message.includes('503') || message.includes('service unavailable')) return 503;

    return 0; // Unknown
  }

  /**
   * API接続テスト
   * @returns {Promise<ApiResponse>} テスト結果
   */
  async testConnection() {
    Logger.info('API接続テスト開始');

    try {
      const testPrompt = 'こんにちは。これは接続テストです。「接続成功」と返答してください。';
      const response = await this.callAI(testPrompt, {
        maxTokens: 50,
        temperature: 0.1,
        systemPrompt: '簡潔に「接続成功」と返答してください。'
      });

      if (response.isSuccess()) {
        Logger.info('API接続テスト成功');
      } else {
        Logger.error('API接続テスト失敗', response.error);
      }

      return response;

    } catch (error) {
      Logger.error('API接続テストエラー', error);

      return new ApiResponse({
        success: false,
        error: error.message,
        timestamp: Date.now()
      });
    }
  }
}
