/*
 * AI業務支援ツール Edge拡張機能 - APIサービス（リファクタリング版）
 * Copyright (c) 2024 AI Business Support Team
 */

import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';
import { HttpUtil } from '../../shared/utils/http.js';
import { API_CONSTANTS } from '../../shared/constants/api.js';
import { ERROR_CONSTANTS } from '../../shared/constants/errors.js';

/**
 * API呼び出しサービス
 */
export class ApiService {
  /**
   * コンストラクタ
   */
  constructor() {
    this.logger = new Logger('ApiService');
    this.validator = new Validator();
    this.httpUtil = new HttpUtil();
  }

  /**
   * AI APIを呼び出し（統一インターフェース）
   * @param {string} prompt - プロンプト
   * @param {Object} settings - API設定
   * @param {Object} options - オプション
   * @returns {Promise<Object>} API応答
   */
  async callAI(prompt, settings, options = {}) {
    const startTime = Date.now();

    try {
      this.logger.info('AI API呼び出し処理を開始します', {
        promptLength: prompt.length,
        provider: settings.provider,
        model: settings.model
      });

      // 入力の妥当性検証
      const validationResult = this.validateApiCall(prompt, settings);
      if (!validationResult.isValid) {
        throw new Error(`API呼び出しパラメータが無効です: ${validationResult.errors.join(', ')}`);
      }

      // APIリクエストを構築
      const requestConfig = this.buildRequestConfig(prompt, settings, options);

      // APIを呼び出し
      const response = await this.executeApiRequest(requestConfig);

      const duration = Date.now() - startTime;
      this.logger.info('AI API呼び出し処理が完了しました', {
        duration: `${duration}ms`,
        success: true
      });

      return {
        success: true,
        data: response.content,
        metadata: {
          provider: settings.provider,
          model: settings.model,
          tokensUsed: response.tokensUsed || 0,
          duration: duration,
          timestamp: new Date().toISOString()
        }
      };

    } catch (error) {
      const duration = Date.now() - startTime;
      this.logger.error('AI API呼び出し処理中にエラーが発生しました', error, {
        duration: `${duration}ms`,
        provider: settings?.provider,
        model: settings?.model
      });

      return {
        success: false,
        error: error.message,
        metadata: {
          provider: settings?.provider,
          model: settings?.model,
          duration: duration,
          timestamp: new Date().toISOString(),
          errorType: this.categorizeError(error)
        }
      };
    }
  }

  /**
   * API呼び出しパラメータの妥当性を検証
   * @param {string} prompt - プロンプト
   * @param {Object} settings - API設定
   * @returns {Object} 検証結果
   * @private
   */
  validateApiCall(prompt, settings) {
    const errors = [];

    try {
      // プロンプトの検証
      if (!this.validator.isString(prompt) || prompt.trim().length === 0) {
        errors.push('プロンプトが無効です');
      } else if (prompt.length > API_CONSTANTS.MAX_PROMPT_LENGTH) {
        errors.push(`プロンプトが長すぎます（最大${API_CONSTANTS.MAX_PROMPT_LENGTH}文字）`);
      }

      // 設定の検証
      if (!settings || typeof settings !== 'object') {
        errors.push('API設定が無効です');
      } else {
        // プロバイダーの検証
        if (!API_CONSTANTS.SUPPORTED_PROVIDERS.includes(settings.provider)) {
          errors.push('サポートされていないAPIプロバイダーです');
        }

        // APIキーの検証
        if (!this.validator.isString(settings.apiKey) || settings.apiKey.trim().length === 0) {
          errors.push('APIキーが設定されていません');
        }

        // モデル名の検証
        if (!this.validator.isString(settings.model) || settings.model.trim().length === 0) {
          errors.push('モデル名が設定されていません');
        }

        // Azure固有の検証
        if (settings.provider === 'azure') {
          if (!this.validator.isUrl(settings.azureEndpoint)) {
            errors.push('AzureエンドポイントURLが無効です');
          }
          if (!this.validator.isString(settings.azureDeploymentName) ||
            settings.azureDeploymentName.trim().length === 0) {
            errors.push('Azureデプロイメント名が設定されていません');
          }
        }

        // 数値パラメータの検証
        if (settings.maxTokens !== undefined &&
          (!this.validator.isPositiveInteger(settings.maxTokens) ||
            settings.maxTokens > API_CONSTANTS.MAX_TOKENS_LIMIT)) {
          errors.push('最大トークン数が無効です');
        }

        if (settings.temperature !== undefined &&
          (!this.validator.isNumber(settings.temperature) ||
            settings.temperature < 0 || settings.temperature > 2)) {
          errors.push('温度パラメータが無効です（0-2の範囲で設定してください）');
        }
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('API呼び出しパラメータ妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['API呼び出しパラメータの妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * APIリクエスト設定を構築
   * @param {string} prompt - プロンプト
   * @param {Object} settings - API設定
   * @param {Object} options - オプション
   * @returns {Object} リクエスト設定
   * @private
   */
  buildRequestConfig(prompt, settings, options) {
    const config = this.getProviderConfig(settings);
    const requestBody = this.buildRequestBody(prompt, settings, options);

    return {
      url: config.endpoint,
      method: 'POST',
      headers: config.headers,
      body: JSON.stringify(requestBody),
      timeout: settings.requestTimeout || API_CONSTANTS.DEFAULT_TIMEOUT,
      retryAttempts: options.retryAttempts || API_CONSTANTS.DEFAULT_RETRY_ATTEMPTS,
      retryDelay: options.retryDelay || API_CONSTANTS.DEFAULT_RETRY_DELAY
    };
  }

  /**
   * プロバイダー固有の設定を取得
   * @param {Object} settings - API設定
   * @returns {Object} プロバイダー設定
   * @private
   */
  getProviderConfig(settings) {
    switch (settings.provider) {
      case 'openai':
        return {
          endpoint: API_CONSTANTS.ENDPOINTS.OPENAI,
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${settings.apiKey}`
          }
        };

      case 'azure':
        const azureEndpoint = `${settings.azureEndpoint}/openai/deployments/${settings.azureDeploymentName}/chat/completions?api-version=${API_CONSTANTS.AZURE_API_VERSION}`;
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
   * @param {string} prompt - プロンプト
   * @param {Object} settings - API設定
   * @param {Object} options - オプション
   * @returns {Object} リクエストボディ
   * @private
   */
  buildRequestBody(prompt, settings, options) {
    const systemPrompt = options.systemPrompt || API_CONSTANTS.DEFAULT_SYSTEM_PROMPT;

    const body = {
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: prompt }
      ],
      max_tokens: options.maxTokens || settings.maxTokens || API_CONSTANTS.DEFAULT_MAX_TOKENS,
      temperature: options.temperature !== undefined ? options.temperature :
        (settings.temperature !== undefined ? settings.temperature : API_CONSTANTS.DEFAULT_TEMPERATURE),
      top_p: options.topP || API_CONSTANTS.DEFAULT_TOP_P,
      frequency_penalty: options.frequencyPenalty || API_CONSTANTS.DEFAULT_FREQUENCY_PENALTY,
      presence_penalty: options.presencePenalty || API_CONSTANTS.DEFAULT_PRESENCE_PENALTY
    };

    // OpenAI の場合はモデル名を含める
    if (settings.provider === 'openai') {
      body.model = settings.model;
    }

    // ストリーミングは現在未対応
    body.stream = false;

    return body;
  }

  /**
   * APIリクエストを実行
   * @param {Object} config - リクエスト設定
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async executeApiRequest(config) {
    try {
      // Chrome拡張のOffscreen documentを使用してCORS制限を回避
      const response = await this.fetchViaOffscreen(config);

      if (!response.success) {
        throw new Error(response.error || 'API呼び出しに失敗しました');
      }

      return this.parseApiResponse(response.data);
    } catch (error) {
      this.logger.error('APIリクエスト実行中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * Offscreen documentを使用してAPIを呼び出し
   * @param {Object} config - リクエスト設定
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async fetchViaOffscreen(config) {
    try {
      // Offscreen documentの存在を確認・作成
      await this.ensureOffscreenDocument();

      // Offscreen documentにメッセージを送信
      const response = await chrome.runtime.sendMessage({
        target: 'offscreen',
        action: 'fetchAPI',
        data: config
      });

      return response || { success: false, error: 'レスポンスが受信されませんでした' };
    } catch (error) {
      this.logger.error('Offscreen document経由のAPI呼び出し中にエラーが発生しました', error);
      throw new Error(`API呼び出しエラー: ${error.message}`);
    }
  }

  /**
   * Offscreen documentの存在を確認・作成
   * @private
   */
  async ensureOffscreenDocument() {
    try {
      // 既存のOffscreen documentをチェック
      const existingContexts = await chrome.runtime.getContexts({
        contextTypes: ['OFFSCREEN_DOCUMENT']
      });

      if (existingContexts.length > 0) {
        return; // 既に存在する
      }

      // Offscreen documentを作成
      await chrome.offscreen.createDocument({
        url: chrome.runtime.getURL('offscreen/offscreen.html'),
        reasons: ['BLOBS'],
        justification: 'API呼び出しのためのCORS制限回避'
      });

      this.logger.debug('Offscreen documentを作成しました');
    } catch (error) {
      this.logger.error('Offscreen document作成中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * APIレスポンスを解析
   * @param {Object} responseData - レスポンスデータ
   * @returns {Object} 解析済みレスポンス
   * @private
   */
  parseApiResponse(responseData) {
    try {
      if (!responseData || !responseData.choices || responseData.choices.length === 0) {
        throw new Error('無効なAPIレスポンス形式です');
      }

      const choice = responseData.choices[0];
      if (!choice.message || !choice.message.content) {
        throw new Error('APIレスポンスにコンテンツが含まれていません');
      }

      return {
        content: choice.message.content.trim(),
        tokensUsed: responseData.usage ? {
          promptTokens: responseData.usage.prompt_tokens || 0,
          completionTokens: responseData.usage.completion_tokens || 0,
          totalTokens: responseData.usage.total_tokens || 0
        } : null,
        finishReason: choice.finish_reason || 'unknown'
      };
    } catch (error) {
      this.logger.error('APIレスポンス解析中にエラーが発生しました', error);
      throw new Error(`レスポンス解析エラー: ${error.message}`);
    }
  }

  /**
   * エラーを分類
   * @param {Error} error - エラーオブジェクト
   * @returns {string} エラータイプ
   * @private
   */
  categorizeError(error) {
    const message = error.message?.toLowerCase() || '';

    if (message.includes('unauthorized') || message.includes('401')) {
      return ERROR_CONSTANTS.TYPES.AUTHENTICATION;
    } else if (message.includes('rate limit') || message.includes('429')) {
      return ERROR_CONSTANTS.TYPES.RATE_LIMIT;
    } else if (message.includes('timeout') || message.includes('408')) {
      return ERROR_CONSTANTS.TYPES.TIMEOUT;
    } else if (message.includes('network') || message.includes('connection')) {
      return ERROR_CONSTANTS.TYPES.NETWORK;
    } else if (message.includes('quota') || message.includes('insufficient')) {
      return ERROR_CONSTANTS.TYPES.QUOTA_EXCEEDED;
    } else {
      return ERROR_CONSTANTS.TYPES.UNKNOWN;
    }
  }

  /**
   * 利用可能なモデル一覧を取得
   * @param {Object} settings - API設定
   * @returns {Promise<Array>} モデル一覧
   */
  async getAvailableModels(settings) {
    try {
      this.logger.info('利用可能モデル一覧取得処理を開始します', { provider: settings.provider });

      // プロバイダー別のモデル一覧を返す
      switch (settings.provider) {
        case 'openai':
          return API_CONSTANTS.MODELS.OPENAI;
        case 'azure':
          // Azureの場合は設定から取得またはデフォルト
          return settings.availableModels || API_CONSTANTS.MODELS.AZURE;
        default:
          throw new Error(`サポートされていないプロバイダー: ${settings.provider}`);
      }
    } catch (error) {
      this.logger.error('利用可能モデル一覧取得処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * API接続をテスト
   * @param {Object} settings - API設定
   * @returns {Promise<Object>} テスト結果
   */
  async testConnection(settings) {
    try {
      this.logger.info('API接続テストを開始します', { provider: settings.provider });

      const testPrompt = 'テスト';
      const testOptions = {
        maxTokens: 10,
        temperature: 0
      };

      const result = await this.callAI(testPrompt, settings, testOptions);

      this.logger.info('API接続テストが完了しました', { success: result.success });
      return {
        success: result.success,
        message: result.success ? '接続成功' : result.error,
        latency: result.metadata?.duration || 0
      };
    } catch (error) {
      this.logger.error('API接続テスト中にエラーが発生しました', error);
      return {
        success: false,
        message: `接続テストに失敗しました: ${error.message}`,
        latency: 0
      };
    }
  }
}
