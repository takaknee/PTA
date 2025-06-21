/*
 * AI業務支援ツール Edge拡張機能 - プロンプトサービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';
import { APP_CONSTANTS } from '../../shared/constants/app.js';

/**
 * プロンプト管理サービス
 */
export class PromptService {
  /**
   * コンストラクタ
   */
  constructor() {
    this.logger = new Logger('PromptService');
    this.validator = new Validator();
    this.templates = this.initializeTemplates();
  }

  /**
   * プロンプトテンプレートを初期化
   * @returns {Object} テンプレート一覧
   * @private
   */
  initializeTemplates() {
    return {
      // 基本テンプレート
      summarize: {
        name: '要約',
        description: 'テキストの要約を作成します',
        systemPrompt: 'あなたは優秀な要約専門家です。与えられたテキストを簡潔で分かりやすく要約してください。',
        template: '以下のテキストを要約してください：\n\n{content}',
        category: 'basic',
        parameters: ['content']
      },
      translate: {
        name: '翻訳',
        description: 'テキストを指定言語に翻訳します',
        systemPrompt: 'あなたは優秀な翻訳専門家です。文脈を理解して自然な翻訳を提供してください。',
        template: '以下のテキストを{targetLanguage}に翻訳してください：\n\n{content}',
        category: 'basic',
        parameters: ['content', 'targetLanguage']
      },
      explain: {
        name: '解説',
        description: '複雑な概念を分かりやすく説明します',
        systemPrompt: 'あなたは優秀な教育者です。複雑な内容を初心者にも分かりやすく説明してください。',
        template: '以下の内容について分かりやすく説明してください：\n\n{content}',
        category: 'basic',
        parameters: ['content']
      },

      // ビジネステンプレート
      businessEmail: {
        name: 'ビジネスメール作成',
        description: 'ビジネスメールの下書きを作成します',
        systemPrompt: 'あなたは優秀なビジネスライターです。適切な敬語と丁寧な表現でメールを作成してください。',
        template: '以下の内容でビジネスメールを作成してください：\n\n件名: {subject}\n宛先: {recipient}\n内容: {content}',
        category: 'business',
        parameters: ['subject', 'recipient', 'content']
      },
      meetingMinutes: {
        name: '議事録作成',
        description: '会議の議事録を整理します',
        systemPrompt: 'あなたは優秀な秘書です。会議の内容を整理して読みやすい議事録を作成してください。',
        template: '以下の会議内容から議事録を作成してください：\n\n会議名: {meetingName}\n日時: {dateTime}\n参加者: {participants}\n内容: {content}',
        category: 'business',
        parameters: ['meetingName', 'dateTime', 'participants', 'content']
      },

      // 学習テンプレート
      quiz: {
        name: 'クイズ作成',
        description: 'テキストからクイズを作成します',
        systemPrompt: 'あなたは優秀な教育者です。学習効果の高いクイズを作成してください。',
        template: '以下のテキストから{questionCount}問のクイズを作成してください：\n\n{content}',
        category: 'education',
        parameters: ['content', 'questionCount']
      },

      // 創作テンプレート
      creative: {
        name: '創作支援',
        description: '創作活動をサポートします',
        systemPrompt: 'あなたは創造性豊かなライターです。魅力的で独創的な内容を創作してください。',
        template: '以下のテーマで{type}を作成してください：\n\nテーマ: {theme}\n要求: {requirements}',
        category: 'creative',
        parameters: ['type', 'theme', 'requirements']
      }
    };
  }

  /**
   * 利用可能なプロンプトテンプレート一覧を取得
   * @param {string} [category] - カテゴリーフィルター
   * @returns {Array} テンプレート一覧
   */
  getTemplates(category = null) {
    try {
      this.logger.debug('プロンプトテンプレート一覧取得を開始します', { category });

      const templates = Object.entries(this.templates).map(([id, template]) => ({
        id,
        ...template
      }));

      const filteredTemplates = category ?
        templates.filter(template => template.category === category) :
        templates;

      this.logger.debug(`プロンプトテンプレート一覧取得が完了しました（${filteredTemplates.length}件）`);
      return filteredTemplates;
    } catch (error) {
      this.logger.error('プロンプトテンプレート一覧取得中にエラーが発生しました', error);
      return [];
    }
  }

  /**
   * プロンプトテンプレートを取得
   * @param {string} templateId - テンプレートID
   * @returns {Object|null} テンプレート
   */
  getTemplate(templateId) {
    try {
      this.logger.debug('プロンプトテンプレート取得を開始します', { templateId });

      if (!this.validator.isString(templateId) || !this.templates[templateId]) {
        this.logger.warn('指定されたテンプレートが見つかりません', { templateId });
        return null;
      }

      const template = { id: templateId, ...this.templates[templateId] };

      this.logger.debug('プロンプトテンプレート取得が完了しました', { templateId });
      return template;
    } catch (error) {
      this.logger.error('プロンプトテンプレート取得中にエラーが発生しました', error);
      return null;
    }
  }

  /**
   * プロンプトを生成
   * @param {string} templateId - テンプレートID
   * @param {Object} parameters - パラメータ
   * @returns {Object} 生成されたプロンプト
   */
  generatePrompt(templateId, parameters = {}) {
    try {
      this.logger.debug('プロンプト生成処理を開始します', { templateId, parameters });

      const template = this.getTemplate(templateId);
      if (!template) {
        throw new Error(`テンプレートが見つかりません: ${templateId}`);
      }

      // パラメータの妥当性検証
      const validationResult = this.validateParameters(template, parameters);
      if (!validationResult.isValid) {
        throw new Error(`パラメータが無効です: ${validationResult.errors.join(', ')}`);
      }

      // プロンプトを生成
      const prompt = this.interpolateTemplate(template.template, parameters);
      const systemPrompt = template.systemPrompt;

      const result = {
        prompt,
        systemPrompt,
        template: template,
        parameters: parameters,
        generatedAt: new Date().toISOString()
      };

      this.logger.debug('プロンプト生成処理が完了しました', { templateId });
      return result;
    } catch (error) {
      this.logger.error('プロンプト生成処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * カスタムプロンプトを作成
   * @param {string} prompt - プロンプト
   * @param {string} [systemPrompt] - システムプロンプト
   * @returns {Object} プロンプト情報
   */
  createCustomPrompt(prompt, systemPrompt = null) {
    try {
      this.logger.debug('カスタムプロンプト作成処理を開始します');

      // プロンプトの妥当性検証
      if (!this.validator.isString(prompt) || prompt.trim().length === 0) {
        throw new Error('プロンプトが無効です');
      }

      if (prompt.length > APP_CONSTANTS.MAX_PROMPT_LENGTH) {
        throw new Error(`プロンプトが長すぎます（最大${APP_CONSTANTS.MAX_PROMPT_LENGTH}文字）`);
      }

      const result = {
        prompt: prompt.trim(),
        systemPrompt: systemPrompt || APP_CONSTANTS.DEFAULT_SYSTEM_PROMPT,
        template: null,
        parameters: {},
        generatedAt: new Date().toISOString()
      };

      this.logger.debug('カスタムプロンプト作成処理が完了しました');
      return result;
    } catch (error) {
      this.logger.error('カスタムプロンプト作成処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * プロンプトを最適化
   * @param {string} prompt - 元のプロンプト
   * @param {Object} options - 最適化オプション
   * @returns {Object} 最適化されたプロンプト
   */
  optimizePrompt(prompt, options = {}) {
    try {
      this.logger.debug('プロンプト最適化処理を開始します');

      if (!this.validator.isString(prompt) || prompt.trim().length === 0) {
        throw new Error('プロンプトが無効です');
      }

      let optimizedPrompt = prompt.trim();

      // 基本的な最適化
      optimizedPrompt = this.applyBasicOptimizations(optimizedPrompt, options);

      // 長さ制限の適用
      if (optimizedPrompt.length > APP_CONSTANTS.MAX_PROMPT_LENGTH) {
        optimizedPrompt = this.truncatePrompt(optimizedPrompt, APP_CONSTANTS.MAX_PROMPT_LENGTH);
      }

      const result = {
        originalPrompt: prompt,
        optimizedPrompt: optimizedPrompt,
        optimizations: this.getAppliedOptimizations(prompt, optimizedPrompt),
        generatedAt: new Date().toISOString()
      };

      this.logger.debug('プロンプト最適化処理が完了しました');
      return result;
    } catch (error) {
      this.logger.error('プロンプト最適化処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * テンプレートのパラメータを検証
   * @param {Object} template - テンプレート
   * @param {Object} parameters - パラメータ
   * @returns {Object} 検証結果
   * @private
   */
  validateParameters(template, parameters) {
    const errors = [];

    try {
      // 必須パラメータの確認
      for (const requiredParam of template.parameters) {
        if (!(requiredParam in parameters)) {
          errors.push(`必須パラメータが不足しています: ${requiredParam}`);
        } else if (!this.validator.isString(parameters[requiredParam]) ||
          parameters[requiredParam].trim().length === 0) {
          errors.push(`パラメータが無効です: ${requiredParam}`);
        }
      }

      // パラメータの長さ制限
      for (const [key, value] of Object.entries(parameters)) {
        if (this.validator.isString(value) && value.length > APP_CONSTANTS.MAX_PARAMETER_LENGTH) {
          errors.push(`パラメータが長すぎます: ${key} (最大${APP_CONSTANTS.MAX_PARAMETER_LENGTH}文字)`);
        }
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('パラメータ妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['パラメータの妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * テンプレートにパラメータを挿入
   * @param {string} template - テンプレート
   * @param {Object} parameters - パラメータ
   * @returns {string} 挿入済みテンプレート
   * @private
   */
  interpolateTemplate(template, parameters) {
    try {
      let result = template;

      for (const [key, value] of Object.entries(parameters)) {
        const placeholder = `{${key}}`;
        result = result.replace(new RegExp(placeholder, 'g'), value);
      }

      // 未置換のプレースホルダーをチェック
      const unresolved = result.match(/\{[^}]+\}/g);
      if (unresolved && unresolved.length > 0) {
        this.logger.warn('未置換のプレースホルダーが見つかりました', { unresolved });
      }

      return result;
    } catch (error) {
      this.logger.error('テンプレート挿入中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 基本的な最適化を適用
   * @param {string} prompt - プロンプト
   * @param {Object} options - オプション
   * @returns {string} 最適化されたプロンプト
   * @private
   */
  applyBasicOptimizations(prompt, options) {
    let optimized = prompt;

    // 余分な空白を削除
    if (options.trimWhitespace !== false) {
      optimized = optimized.replace(/\s+/g, ' ').trim();
    }

    // 重複する文を削除
    if (options.removeDuplicates) {
      const sentences = optimized.split(/[.!?]+/).filter(s => s.trim().length > 0);
      const uniqueSentences = [...new Set(sentences)];
      optimized = uniqueSentences.join('. ') + '.';
    }

    // 丁寧語の追加
    if (options.addPoliteLanguage) {
      if (!optimized.includes('ください') && !optimized.includes('お願い')) {
        optimized += ' よろしくお願いします。';
      }
    }

    return optimized;
  }

  /**
   * プロンプトを切り詰め
   * @param {string} prompt - プロンプト
   * @param {number} maxLength - 最大長
   * @returns {string} 切り詰められたプロンプト
   * @private
   */
  truncatePrompt(prompt, maxLength) {
    if (prompt.length <= maxLength) {
      return prompt;
    }

    // 文の境界で切り詰める
    const truncated = prompt.substring(0, maxLength);
    const lastSentenceEnd = Math.max(
      truncated.lastIndexOf('.'),
      truncated.lastIndexOf('!'),
      truncated.lastIndexOf('?'),
      truncated.lastIndexOf('。')
    );

    if (lastSentenceEnd > maxLength * 0.7) {
      return truncated.substring(0, lastSentenceEnd + 1);
    }

    return truncated + '...';
  }

  /**
   * 適用された最適化の一覧を取得
   * @param {string} original - 元のプロンプト
   * @param {string} optimized - 最適化されたプロンプト
   * @returns {Array} 最適化一覧
   * @private
   */
  getAppliedOptimizations(original, optimized) {
    const optimizations = [];

    if (original.length !== optimized.length) {
      optimizations.push('長さ調整');
    }

    if (original.replace(/\s+/g, ' ') !== optimized.replace(/\s+/g, ' ')) {
      optimizations.push('空白整理');
    }

    return optimizations;
  }
}
