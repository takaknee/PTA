/*
 * AI業務支援ツール Edge拡張機能 - プロンプト管理サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { Logger } from '../lib/utils.js';

/**
 * AIプロンプト管理サービス
 */
export class PromptService {
  constructor() {
    this.prompts = new Map();
    this.initializePrompts();
  }

  /**
   * プロンプトを初期化
   */
  initializePrompts() {
    // メール解析プロンプト
    this.registerPrompt('emailAnalysis', {
      system: `あなたは優秀なメール解析アシスタントです。
受信したメールの内容を分析し、以下の形式で日本語で回答してください：

## 📧 メール要約
- **差出人**: [差出人名]
- **件名**: [件名]
- **重要度**: [高/中/低]
- **カテゴリ**: [業務/個人/緊急/情報/その他]

## 📝 内容要約
[メールの主要な内容を3-5行で要約]

## 🎯 アクションアイテム
[必要なアクションがあれば箇条書きで記載、なければ「アクションなし」]

## 💡 推奨対応
[返信が必要かどうか、対応期限の目安など]

専門的な内容でも分かりやすく、簡潔に分析してください。`,

      template: (emailData) => `
以下のメールを分析してください：

**差出人**: ${emailData.from || '不明'}
**件名**: ${emailData.subject || '件名なし'}
**受信日時**: ${emailData.date || '不明'}

**本文**:
${emailData.body || 'メール本文が取得できませんでした'}
`
    });

    // ページ解析プロンプト
    this.registerPrompt('pageAnalysis', {
      system: `あなたは優秀なWeb情報解析アシスタントです。
Webページの内容を分析し、以下の形式で日本語で回答してください：

## 🌐 ページ情報
- **タイトル**: [ページタイトル]
- **URL**: [ページURL]
- **サイト種別**: [企業サイト/ニュース/ブログ/ECサイト/その他]

## 📊 内容要約
[ページの主要な内容を5-7行で要約]

## 🔑 キーポイント
[重要なポイントを3-5個の箇条書きで]

## 🔗 関連情報
[関連リンクや参考情報があれば記載]

## 💼 ビジネス活用
[業務での活用方法や注意点があれば記載]

読みやすく、実用的な形で情報をまとめてください。`,

      template: (pageData) => `
以下のWebページを分析してください：

**URL**: ${pageData.url || '不明'}
**タイトル**: ${pageData.title || 'タイトル不明'}

**ページ内容**:
${pageData.content || 'ページ内容が取得できませんでした'}
`
    });

    // 選択テキスト解析プロンプト
    this.registerPrompt('selectionAnalysis', {
      system: `あなたは優秀なテキスト解析アシスタントです。
選択されたテキストを分析し、以下の形式で日本語で回答してください：

## 📝 テキスト分析
- **種別**: [文書/記事/コード/データ/その他]
- **言語**: [日本語/英語/その他]
- **文字数**: [概算文字数]

## 📋 内容要約
[選択されたテキストの内容を3-5行で要約]

## 🎯 重要なポイント
[重要な情報やキーワードを箇条書きで]

## 💡 解釈・補足
[内容の解釈や補足説明があれば記載]

## 🔄 関連アクション
[このテキストに対して推奨されるアクション]

明確で理解しやすい分析を提供してください。`,

      template: (selectionData) => `
以下の選択されたテキストを分析してください：

**ソースURL**: ${selectionData.url || '不明'}
**ページタイトル**: ${selectionData.pageTitle || '不明'}

**選択テキスト**:
${selectionData.selectedText || '選択テキストが取得できませんでした'}
`
    });

    // メール作成プロンプト
    this.registerPrompt('emailComposition', {
      system: `あなたは優秀なメール作成アシスタントです。
以下の指示に基づいて、適切なビジネスメールを日本語で作成してください：

## メール作成ガイドライン
- 件名は簡潔で内容が分かりやすいものにする
- 宛先に応じた適切な敬語を使用
- 内容は論理的で読みやすい構成にする
- 必要に応じて箇条書きを活用
- 結びの挨拶を含める

## 出力形式
**件名**: [メール件名]

**本文**:
[メール本文をそのまま記載]

丁寧で分かりやすいメールを作成してください。`,

      template: (requestData) => `
以下の条件でメールを作成してください：

**メール種別**: ${requestData.type || '一般'}
**宛先**: ${requestData.recipient || '相手先'}
**目的**: ${requestData.purpose || '目的不明'}

**詳細指示**:
${requestData.details || '詳細な指示をご提供ください'}

${requestData.referenceEmail ? `**参考メール**:\n${requestData.referenceEmail}` : ''}
`
    });

    // 翻訳プロンプト
    this.registerPrompt('translation', {
      system: `あなたは優秀な翻訳アシスタントです。
以下のテキストを指定された言語に翻訳してください：

## 翻訳ガイドライン
- 自然で読みやすい翻訳を心がける
- 文脈を考慮した適切な訳語を選択
- 専門用語は正確に翻訳
- 翻訳不可能な固有名詞はそのまま記載

翻訳結果のみを出力してください。`,

      template: (translationData) => `
以下のテキストを${translationData.targetLanguage || '日本語'}に翻訳してください：

${translationData.text || '翻訳対象テキストが指定されていません'}
`
    });

    // URL抽出プロンプト
    this.registerPrompt('urlExtraction', {
      system: `あなたは優秀な情報抽出アシスタントです。
テキストからURLを抽出し、以下の形式で整理してください：

## 🔗 抽出されたURL一覧

各URLについて以下の情報を提供：
- **URL**: [完全なURL]
- **リンクテキスト**: [元のリンクテキスト]
- **種別**: [公式サイト/記事/資料/その他]
- **説明**: [URLの内容の簡単な説明]

見つからない場合は「URLが見つかりませんでした」と回答してください。`,

      template: (extractionData) => `
以下のテキストからURLを抽出してください：

${extractionData.text || 'テキストが提供されていません'}
`
    });

    Logger.info('プロンプト初期化完了', { count: this.prompts.size });
  }

  /**
   * プロンプトを登録
   * @param {string} key プロンプトキー
   * @param {Object} promptData プロンプトデータ
   */
  registerPrompt(key, promptData) {
    this.prompts.set(key, {
      system: promptData.system,
      template: promptData.template,
      metadata: promptData.metadata || {}
    });

    Logger.debug('プロンプト登録', { key });
  }

  /**
   * プロンプトを取得
   * @param {string} key プロンプトキー
   * @param {Object} data テンプレートデータ
   * @returns {Object} プロンプト情報
   */
  getPrompt(key, data = {}) {
    const promptData = this.prompts.get(key);

    if (!promptData) {
      Logger.warn('プロンプトが見つかりません', { key });
      return this.getDefaultPrompt(data);
    }

    try {
      const userPrompt = typeof promptData.template === 'function'
        ? promptData.template(data)
        : promptData.template || JSON.stringify(data);

      return {
        systemPrompt: promptData.system,
        userPrompt: userPrompt.trim(),
        metadata: promptData.metadata
      };

    } catch (error) {
      Logger.error('プロンプト生成エラー', { key, error: error.message });
      return this.getDefaultPrompt(data);
    }
  }

  /**
   * デフォルトプロンプトを取得
   * @param {Object} data データ
   * @returns {Object} デフォルトプロンプト
   */
  getDefaultPrompt(data) {
    return {
      systemPrompt: 'あなたは優秀なAIアシスタントです。日本語で分かりやすく回答してください。',
      userPrompt: JSON.stringify(data),
      metadata: {}
    };
  }

  /**
   * メール解析用プロンプトを生成
   * @param {Object} emailData メールデータ
   * @returns {Object} プロンプト情報
   */
  createEmailAnalysisPrompt(emailData) {
    return this.getPrompt('emailAnalysis', emailData);
  }

  /**
   * ページ解析用プロンプトを生成
   * @param {Object} pageData ページデータ
   * @returns {Object} プロンプト情報
   */
  createPageAnalysisPrompt(pageData) {
    return this.getPrompt('pageAnalysis', pageData);
  }

  /**
   * 選択テキスト解析用プロンプトを生成
   * @param {Object} selectionData 選択データ
   * @returns {Object} プロンプト情報
   */
  createSelectionAnalysisPrompt(selectionData) {
    return this.getPrompt('selectionAnalysis', selectionData);
  }

  /**
   * メール作成用プロンプトを生成
   * @param {Object} requestData リクエストデータ
   * @returns {Object} プロンプト情報
   */
  createEmailCompositionPrompt(requestData) {
    return this.getPrompt('emailComposition', requestData);
  }

  /**
   * 翻訳用プロンプトを生成
   * @param {Object} translationData 翻訳データ
   * @returns {Object} プロンプト情報
   */
  createTranslationPrompt(translationData) {
    return this.getPrompt('translation', translationData);
  }

  /**
   * URL抽出用プロンプトを生成
   * @param {Object} extractionData 抽出データ
   * @returns {Object} プロンプト情報
   */
  createUrlExtractionPrompt(extractionData) {
    return this.getPrompt('urlExtraction', extractionData);
  }

  /**
   * カスタムプロンプトを作成
   * @param {string} systemPrompt システムプロンプト
   * @param {string} userPrompt ユーザープロンプト
   * @param {Object} metadata メタデータ
   * @returns {Object} プロンプト情報
   */
  createCustomPrompt(systemPrompt, userPrompt, metadata = {}) {
    return {
      systemPrompt: systemPrompt || 'あなたは優秀なAIアシスタントです。日本語で回答してください。',
      userPrompt: userPrompt || '',
      metadata: metadata
    };
  }

  /**
   * 登録されているプロンプトの一覧を取得
   * @returns {Array} プロンプトキー一覧
   */
  getAvailablePrompts() {
    return Array.from(this.prompts.keys());
  }

  /**
   * プロンプトの詳細情報を取得
   * @param {string} key プロンプトキー
   * @returns {Object} プロンプト詳細
   */
  getPromptInfo(key) {
    const promptData = this.prompts.get(key);

    if (!promptData) {
      return null;
    }

    return {
      key: key,
      hasTemplate: typeof promptData.template === 'function',
      metadata: promptData.metadata,
      systemPromptLength: promptData.system.length
    };
  }

  /**
   * プロンプトを削除
   * @param {string} key プロンプトキー
   * @returns {boolean} 削除成功可否
   */
  removePrompt(key) {
    const existed = this.prompts.has(key);
    this.prompts.delete(key);

    if (existed) {
      Logger.debug('プロンプト削除', { key });
    }

    return existed;
  }

  /**
   * すべてのプロンプトをクリア
   */
  clearPrompts() {
    const count = this.prompts.size;
    this.prompts.clear();
    Logger.info('プロンプト全削除', { count });
  }

  /**
   * プロンプトテンプレートの妥当性をチェック
   * @param {Function} template テンプレート関数
   * @param {Object} testData テストデータ
   * @returns {Object} 検証結果
   */
  validateTemplate(template, testData = {}) {
    try {
      if (typeof template !== 'function') {
        return { isValid: false, error: 'テンプレートは関数である必要があります' };
      }

      const result = template(testData);

      if (typeof result !== 'string') {
        return { isValid: false, error: 'テンプレートは文字列を返す必要があります' };
      }

      return { isValid: true, result };

    } catch (error) {
      return { isValid: false, error: error.message };
    }
  }
}
