/**
 * HTML セキュリティ・サニタイゼーション モジュール
 * PTA Edge Extension 共通ライブラリ
 * 
 * 用途：
 * - HTMLコンテンツの安全なサニタイゼーション
 * - XSS攻撃の防止
 * - 安全なテキスト抽出
 * 
 * セキュリティ基準：
 * - OWASP Top 10 対応
 * - CodeQL セキュリティ要件準拠
 * - Chrome Extension セキュリティポリシー準拠
 * 
 * 【重要】セキュリティ方針変更：
 * - DOMPurifyを必須要件とする
 * - フォールバックサニタイザーは廃止（セキュリティリスク軽減）
 * - DOMPurifyが利用できない場合はエラーを投げる
 * 
 * Service Worker対応：
 * - importScripts() で読み込み可能
 * - globalThis.PTASanitizer で利用可能
 * - ES Modules は使用しない（Service Worker制限のため）
 */

// Service Worker環境での読み込み確認
(() => {
  'use strict';

  try {
    console.log('HTMLサニタイゼーションモジュール: 初期化開始');

    // グローバルスコープの確認
    if (typeof globalThis === 'undefined') {
      throw new Error('globalThis がサポートされていません');
    }

    console.log('HTMLサニタイゼーションモジュール: 環境確認完了');

  } catch (error) {
    console.error('HTMLサニタイゼーションモジュール: 初期化エラー', error);
    return;
  }
})();

/**
 * DOMPurifyベースの統一サニタイザー実装
 * セキュリティを最優先とし、DOMPurifyを必須要件とする
 * フォールバック実装は廃止（セキュリティリスク軽減のため）
 */
function createUnifiedSanitizer() {
  'use strict';

  // DOMPurifyの利用可能性を確認
  const isDOMPurifyAvailable = () => {
    return typeof globalThis !== 'undefined' &&
      typeof globalThis.DOMPurify !== 'undefined' &&
      typeof globalThis.DOMPurify.sanitize === 'function';
  };

  /**
   * DOMPurifyを使用した安全なHTMLサニタイゼーション
   * @param {string} html - サニタイズするHTML文字列
   * @returns {string} - サニタイズされた安全な文字列
   * @throws {Error} DOMPurifyが利用できない場合
   */
  function extractSafeTextWithDOMPurify(html) {
    if (!html || typeof html !== 'string') return '';

    // DOMPurifyの必須チェック
    if (!isDOMPurifyAvailable()) {
      throw new Error('DOMPurifyが必要です。ライブラリを読み込んでください。');
    }

    try {
      // DOMPurifyの設定
      const config = {
        ALLOWED_TAGS: [], // すべてのタグを除去してテキストのみ抽出
        ALLOWED_ATTR: [],
        KEEP_CONTENT: true, // タグ内のテキストコンテンツは保持
        RETURN_DOM: false,
        RETURN_DOM_FRAGMENT: false,
        RETURN_TRUSTED_TYPE: false
      };

      // DOMPurifyでサニタイズ実行
      const sanitized = globalThis.DOMPurify.sanitize(html, config);

      // 追加の正規化処理
      return sanitized
        .replace(/\s+/g, ' ')      // 連続空白を統合
        .trim()                     // 前後の空白を除去
        .substring(0, 10000);       // 長さ制限

    } catch (error) {
      console.error('DOMPurifyサニタイゼーション中にエラー:', error);
      throw new Error(`サニタイゼーション処理に失敗しました: ${error.message}`);
    }
  }

  /**
   * 統一されたextractSafeText関数
   * DOMPurifyを必須要件とし、セキュリティを最優先
   * @param {string} html - サニタイズするHTML文字列
   * @returns {string} - サニタイズされた安全な文字列
   * @throws {Error} DOMPurifyが利用できない場合
   */
  function extractSafeText(html) {
    console.log('統一サニタイザー: extractSafeText実行開始');

    if (!isDOMPurifyAvailable()) {
      const errorMsg = 'DOMPurifyが必要です。セキュリティ上の理由により、DOMPurifyなしでの実行はサポートされていません。';
      console.error(errorMsg);
      throw new Error(errorMsg);
    }

    console.log('DOMPurify使用可能 - DOMPurifyベースのサニタイゼーションを実行');
    return extractSafeTextWithDOMPurify(html);
  }

  /**
   * 高速なHTMLタグ除去（軽量版）
   * DOMPurifyを使用した安全な実装
   * @param {string} html - 処理するHTML文字列
   * @returns {string} - 処理された安全な文字列
   * @throws {Error} DOMPurifyが利用できない場合
   */
  function fastStripTags(html) {
    if (!html || typeof html !== 'string') return '';

    if (!isDOMPurifyAvailable()) {
      throw new Error('DOMPurifyが必要です。fastStripTagsにはDOMPurifyが必須です。');
    }

    try {
      // DOMPurifyの高速設定
      const config = {
        ALLOWED_TAGS: [],
        ALLOWED_ATTR: [],
        KEEP_CONTENT: true,
        RETURN_DOM: false
      };

      const sanitized = globalThis.DOMPurify.sanitize(html, config);

      return sanitized
        .replace(/\s+/g, ' ')
        .trim()
        .substring(0, 1000);

    } catch (error) {
      console.error('fastStripTags処理中にエラー:', error);
      throw new Error(`高速タグ除去処理に失敗しました: ${error.message}`);
    }
  }

  // 公開インターフェース
  return {
    extractSafeText,
    fastStripTags,
    isDOMPurifyAvailable,
    // 後方互換性のために旧関数名も提供
    sanitizeHTML: extractSafeText,
    stripTags: fastStripTags
  };
}

// グローバルな統一サニタイザーインスタンスを作成
const PTAUnifiedSanitizer = createUnifiedSanitizer();

// PTASanitizerとして公開（既存コードとの互換性）
if (typeof globalThis !== 'undefined') {
  globalThis.PTASanitizer = PTAUnifiedSanitizer;
  console.log('✅ PTASanitizer（統一版）がグローバルスコープで利用可能になりました');

  // 初期化テスト（DOMPurify必須チェック含む）
  try {
    if (!globalThis.PTASanitizer.isDOMPurifyAvailable()) {
      console.warn('⚠️ DOMPurifyが利用できません。このライブラリはDOMPurifyを必要とします。');
    } else {
      const testResult = globalThis.PTASanitizer.extractSafeText('<p>テスト<script>alert("XSS")</script></p>');
      console.log('✅ 統一サニタイザー初期化テスト成功:', testResult);
    }
  } catch (error) {
    console.error('❌ 統一サニタイザー初期化テスト失敗:', error);
  }
} else {
  console.error('❌ globalThis が利用できません - PTASanitizerを登録できませんでした');
}

// Service Worker環境での利用のためにモジュールをエクスポート
// （後方互換性のために古い関数も含める）

/**
 * レガシー互換性のためのラッパー関数群
 * 既存コードが動作し続けるように提供
 */

// 旧関数の統一サニタイザーへの移行
function extractSafeText(html) {
  return PTAUnifiedSanitizer.extractSafeText(html);
}

function fastStripTags(html) {
  return PTAUnifiedSanitizer.fastStripTags(html);
}

function sanitizeHTML(html) {
  return PTAUnifiedSanitizer.extractSafeText(html);
}

// 旧コードで参照される可能性のある関数名もサポート
function stripTags(html) {
  return PTAUnifiedSanitizer.fastStripTags(html);
}

// モジュールとしてのエクスポート（Service Worker用）
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    extractSafeText,
    fastStripTags,
    sanitizeHTML,
    stripTags,
    PTAUnifiedSanitizer
  };
}

console.log('✅ HTML サニタイゼーションモジュール（統一版）の初期化完了');
