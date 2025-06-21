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
 * 危険なHTMLタグのリスト
 * XSS攻撃やコード実行に使用される可能性があるタグ
 */
const DANGEROUS_TAGS = [
  'script', 'style', 'iframe', 'frame', 'frameset', 'noframes',
  'object', 'embed', 'applet', 'meta', 'link', 'base',
  'form', 'input', 'button', 'select', 'textarea', 'option',
  'img', 'audio', 'video', 'source', 'track',
  'svg', 'math', 'canvas'
];

/**
 * 危険な属性とプロトコルのリスト
 * JavaScriptコード実行や外部リソース読み込みに使用される可能性がある
 */
const DANGEROUS_ATTRIBUTES = [
  'onabort', 'onblur', 'onchange', 'onclick', 'ondblclick',
  'onerror', 'onfocus', 'onkeydown', 'onkeypress', 'onkeyup',
  'onload', 'onmousedown', 'onmousemove', 'onmouseout', 'onmouseover',
  'onmouseup', 'onreset', 'onresize', 'onselect', 'onsubmit',
  'onunload', 'onafterprint', 'onbeforeprint', 'onbeforeunload',
  'onhashchange', 'onmessage', 'onoffline', 'ononline', 'onpagehide',
  'onpageshow', 'onpopstate', 'onstorage', 'onunload',
  'javascript:', 'vbscript:', 'data:', 'blob:', 'file:'
];

/**
 * HTMLエンティティのマッピング
 * 安全な文字への変換用
 */
const HTML_ENTITY_MAP = {
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&#x27;': "'",
  '&#x2F;': '/',
  '&#x60;': '`',
  '&#x3D;': '='
};

/**
 * セキュリティ強化版：複数文字パターンのサニタイゼーション
 * CodeQL脆弱性「Incomplete multi-character sanitization」への対応
 * 
 * @param {string} input - サニタイズ対象の文字列
 * @returns {string} - サニタイズされた文字列
 */
function sanitizeMultiCharacterPatterns(input) {
  if (!input || typeof input !== 'string') return '';

  let sanitized = input;
  const maxIterations = 100; // 無限ループ防止

  // 各パターンを個別に処理（より安全な方式）
  const patterns = [{
    name: 'HTMLコメント',
    regex: /<!--[\s\S]*?-->/g,
    // 専用の安全な関数を使用してHTMLコメントを除去
    safeAlternative: (str) => removeHTMLCommentsSafely(str)
  }, {
    name: 'CDATA',
    regex: /<!\[CDATA\[[\s\S]*?\]\]>/g,
    // 専用の安全な関数を使用してCDATAを除去
    safeAlternative: (str) => removeCDATASafely(str)
  },
  {
    name: '処理命令',
    regex: /<\?[\s\S]*?\?>/g,
    // 専用の安全な関数を使用して処理命令を除去
    safeAlternative: (str) => removeProcessingInstructionsSafely(str)
  },
  {
    name: 'DOCTYPE宣言',
    regex: /<!DOCTYPE[\s\S]*?>/gi,
    // 専用の安全な関数を使用してDOCTYPE宣言を除去
    safeAlternative: (str) => removeDOCTYPESafely(str)
  }
  ];

  // 各パターンに対して安全な除去処理を実行
  patterns.forEach(pattern => {
    try {
      // まず従来の方式で基本的な除去
      let previous;
      let localIteration = 0;

      do {
        previous = sanitized;
        sanitized = sanitized.replace(pattern.regex, '');
        localIteration++;
      } while (sanitized !== previous && localIteration < maxIterations);

      // 安全な代替方式で追加処理
      sanitized = pattern.safeAlternative(sanitized);

    } catch (error) {
      console.warn(`パターン ${pattern.name} の処理中にエラー:`, error);
      // エラー時は最も安全な方式：危険な文字を単純置換
      sanitized = sanitized.replace(/[<>]/g, '');
    }
  });

  return sanitized;
}

/**
 * 危険なタグの除去
 * ネストしたタグや異常な形式にも対応
 * 
 * @param {string} html - 処理対象のHTML文字列
 * @returns {string} - 危険なタグが除去されたHTML文字列
 */
function removeDangerousTags(html) {
  if (!html || typeof html !== 'string') return '';

  let sanitized = html;

  DANGEROUS_TAGS.forEach(tag => {
    try {
      // 開始タグと終了タグのペア（ネスト対応）
      const regex1 = new RegExp(`<${tag}\\s*[^>]*>[\\s\\S]*?<\\/${tag}\\s*>`, 'gi');
      sanitized = sanitized.replace(regex1, '');

      // 自己完結タグ
      const regex2 = new RegExp(`<${tag}\\s*[^>]*\\s*/?>`, 'gi');
      sanitized = sanitized.replace(regex2, '');

      // 不正な形式のタグ（属性なし）
      const regex3 = new RegExp(`<\\s*\\/?\\s*${tag}\\s*>`, 'gi');
      sanitized = sanitized.replace(regex3, '');

    } catch (error) {
      console.warn(`タグ ${tag} の除去処理中にエラー:`, error);
    }
  });

  return sanitized;
}

/**
 * 危険な属性の除去
 * イベントハンドラーや危険なプロトコルを含む属性を除去
 * 
 * @param {string} html - 処理対象のHTML文字列
 * @returns {string} - 危険な属性が除去されたHTML文字列
 */
function removeDangerousAttributes(html) {
  if (!html || typeof html !== 'string') return '';

  return html.replace(/<[^>]+>/g, (match) => {
    const lowerMatch = match.toLowerCase();

    // 危険な属性が含まれているかチェック
    for (const attr of DANGEROUS_ATTRIBUTES) {
      if (lowerMatch.includes(attr)) {
        return ' '; // 危険な属性を含むタグは完全除去
      }
    }    // style属性の特別処理（CSSインジェクション対策）
    if (lowerMatch.includes('style=')) {
      return removeStyleAttributesSafely(match);
    }

    return ' '; // 最終的にはタグを除去してテキスト化
  });
}

/**
 * HTMLエンティティの安全なデコード
 * 数値エンティティも含めて適切に処理
 * 
 * @param {string} text - デコード対象のテキスト
 * @returns {string} - デコードされたテキスト
 */
function decodeHTMLEntities(text) {
  if (!text || typeof text !== 'string') return '';

  let decoded = text;

  try {
    // 名前付きエンティティのデコード
    Object.keys(HTML_ENTITY_MAP).forEach(entity => {
      const regex = new RegExp(entity.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
      decoded = decoded.replace(regex, HTML_ENTITY_MAP[entity]);
    });

    // 数値エンティティのデコード（安全な範囲のみ）
    decoded = decoded.replace(/&#(\d+);/g, (match, num) => {
      const code = parseInt(num, 10);
      // 安全な範囲の文字のみデコード（基本的なASCII文字）
      if (code >= 32 && code <= 126) {
        return String.fromCharCode(code);
      }
      return ' '; // 安全でない文字は空白に置換
    });

    // 16進数エンティティのデコード（安全な範囲のみ）
    decoded = decoded.replace(/&#x([0-9a-fA-F]+);/g, (match, hex) => {
      const code = parseInt(hex, 16);
      if (code >= 32 && code <= 126) {
        return String.fromCharCode(code);
      }
      return ' ';
    });

  } catch (error) {
    console.warn('HTMLエンティティデコード中にエラー:', error);
    // エラー時は元のテキストを返す
    return text;
  }

  return decoded;
}

/**
 * テキストの正規化
 * 空白文字や制御文字の適切な処理
 * 
 * @param {string} text - 正規化対象のテキスト
 * @returns {string} - 正規化されたテキスト
 */
function normalizeText(text) {
  if (!text || typeof text !== 'string') return '';

  return text
    .replace(/[\r\n\t]/g, ' ')      // 改行・タブを空白に
    .replace(/\s+/g, ' ')           // 連続空白を単一空白に
    .replace(/^\s+|\s+$/g, '')      // 前後の空白を除去
    .substring(0, 10000);           // 長すぎるテキストの制限
}

/**
 * メインのHTMLサニタイゼーション関数
 * 安全なテキスト抽出を行う
 * 
 * @param {string} html - サニタイズ対象のHTML文字列
 * @returns {string} - 安全にサニタイズされたテキスト
 */
function extractSafeText(html) {
  if (!html || typeof html !== 'string') return '';

  try {
    let sanitized = html;

    // 1. 危険なタグの完全除去
    sanitized = removeDangerousTags(sanitized);

    // 2. コメントと特殊なセクションの除去
    sanitized = sanitizeMultiCharacterPatterns(sanitized);

    // 3. 危険な属性の除去
    sanitized = removeDangerousAttributes(sanitized);

    // 4. HTMLエンティティの安全なデコード
    sanitized = decodeHTMLEntities(sanitized);

    // 5. テキストの正規化
    sanitized = normalizeText(sanitized);

    return sanitized;

  } catch (error) {
    console.error('HTMLサニタイゼーション処理中にエラー:', error);

    // セキュリティ重視のフォールバック処理
    try {
      let fallbackSanitized = html
        .replace(/[<>]/g, '')           // すべての < > を除去
        .replace(/&[^;]*;/g, ' ')       // HTMLエンティティを空白に
        .replace(/javascript:/gi, '')   // javascript: プロトコル除去
        .replace(/vbscript:/gi, '')     // vbscript: プロトコル除去
        .replace(/data:/gi, '')         // data: プロトコル除去
        .replace(/[\r\n\t]/g, ' ')      // 制御文字を空白に
        .replace(/\s+/g, ' ')           // 連続空白を統合
        .trim()
        .substring(0, 10000);           // 長さ制限

      console.warn('フォールバック処理によるサニタイゼーションを実行しました');
      return fallbackSanitized;

    } catch (fallbackError) {
      console.error('フォールバック処理も失敗:', fallbackError);
      return ''; // 最後の手段：空文字列を返す
    }
  }
}

/**
 * 高速なHTMLタグ除去（軽量版）
 * 基本的なHTMLタグのみを除去する軽量版
 * 
 * @param {string} html - 処理対象のHTML文字列
 * @returns {string} - HTMLタグが除去されたテキスト
 */
function stripHTMLTags(html) {
  if (!html || typeof html !== 'string') return '';

  try {
    return html
      .replace(/<[^>]*>/g, ' ')       // すべてのHTMLタグを除去
      .replace(/\s+/g, ' ')           // 連続空白を統合
      .trim();
  } catch (error) {
    console.error('HTMLタグ除去中にエラー:', error);
    return html.replace(/[<>]/g, ''); // 最低限の処理
  }
}

/**
 * サニタイゼーション機能のテスト関数
 * 開発・デバッグ時に使用
 * 
 * @returns {boolean} - すべてのテストが合格した場合はtrue
 */
function testSanitization() {
  const testCases = [
    {
      name: '基本的なHTMLコメント',
      input: '<!-- comment -->text',
      expected: 'text'
    },
    {
      name: 'ネストされたHTMLコメント攻撃',
      input: '<!<!--- comment --->>text',
      expected: 'text'
    },
    {
      name: 'CDATA攻撃',
      input: '<![CDATA[<script>alert(1)</script>]]>text',
      expected: 'text'
    },
    {
      name: '処理命令攻撃',
      input: '<?php echo "test"; ?>text',
      expected: 'text'
    },
    {
      name: '複合攻撃',
      input: '<!--<![CDATA[--><script>alert(1)</script><!--]]>-->text',
      expected: 'text'
    },
    {
      name: 'JavaScript プロトコル',
      input: '<a href="javascript:alert(1)">text</a>',
      expected: 'text'
    },
    {
      name: 'イベントハンドラー属性',
      input: '<div onclick="alert(1)">text</div>',
      expected: 'text'
    },
    {
      name: 'HTMLエンティティ',
      input: '&lt;script&gt;alert(1)&lt;/script&gt;',
      expected: '<script>alert(1)</script>'
    }
  ];

  console.log('=== HTMLサニタイゼーションテスト開始 ===');

  let allPassed = true;

  testCases.forEach((testCase, index) => {
    try {
      const result = extractSafeText(testCase.input);
      const passed = result.trim() === testCase.expected;

      if (!passed) allPassed = false;

      console.log(`テスト ${index + 1}: ${testCase.name}`);
      console.log(`  入力: ${testCase.input}`);
      console.log(`  期待値: ${testCase.expected}`);
      console.log(`  結果: ${result}`);
      console.log(`  判定: ${passed ? '✅ 合格' : '❌ 不合格'}`);
    } catch (error) {
      console.error(`テスト ${index + 1} でエラー:`, error);
      allPassed = false;
    }
  });

  console.log('=== HTMLサニタイゼーションテスト終了 ===');
  console.log(`総合結果: ${allPassed ? '✅ 全テスト合格' : '❌ 一部テスト不合格'}`);

  return allPassed;
}

/**
 * HTMLコメント専用の安全な除去処理
 * CodeQL脆弱性「js/bad-tag-filter」への対応
 * 不正な形式のコメント終了タグ（--!>など）も適切に処理
 * 
 * @param {string} input - 処理対象の文字列
 * @returns {string} - HTMLコメントが安全に除去された文字列
 */
function removeHTMLCommentsSafely(input) {
  if (!input || typeof input !== 'string') return '';

  let result = input;
  const maxIterations = 100; // 無限ループ防止
  let changed = true;
  let iter = 0;

  while (changed && iter < maxIterations) {
    const before = result;
    // Consolidated and repeated sanitization logic
    result = result.replace(/<!--[\s\S]*?-->|<!--|--!?>|-!>|--\w>|<!-+|-+>|<!-[\s\S]*?>/g, '');
    changed = (result !== before);
    iter++;
  }

  return result;
}

/**
 * CDATA専用の安全な除去処理
 * CodeQL脆弱性「js/bad-tag-filter」への対応
 * 不正な形式のCDATA終了タグも適切に処理
 * 
 * @param {string} input - 処理対象の文字列
 * @returns {string} - CDATAが安全に除去された文字列
 */
function removeCDATASafely(input) {
  if (!input || typeof input !== 'string') return '';

  let result = input;
  const maxIterations = 100; // 無限ループ防止
  let changed = true;
  let iter = 0;

  while (changed && iter < maxIterations) {
    const before = result.length;

    // 1. 正規のCDATAを除去（<![CDATA[...]]>）
    result = result.replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '');

    // 2. 不正な形式のCDATA開始タグを除去
    result = result.replace(/<!\[CDATA\[/g, '');

    // 3. 不正な形式のCDATA終了タグを除去
    result = result.replace(/\]\]>/g, '');
    result = result.replace(/\]!>/g, ''); // ]!> のような変形
    result = result.replace(/\]\w>/g, ''); // ]a>, ]b> など

    // 4. その他の潜在的な不正パターンを除去
    result = result.replace(/<!\[\w*\[/g, ''); // <![CDATA[, <![DATA[, etc.
    result = result.replace(/\]+>/g, ''); // ]>, ]]>, ]]]>, etc.

    changed = (result.length !== before);
    iter++;
  }

  return result;
}

/**
 * 処理命令専用の安全な除去処理
 * CodeQL脆弱性「js/bad-tag-filter」への対応
 * 不正な形式の処理命令終了タグも適切に処理
 * 
 * @param {string} input - 処理対象の文字列
 * @returns {string} - 処理命令が安全に除去された文字列
 */
function removeProcessingInstructionsSafely(input) {
  if (!input || typeof input !== 'string') return '';

  let result = input;
  const maxIterations = 100; // 無限ループ防止
  let changed = true;
  let iter = 0;

  while (changed && iter < maxIterations) {
    const before = result.length;

    // 1. 正規の処理命令を除去（<?...?>）
    result = result.replace(/<\?[\s\S]*?\?>/g, '');

    // 2. 不正な形式の処理命令開始タグを除去
    result = result.replace(/<\?/g, '');

    // 3. 不正な形式の処理命令終了タグを除去
    result = result.replace(/\?>/g, '');
    result = result.replace(/\?!>/g, ''); // ?!> のような変形
    result = result.replace(/\?\w>/g, ''); // ?a>, ?b> など

    // 4. その他の潜在的な不正パターンを除去
    result = result.replace(/<\?\w*/g, ''); // <?php, <?xml, etc.
    result = result.replace(/\?+>/g, ''); // ?>, ??>, ???>, etc.

    changed = (result.length !== before);
    iter++;
  }

  return result;
}

/**
 * DOCTYPE宣言専用の安全な除去処理
 * CodeQL脆弱性「js/bad-tag-filter」への対応
 * 不正な形式のDOCTYPE宣言も適切に処理
 * 
 * @param {string} input - 処理対象の文字列
 * @returns {string} - DOCTYPE宣言が安全に除去された文字列
 */
function removeDOCTYPESafely(input) {
  if (!input || typeof input !== 'string') return '';

  let result = input;
  const maxIterations = 100; // 無限ループ防止
  let changed = true;
  let iter = 0;

  while (changed && iter < maxIterations) {
    const before = result.length;

    // 1. 正規のDOCTYPE宣言を除去（<!DOCTYPE...>）
    result = result.replace(/<!DOCTYPE[\s\S]*?>/gi, '');

    // 2. 不正な形式のDOCTYPE開始を除去
    result = result.replace(/<!DOCTYPE/gi, '');

    // 3. 孤立した > タグの除去（より安全な方式）
    // 単純な > 除去ではなく、DOCTYPE関連のみに限定
    result = result.replace(/<!doctype[\s\S]*?>/gi, ''); // 小文字版
    result = result.replace(/<!Doctype[\s\S]*?>/gi, ''); // 混在版

    // 4. その他の潜在的な不正パターンを除去
    result = result.replace(/<!doc\w*/gi, ''); // <!doctype, <!document, etc.

    changed = (result.length !== before);
    iter++;
  }

  return result;
}

/**
 * Style属性専用の安全な除去処理
 * CodeQL脆弱性「js/bad-tag-filter」への対応
 * 複雑な引用符パターンや不正なエスケープにも対応
 * 
 * @param {string} input - 処理対象の文字列
 * @returns {string} - style属性が安全に除去された文字列
 */
function removeStyleAttributesSafely(input) {
  if (!input || typeof input !== 'string') return '';

  let result = input;
  const maxIterations = 100; // 無限ループ防止
  let changed = true;
  let iter = 0;

  while (changed && iter < maxIterations) {
    const before = result.length;

    // 1. 正規のstyle属性を除去（様々な引用符パターンに対応）
    result = result.replace(/style\s*=\s*"[^"]*"/gi, '');
    result = result.replace(/style\s*=\s*'[^']*'/gi, '');
    result = result.replace(/style\s*=\s*[^"'\s>][^\s>]*/gi, ''); // 引用符なし

    // 2. 不正なエスケープを含むstyle属性
    result = result.replace(/style\s*=\s*"[^"]*\\["'][^"]*"/gi, '');
    result = result.replace(/style\s*=\s*'[^']*\\['"][^']*'/gi, '');

    // 3. 混在した引用符パターン
    result = result.replace(/style\s*=\s*"[^"]*'[^"]*"/gi, '');
    result = result.replace(/style\s*=\s*'[^']*"[^']*'/gi, '');    // 4. CSSインジェクション対策：危険なCSS関数の除去
    result = result.replace(/expression\s*\([^)]*\)/gi, '');
    result = result.replace(/javascript\s*:/gi, '');
    result = result.replace(/data\s*:/gi, '');
    result = result.replace(/url\s*\(\s*["']?javascript:/gi, '');

    // 5. style属性の不正な形式の除去
    result = result.replace(/style\s*=/gi, ''); // 孤立したstyle=

    changed = (result.length !== before);
    iter++;
  }

  return result;
}

/**
 * セキュリティ強化版：危険なパターンの包括的検知
 * CodeQL「Bad HTML filtering regexp」脆弱性への対応
 * 
 * @param {string} html - 検査対象のHTMLコンテンツ
 * @returns {Object} - 検知結果とリスク評価
 */
function detectAdvancedSecurityPatterns(html) {
  if (!html || typeof html !== 'string') {
    return { safe: true, riskLevel: 'none', detectedPatterns: [] };
  }

  // より堅牢な正規表現パターン（CodeQL要件準拠）
  const securityPatterns = [
    // script タグ - 開始・終了を個別検知（スペース対応）
    {
      name: 'script 開始タグ検知',
      pattern: /<script(?:\s+[^>]*)?>/gi,
      severity: 'critical',
      description: 'JavaScript実行可能なscript開始タグ'
    },
    {
      name: 'script 終了タグ検知',
      pattern: /<\/script(?:\s*)>/gi,
      severity: 'critical',
      description: 'script終了タグ（スペース含む）'
    },
    // JavaScript プロトコル
    {
      name: 'JavaScript プロトコル',
      pattern: /(?:javascript|vbscript|data|blob|file)\s*:/gi,
      severity: 'high',
      description: '危険なURLプロトコル'
    },
    // イベントハンドラー（包括的）
    {
      name: 'イベントハンドラー',
      pattern: /\bon\w+\s*=(?:\s*["']?[^"'>]*["']?|[^>\s]*)/gi,
      severity: 'high',
      description: 'HTML要素のイベントハンドラー'
    },
    // iframe タグ
    {
      name: 'iframe 開始タグ',
      pattern: /<iframe(?:\s+[^>]*)?>/gi,
      severity: 'high',
      description: '外部コンテンツ埋め込み可能なiframe'
    },
    {
      name: 'iframe 終了タグ',
      pattern: /<\/iframe(?:\s*)>/gi,
      severity: 'high',
      description: 'iframe終了タグ'
    },
    // その他の危険なタグ
    {
      name: 'object/embed タグ',
      pattern: /<(?:object|embed)(?:\s+[^>]*)?>/gi,
      severity: 'medium',
      description: 'プラグイン実行可能なタグ'
    },
    {
      name: 'form/input タグ',
      pattern: /<(?:form|input|button|select|textarea)(?:\s+[^>]*)?>/gi,
      severity: 'medium',
      description: 'ユーザー入力可能なフォーム要素'
    },
    // 高度な回避パターン
    {
      name: 'HTMLエンコード回避',
      pattern: /&(?:#(?:x[0-9a-f]+|[0-9]+)|[a-z]+);/gi,
      severity: 'low',
      description: 'HTMLエンティティエンコーディング'
    },
    {
      name: 'Base64疑似パターン',
      pattern: /(?:data:[\w\/]+;base64,|btoa|atob)[A-Za-z0-9+\/=]{20,}/gi,
      severity: 'medium',
      description: 'Base64エンコードされた可能性のあるデータ'
    }
  ];

  const detectedPatterns = [];
  let highestSeverity = 'none';

  // パターンマッチング実行
  securityPatterns.forEach(pattern => {
    try {
      const matches = html.match(pattern.pattern);
      if (matches && matches.length > 0) {
        detectedPatterns.push({
          name: pattern.name,
          count: matches.length,
          severity: pattern.severity,
          description: pattern.description,
          examples: matches.slice(0, 2).map(match => 
            match.length > 50 ? match.substring(0, 50) + '...' : match
          )
        });

        // 最高リスクレベルを更新
        const severityLevels = { 'none': 0, 'low': 1, 'medium': 2, 'high': 3, 'critical': 4 };
        if (severityLevels[pattern.severity] > severityLevels[highestSeverity]) {
          highestSeverity = pattern.severity;
        }
      }
    } catch (error) {
      console.error(`セキュリティパターン検知エラー (${pattern.name}):`, error);
    }
  });

  // 結果の評価
  const riskLevel = highestSeverity === 'none' ? 'none' :
                   highestSeverity === 'critical' ? 'critical' :
                   highestSeverity === 'high' ? 'high' :
                   highestSeverity === 'medium' ? 'medium' : 'low';

  return {
    safe: detectedPatterns.length === 0,
    riskLevel: riskLevel,
    detectedPatterns: detectedPatterns,
    summary: {
      totalPatterns: detectedPatterns.length,
      criticalCount: detectedPatterns.filter(p => p.severity === 'critical').length,
      highCount: detectedPatterns.filter(p => p.severity === 'high').length,
      mediumCount: detectedPatterns.filter(p => p.severity === 'medium').length,
      lowCount: detectedPatterns.filter(p => p.severity === 'low').length
    }
  };
}

// エクスポート（Service Worker環境対応）
const sanitizerAPI = {
  extractSafeText,
  stripHTMLTags,
  sanitizeMultiCharacterPatterns,
  removeDangerousTags,
  removeDangerousAttributes,
  decodeHTMLEntities,
  normalizeText,
  testSanitization,
  detectAdvancedSecurityPatterns
};

// Service Worker環境で直接グローバルスコープに設定
try {
  if (typeof globalThis !== 'undefined') {
    globalThis.PTASanitizer = sanitizerAPI;
    console.log('HTMLサニタイゼーションモジュール: グローバルスコープに正常に設定されました');
  } else {
    throw new Error('globalThis が利用できません');
  }
} catch (error) {
  console.error('HTMLサニタイゼーションモジュール: グローバル設定エラー', error);
}

// ブラウザ環境での下位互換性
try {
  if (typeof window !== 'undefined') {
    window.PTASanitizer = sanitizerAPI;
    console.log('HTMLサニタイゼーションモジュール: window に設定されました');
  }
} catch (error) {
  console.warn('HTMLサニタイゼーションモジュール: window 設定エラー', error);
}

// Node.js環境での互換性
try {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = sanitizerAPI;
    console.log('HTMLサニタイゼーションモジュール: module.exports に設定されました');
  }
} catch (error) {
  console.warn('HTMLサニタイゼーションモジュール: module.exports 設定エラー', error);
}

// 初期化完了ログ
console.log('HTMLサニタイゼーションモジュール: 初期化完了');
console.log('利用可能な機能:', Object.keys(sanitizerAPI));
