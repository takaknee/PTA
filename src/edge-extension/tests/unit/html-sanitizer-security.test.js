/**
 * HTMLサニタイザー セキュリティテスト
 * 
 * このテストファイルは、html-sanitizer.jsのセキュリティ脆弱性を確認するためのものです。
 * 特にCodeQLが指摘したHTMLコメント処理の脆弱性（js/bad-tag-filter）をテストします。
 */

// HTMLサニタイザーの読み込み（実際のテスト実行時に必要）
// const { removeHTMLCommentsSafely, sanitizeMultiCharacterPatterns } = require('../../infrastructure/html-sanitizer.js');

/**
 * HTMLコメント処理のセキュリティテストケース
 * CodeQL脆弱性「js/bad-tag-filter」への対応確認
 */
const HTML_COMMENT_SECURITY_TEST_CASES = [
  // 正規のHTMLコメント
  {
    input: '<!-- 正規のコメント -->テキスト',
    expected: 'テキスト',
    description: '正規のHTMLコメントの除去'
  },
  {
    input: 'テキスト1<!-- 複数行\nコメント -->テキスト2',
    expected: 'テキスト1テキスト2',
    description: '複数行HTMLコメントの除去'
  },

  // CodeQLが指摘した不正な形式のコメント終了タグ
  {
    input: '<!--悪意のあるコード--!>',
    expected: '',
    description: '不正なコメント終了タグ --!> の安全な処理'
  },
  {
    input: '<!--スクリプト-!>',
    expected: '',
    description: '不正なコメント終了タグ -!> の安全な処理'
  },
  {
    input: '<!--攻撃コード--a>',
    expected: '',
    description: '不正なコメント終了タグ --a> の安全な処理'
  },

  // 複雑な攻撃パターン
  {
    input: '<!-攻撃-->正常テキスト<!---攻撃->',
    expected: '正常テキスト',
    description: '複数の不正コメントパターンの処理'
  },
  {
    input: 'テキスト<!--コメント1--><!--コメント2--!>テキスト2',
    expected: 'テキストテキスト2',
    description: '混在するコメントパターンの処理'
  },
  {
    input: '<!---攻撃--->',
    expected: '',
    description: '過剰なハイフンを含む不正コメント'
  },

  // ネストしたコメント攻撃
  {
    input: '<!--外側<!--内側-->外側続き-->',
    expected: '',
    description: 'ネストしたコメント構造の安全な処理'
  },

  // 不完全なコメント構造
  {
    input: '<!--不完全',
    expected: '',
    description: '不完全なコメント開始タグの処理'
  },
  {
    input: '不完全-->',
    expected: '不完全',
    description: '不完全なコメント終了タグの処理'
  }
];

/**
 * XSS攻撃を模したHTMLコメント内コード実行テスト
 */
const XSS_HTML_COMMENT_TEST_CASES = [
  {
    input: '<!--<script>alert("XSS")</script>-->',
    expected: '',
    description: 'コメント内スクリプトタグの完全除去'
  },
  {
    input: '<!--<img src=x onerror=alert(1)>--!>',
    expected: '',
    description: 'コメント内イベントハンドラーの安全な処理'
  },
  {
    input: '<!--javascript:alert(1)--!>',
    expected: '',
    description: 'コメント内javascript:プロトコルの処理'
  }
];

/**
 * パフォーマンステスト用の大量データ
 */
const PERFORMANCE_TEST_CASES = [
  {
    input: '<!--' + 'A'.repeat(10000) + '-->',
    description: '大量データを含むHTMLコメントの処理性能'
  },
  {
    input: Array(1000).fill('<!--test--!>').join(''),
    description: '大量の不正コメントタグの処理性能'
  }
];

/**
 * テスト実行関数（モックアップ）
 * 実際のテスト実行時には、適切なテストフレームワークを使用
 */
function runSecurityTests() {
  console.log('=== HTMLコメント セキュリティテスト開始 ===');

  // 基本的なセキュリティテスト
  HTML_COMMENT_SECURITY_TEST_CASES.forEach((testCase, index) => {
    console.log(`テスト ${index + 1}: ${testCase.description}`);
    console.log(`入力: "${testCase.input}"`);
    console.log(`期待値: "${testCase.expected}"`);

    // 実際のテスト実行（要実装）
    // const result = removeHTMLCommentsSafely(testCase.input);
    // console.log(`結果: "${result}"`);
    // console.log(`判定: ${result === testCase.expected ? 'PASS' : 'FAIL'}`);
    console.log('---');
  });

  // XSS攻撃テスト
  console.log('\n=== XSS攻撃防止テスト ===');
  XSS_HTML_COMMENT_TEST_CASES.forEach((testCase, index) => {
    console.log(`XSSテスト ${index + 1}: ${testCase.description}`);
    console.log(`入力: "${testCase.input}"`);
    console.log(`期待値: "${testCase.expected}"`);
    console.log('---');
  });

  console.log('\n=== テスト完了 ===');
}

/**
 * 回帰テスト用の既知の脆弱性パターン
 * 将来的な修正で再発しないことを確認
 */
const REGRESSION_TEST_CASES = [
  {
    cveId: 'CodeQL-js/bad-tag-filter',
    input: '<!--攻撃コード--!>',
    expectedBehavior: '完全に除去される',
    description: 'CodeQL指摘の不正コメント終了タグ処理'
  }
];

// テスト結果のエクスポート（実際のテスト環境で使用）
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    HTML_COMMENT_SECURITY_TEST_CASES,
    XSS_HTML_COMMENT_TEST_CASES,
    PERFORMANCE_TEST_CASES,
    REGRESSION_TEST_CASES,
    runSecurityTests
  };
}
