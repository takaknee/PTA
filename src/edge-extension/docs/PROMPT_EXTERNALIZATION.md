# プロンプト設定ファイルの外だし化

## 概要

AI 機能のプロンプト設定を独立したファイル (`config/prompts.js`) に外だしして、保守性とカスタマイズ性を向上させました。

## 実装内容

### 1. プロンプト設定ファイルの作成

**ファイル**: `config/prompts.js`

#### 対応機能

- VSCode 設定解析
- メール解析
- メール返信作成
- ページ解析
- 選択テキスト解析
- Teams 転送コンテンツ
- カレンダー追加コンテンツ

#### プロンプト構造

```javascript
const VSCODE_ANALYSIS_PROMPT = {
  template: `プロンプトテンプレート（{{変数}}を含む）`,
  variables: ["pageTitle", "pageUrl", "pageContent"],
  build: function (data) {
    // 変数を実際の値に置換してプロンプトを生成
    let prompt = this.template;
    prompt = prompt.replace("{{pageTitle}}", data.pageTitle || "");
    // ...
    return prompt;
  },
};
```

### 2. プロンプト管理クラス

```javascript
class PromptManager {
    static getPrompt(type, data) {
        switch (type) {
            case 'vscode-analysis':
                return VSCODE_ANALYSIS_PROMPT.build(data);
            case 'email-analysis':
                return EMAIL_ANALYSIS_PROMPT.build(data);
            // ...
        }
    }

    static getAvailableTypes() {
        return ['vscode-analysis', 'email-analysis', ...];
    }
}
```

### 3. 背景スクリプトでの使用

**修正前（background.js 内でプロンプトをハードコーディング）**:

```javascript
const analysisPrompt = `あなたはVSCode設定の専門家です。
ページタイトル: ${data.pageTitle || ""}
...`;
```

**修正後（プロンプト設定ファイルを使用）**:

```javascript
// config/prompts.jsを読み込み
importScripts("config/prompts.js");

// プロンプトを動的生成
const analysisPrompt = PromptManager.getPrompt("vscode-analysis", {
  pageTitle: data.pageTitle || "",
  pageUrl: data.pageUrl || "",
  pageContent: pageContent,
});
```

## 利点

### 1. 保守性の向上

- ✅ プロンプトが独立したファイルで管理される
- ✅ コードとプロンプトの責任が分離される
- ✅ プロンプトの修正時にメインロジックに影響しない

### 2. カスタマイズ性の向上

- ✅ プロンプトのみを編集して機能を調整可能
- ✅ テンプレート変数システムで柔軟な構成
- ✅ 新しいプロンプトタイプの追加が容易

### 3. 再利用性の向上

- ✅ 同じプロンプトを複数の場所で使用可能
- ✅ 共通パターンの抽象化
- ✅ バージョン管理が容易

### 4. テストの容易さ

- ✅ プロンプト生成ロジックを独立してテスト可能
- ✅ 変数の置換テストが可能
- ✅ プロンプトの品質管理が向上

## 技術詳細

### ファイル構成

```
src/edge-extension/
├── config/
│   └── prompts.js          # プロンプト設定ファイル（新規作成）
├── background/
│   └── background.js       # プロンプト設定ファイルを使用
└── manifest.json           # 特別な変更は不要
```

### 変数置換システム

- テンプレート内で `{{変数名}}` 形式で変数を定義
- `build()` 関数内で実際の値に置換
- 未設定の変数は空文字列として処理

### グローバル関数化

Chrome 拡張機能の Service Worker 環境で使用するため、`globalThis`を使用してグローバルにアクセス可能にしています：

```javascript
if (typeof globalThis !== "undefined") {
  globalThis.PromptManager = PromptManager;
  globalThis.VSCODE_ANALYSIS_PROMPT = VSCODE_ANALYSIS_PROMPT;
  // ...
}
```

## 今後の拡張性

### 1. 設定ファイルベースのカスタマイズ

- ユーザー定義プロンプトの追加
- 設定画面からのプロンプト編集
- プロンプトテンプレートのインポート/エクスポート

### 2. 多言語対応

- 言語別プロンプトセットの管理
- 動的言語切り替え機能

### 3. A/B テスト対応

- 複数のプロンプトバリエーションの管理
- パフォーマンス計測機能

## 対応状況

- [x] VSCode 設定解析プロンプトの外だし
- [x] プロンプト管理クラスの実装
- [x] 背景スクリプトでの使用実装
- [ ] 他の AI 機能への適用（今後の作業）
- [ ] 設定画面での編集機能（将来的な改善）

---

**更新日**: 2025 年 6 月 17 日
**対象ファイル**:

- `config/prompts.js` (新規作成)
- `background/background.js` (プロンプト生成部分を修正)
