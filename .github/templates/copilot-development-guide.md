# Copilot 開発ガイド

## 概要

このドキュメントは、PTA プロジェクトで GitHub Copilot を効率的に使用するためのガイドです。

## プロンプト例集

### 1. 新機能開発

#### Google Apps Script

```
PTAプロジェクトで「メール自動送信機能」を実装してください。

要件:
- 対象: PTA会員への一斉メール配信
- 入力: スプレッドシートの会員リスト
- 機能: テンプレート使用、送信結果記録
- エラー対応: 送信失敗時の再試行機能

既存パターン参考:
- src/gsuite/pta/informationDistribution.gs
- 日本語ログ出力: logInfo/logError関数使用
- 設定管理: PropertiesService使用
```

#### Excel VBA + AI 連携

```
Excel VBAでOpenAI APIを使用した「データ分析機能」を実装してください。

要件:
- 機能: 選択したセル範囲のデータを分析
- API: OpenAI GPT-4使用
- 出力: 日本語での分析結果表示
- エラー処理: API接続エラー、タイムアウト対応

既存パターン参考:
- src/excel/AI_CellProcessor.bas
- API呼び出し: CallOpenAIAPI関数パターン
- エラー表示: ShowError関数使用
```

### 2. バグ修正

```
以下のGoogle Apps Scriptエラーを修正してください:

エラー:
「参照エラー: logInfo は定義されていません」
ファイル: src/gsuite/pta/memberManagement.gs 35行目

要求:
- 既存の動作を壊さずに修正
- 日本語ログ出力機能の追加
- エラーハンドリングの改善
```

### 3. コードレビュー

```
以下のVBAコードをセキュリティ・パフォーマンス観点でレビューしてください:

[コード貼り付け]

チェックポイント:
- API キーのハードコーディング有無
- エラーハンドリングの適切性
- パフォーマンス（ループ処理、API呼び出し回数）
- 日本語化の状況
```

### 4. リファクタリング

```
以下の関数を保守性・可読性向上のためリファクタリングしてください:

[既存コード]

要件:
- 機能を変更せずに改善
- 日本語コメントの充実
- エラーハンドリングの強化
- PTA特有の命名規則適用
```

## よくある問題と解決法

### タスク実行中の典型的な課題

#### 1. 大きなタスクの分割方法

```
例：「メール配信システムの全面改修」

❌ 悪い分割例:
- メール配信システムを作り直す

✅ 良い分割例:
1. 既存システムの分析と要件整理
2. 新しいメール送信APIの調査
3. テンプレート管理機能の設計
4. 会員管理データベースの設計
5. メール送信ロジックの実装
6. エラーハンドリングの実装
7. ログ機能の実装
8. テスト実装と実行
9. 既存システムからの移行
10. ドキュメント更新
```

#### 2. 技術的負債への対応

```javascript
// レガシーコードの段階的改善例
function improveExistingCode() {
  // Phase 1: 現状の動作を保証するテストを作成
  testExistingFunctionality();

  // Phase 2: 小さな単位でのリファクタリング
  refactorSmallFunctions();

  // Phase 3: 大きな構造の改善
  refactorArchitecture();

  // Phase 4: 新機能の追加
  addNewFeatures();
}
```

### Google Apps Script

**Q: undefined function エラーが発生する**

```javascript
// A: 関数の定義順序とスコープを確認
// 共通関数は最上部で定義、またはライブラリ化

// ❌ 問題のあるパターン
function main() {
  helperFunction(); // エラー: 未定義
}

function helperFunction() {
  // 実装
}

// ✅ 修正後
function helperFunction() {
  // 実装
}

function main() {
  helperFunction(); // OK
}
```

**Q: スプレッドシートアクセスエラー**

```javascript
// A: シートの存在確認とエラーハンドリング
function getSafeSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート '${sheetName}' が見つかりません`);
  }
  return sheet;
}
```

### VBA

**Q: API 接続エラー (HTTP 401)**

```vba
' A: 認証情報の確認と適切なヘッダー設定
Private Sub ValidateAPICredentials()
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Then
        ShowError "設定エラー", "APIキーが設定されていません"
        Exit Sub
    End If
End Sub
```

**Q: 文字化けが発生する**

```vba
' A: UTF-8エンコーディングの明示的指定
http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
```

## コード品質チェックリスト

### 実装前チェック

- [ ] 要件の明確化
- [ ] 既存パターンの確認
- [ ] セキュリティ要件の把握
- [ ] エラーケースの洗い出し

### 実装中チェック

- [ ] 日本語コメントの記述
- [ ] エラーハンドリングの実装
- [ ] ログ出力の追加
- [ ] 入力値検証の実装

### 実装後チェック

- [ ] 単体テストの実行
- [ ] エラーケースのテスト
- [ ] 既存機能への影響確認
- [ ] ドキュメントの更新

## パフォーマンス最適化

### Google Apps Script

```javascript
// ✅ バッチ処理でAPI制限回避
function processInBatches(data, batchSize = 50) {
  for (let i = 0; i < data.length; i += batchSize) {
    const batch = data.slice(i, i + batchSize);
    processBatch(batch);

    if (i + batchSize < data.length) {
      Utilities.sleep(1000); // レート制限対策
    }
  }
}

// ✅ スプレッドシート読み書きの最適化
function optimizedSheetAccess() {
  const sheet = getSafeSheet("データ");
  const data = sheet.getDataRange().getValues(); // 一括取得

  // データ処理
  const processedData = data.map((row) => processRow(row));

  sheet.clear();
  sheet
    .getRange(1, 1, processedData.length, processedData[0].length)
    .setValues(processedData); // 一括書き込み
}
```

### VBA

```vba
' ✅ API呼び出しの最適化
Private Sub OptimizedAPICall()
    ' バッチ処理でリクエスト数削減
    Dim batchContent As String
    batchContent = CombineMultipleInputs()

    ' 一回のAPI呼び出しで複数処理
    Dim result As String
    result = CallOpenAIAPI(batchContent, SYSTEM_PROMPT)

    ' 結果の分割処理
    Call ProcessBatchResult(result)
End Sub
```

## セキュリティベストプラクティス

### 機密情報管理

```javascript
// ✅ 推奨: PropertiesService使用
const apiKey = PropertiesService.getScriptProperties().getProperty("API_KEY");

// ❌ 禁止: ハードコーディング
const apiKey = "sk-1234567890"; // 絶対に避ける
```

### 入力値検証

```javascript
function validateInput(input) {
  if (!input || typeof input !== "string") {
    throw new Error("無効な入力値です");
  }

  // SQLインジェクション対策
  if (input.includes("<script>") || input.includes("DROP TABLE")) {
    throw new Error("不正な文字列が検出されました");
  }

  return input.trim();
}
```

## トラブルシューティング手順

### 体系的問題解決アプローチ

#### 1. 問題の分類と優先度付け

```
問題分類マトリクス:
                     影響度
                   低    中    高
重要度   高    |  B  |  A  |  A  |
        中    |  C  |  B  |  A  |
        低    |  D  |  C  |  B  |

A: 即座に対応 (1時間以内)
B: 当日中に対応 (8時間以内)
C: 週内に対応 (1週間以内)
D: 計画的に対応 (次回メンテナンス時)
```

#### 2. 段階的デバッグ手法

```javascript
// デバッグのための段階的アプローチ
function systematicDebugging(issue) {
  console.log(`問題調査開始: ${issue.description}`);

  // Step 1: 再現性の確認
  const isReproducible = tryToReproduce(issue);
  if (!isReproducible) {
    console.log("一時的な問題の可能性 - 監視継続");
    return;
  }

  // Step 2: 最小限の再現ケース作成
  const minimalCase = createMinimalReproduction(issue);

  // Step 3: 変更履歴の確認
  const recentChanges = checkRecentChanges();

  // Step 4: 段階的な機能無効化
  const isolatedProblem = isolateIssue(minimalCase);

  // Step 5: 解決策の実装と検証
  const solution = implementSolution(isolatedProblem);
  validateSolution(solution);

  console.log("問題解決完了");
}
```

## チーム開発とコラボレーション

### 1. GitHub Copilot を使ったペアプログラミング

```
セッション準備:
1. 共通のコンテキスト設定
   - プロジェクト概要の共有
   - 作業対象コードの確認
   - 目標と制約の合意

2. ロール分担
   - ドライバー: コードを実際に書く人
   - ナビゲーター: 全体設計と品質チェック
   - Copilot: 実装パターンの提案

3. 進行方法
   - 15分単位でのロール交代
   - 各ステップでの実装方針確認
   - リアルタイムでのコードレビュー
```

### 2. 知識継承のためのドキュメント化

````markdown
## 実装記録テンプレート

### 概要

- 機能名: [機能名]
- 担当者: [担当者名]
- 実装期間: [開始日] - [完了日]

### 技術的判断

- 選択した技術: [技術名]
- 選択理由: [理由]
- 代替案と比較: [比較内容]

### 実装パターン

```javascript
// 主要な実装パターン
function examplePattern() {
  // 実装内容
}
```
````

### 学習ポイント

- 新しく学んだこと: [内容]
- 既存知識との関連: [関連性]
- 今後の応用可能性: [応用例]

### 注意点とトラップ

- 陥りやすい問題: [問題内容]
- 回避方法: [回避策]
- 参考資料: [URL 等]

````

### 3. レビュー効率化のためのChecklists
```markdown
## 機能実装レビューチェックリスト

### 基本項目
- [ ] 要件に対する実装の完全性
- [ ] 既存機能への影響なし
- [ ] エラーハンドリングの適切性
- [ ] ログ出力の日本語化
- [ ] セキュリティ要件の遵守

### PTA固有項目
- [ ] 会員情報の適切な取り扱い
- [ ] メール配信の設定確認
- [ ] スプレッドシート権限の確認
- [ ] API制限の考慮

### 品質項目
- [ ] コードの可読性
- [ ] テストカバレッジ
- [ ] パフォーマンス要件
- [ ] ドキュメント更新
````

1. **エラーログの確認**

   - Google Apps Script: 実行ログ画面
   - VBA: Debug.Print 出力

2. **権限・設定の確認**

   - API キーの有効性
   - スコープ・権限設定
   - レート制限状況

3. **段階的デバッグ**

   - 小さな単位でのテスト
   - ログ出力の追加
   - エラー再現の最小化

4. **既存コードとの比較**
   - 動作するコードとの差分確認
   - パターンの一貫性チェック
