# Copilot開発ガイド

## 概要
このドキュメントは、PTAプロジェクトでGitHub Copilotを効率的に使用するためのガイドです。

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

#### VBA + AI連携
```
Outlook VBAでOpenAI APIを使用した「メール要約機能」を実装してください。

要件:
- 機能: 選択メールの内容を要約
- API: OpenAI GPT-4使用
- 出力: 日本語での要約結果表示
- エラー処理: API接続エラー、タイムアウト対応

既存パターン参考:
- src/outlook/OutlookAI_Unified.bas
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

**Q: API接続エラー (HTTP 401)**
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
  const processedData = data.map(row => processRow(row));
  
  sheet.clear();
  sheet.getRange(1, 1, processedData.length, processedData[0].length)
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
const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');

// ❌ 禁止: ハードコーディング  
const apiKey = "sk-1234567890"; // 絶対に避ける
```

### 入力値検証
```javascript
function validateInput(input) {
  if (!input || typeof input !== 'string') {
    throw new Error('無効な入力値です');
  }
  
  // SQLインジェクション対策
  if (input.includes('<script>') || input.includes('DROP TABLE')) {
    throw new Error('不正な文字列が検出されました');
  }
  
  return input.trim();
}
```

## トラブルシューティング手順

1. **エラーログの確認**
   - Google Apps Script: 実行ログ画面
   - VBA: Debug.Print出力

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