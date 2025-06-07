# Copilot 指示ファイル（効率化強化版）

## プロジェクト概要
このプロジェクトは **PTA情報配信システム** のための Google Workspace および Microsoft 365 自動化スクリプト集です。

### 主要技術スタック
- **Google Apps Script (GAS)**: Gmail自動化、スプレッドシート操作、カレンダー管理
- **VBA (Outlook)**: メール処理、AI連携（OpenAI API）、検索フォルダ分析
- **JavaScript/Node.js**: ビルドツール、テスト、CI/CD

### アーキテクチャ
```
PTA/
├── src/
│   ├── gsuite/          # Google Apps Script ファイル
│   │   ├── gmail/       # メール処理自動化
│   │   └── pta/         # PTA業務特化機能
│   ├── outlook/         # Outlook VBA マクロ
│   └── excel/           # Excel VBA AI連携
├── docs/                # 技術文書・仕様書
└── .github/            # CI/CD・品質管理
```

## コーディング方針とパターン

### 1. 日本語中心開発
- **すべてのやり取りを日本語で行う**
- **コード内のコメントは日本語で記述**
- **エラーメッセージやログ出力も日本語で実装**
- **変数名や関数名の説明も日本語で提供**
- **ドキュメントは日本語で作成**

### 2. 命名規則とコードスタイル
```javascript
// ✅ 推奨パターン
function sendEmailToMembers(memberList, subject, body) {
  // メンバーリストへのメール送信処理
  try {
    memberList.forEach(member => {
      // 各メンバーへの送信ロジック
    });
    logInfo('メール送信完了');
  } catch (error) {
    logError('メール送信エラー', error.message);
    throw error;
  }
}

// ✅ VBA推奨パターン
Private Sub ProcessEmailWithAI()
    On Error GoTo ErrorHandler
    
    ' メール処理のメインロジック
    Dim emailContent As String
    emailContent = GetSelectedEmailContent()
    
    ' AI処理実行
    Dim result As String
    result = CallOpenAIAPI(emailContent)
    
    ShowMessage result, "AI処理結果"
    Exit Sub
    
ErrorHandler:
    ShowError "メール処理中にエラーが発生しました", Err.Description
End Sub
```

### 3. エラーハンドリングパターン
```javascript
// Google Apps Script 標準パターン
function robustFunction() {
  try {
    // メイン処理
    return successResult;
  } catch (error) {
    // 日本語エラーログ
    console.error(`処理エラー: ${error.message}`);
    Logger.log(`エラー詳細: ${error.stack}`);
    
    // ユーザー向けメッセージ
    throw new Error(`処理に失敗しました: ${error.message}`);
  }
}
```

## プロジェクト固有パターンとベストプラクティス

### Google Apps Script 特有パターン
```javascript
// 設定値の安全な取得
function getConfigValue(key, defaultValue = null) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value || defaultValue;
  } catch (error) {
    logError(`設定取得エラー: ${key}`, error.message);
    return defaultValue;
  }
}

// スプレッドシート操作の標準パターン
function safeSheetOperation(sheetName, operation) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート '${sheetName}' が見つかりません`);
  }
  
  return operation(sheet);
}

// バッチ処理パターン（大量データ対応）
function processMembersInBatches(members, batchSize = 50) {
  const results = [];
  for (let i = 0; i < members.length; i += batchSize) {
    const batch = members.slice(i, i + batchSize);
    results.push(...processBatch(batch));
    
    // レート制限対策
    if (i + batchSize < members.length) {
      Utilities.sleep(1000); // 1秒待機
    }
  }
  return results;
}
```

### VBA + OpenAI API パターン
```vba
' API呼び出しの標準パターン
Private Function CallOpenAIAPI(content As String, systemPrompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' リクエスト構築
    Dim requestBody As String
    requestBody = BuildAPIRequest(content, systemPrompt)
    
    ' API呼び出し実行
    http.Open "POST", OPENAI_API_ENDPOINT, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & OPENAI_API_KEY
    http.send requestBody
    
    ' レスポンス処理
    If http.Status = 200 Then
        CallOpenAIAPI = ParseAPIResponse(http.responseText)
    Else
        Err.Raise vbObjectError + 1000, , "API呼び出しエラー: " & http.Status
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "API処理エラー", "詳細: " & Err.Description
    CallOpenAIAPI = ""
End Function

' 設定管理パターン
Private Function GetAPISetting(key As String, defaultValue As String) As String
    ' レジストリまたは設定ファイルからの安全な取得
    On Error Resume Next
    GetAPISetting = GetSetting("OutlookAI", "Settings", key, defaultValue)
    On Error GoTo 0
End Function
```

## セキュリティとコンプライアンス

### 機密情報の安全な管理
```javascript
// ✅ 推奨: PropertiesService使用
const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

// ❌ 禁止: ハードコーディング
const apiKey = "sk-1234567890abcdef"; // 絶対に避ける
```

```vba
' ✅ 推奨: 定数での管理（開発時）+ 将来的な外部設定
Const OPENAI_API_KEY As String = "YOUR_API_KEY_HERE" ' プレースホルダー
Const OPENAI_API_ENDPOINT As String = "YOUR_ENDPOINT_HERE"

' ✅ 推奨: 設定ファイルからの読み込み
Private Function GetAPIKey() As String
    GetAPIKey = GetAPISetting("APIKey", "YOUR_API_KEY_HERE")
End Function
```

### データ保護要件
- **個人情報**: メールアドレス、氏名等は必要最小限の使用
- **API制限**: レート制限の遵守（1分間60リクエスト以下）
- **ログ記録**: 個人情報を含まない形での実行ログ記録
- **エラー処理**: 機密情報を含まないエラーメッセージ

## 推奨ライブラリとAPI

### Google Apps Script
```javascript
// 基本サービス
SpreadsheetApp    // スプレッドシート操作
DriveApp         // ファイル・フォルダ管理  
GmailApp         // メール送受信
CalendarApp      // カレンダー操作
FormApp          // フォーム作成・回答処理

// ユーティリティ
Logger           // ログ出力
Utilities        // 各種ユーティリティ
PropertiesService // 設定値管理
LockService      // 排他制御
HtmlService      // UI作成
```

### Microsoft 365 / VBA
```vba
' Outlook VBA
Application.ActiveInspector    ' アクティブメール
Application.Session           ' Outlookセッション
MSXML2.XMLHTTP               ' HTTP通信（API呼び出し）

' Excel VBA  
Application.Workbooks        ' ワークブック操作
Range                       ' セル範囲操作
Worksheet                   ' ワークシート操作
```

### 外部API
- **OpenAI API**: GPT-4によるテキスト処理
- **Microsoft Graph API**: M365データアクセス  
- **Google Workspace Admin SDK**: 管理者機能

## Copilot効率化のための具体的指示

### コード生成時の必須要件
1. **日本語での応答とコメント**
2. **エラーメッセージも日本語で実装**  
3. **ログ出力も日本語で記録**
4. **関数や変数の説明も日本語で提供**
5. **既存のコードパターンとの整合性確保**

### プロンプト例とテンプレート

#### 新機能開発時
```
このPTAプロジェクトで[機能名]を実装してください。
- 技術: [GAS/VBA/JavaScript]
- 要件: [具体的な要件]
- 既存パターン: src/[関連ディレクトリ]のコードを参考に
- エラーハンドリング: 日本語メッセージで実装
- ログ: logInfo/logError関数を使用
```

#### バグ修正時  
```
以下のエラーを修正してください:
- ファイル: [ファイルパス]
- エラー内容: [エラーメッセージ]
- 期待される動作: [正常な動作]
- 制約: 既存の動作を壊さずに最小限の変更で修正
```

#### コードレビュー時
```
以下のコードをレビューしてください:
- セキュリティ面での問題点
- パフォーマンスの改善点  
- 日本語化の改善提案
- PTA特有の要件への適合性
```

### よくあるタスクパターン

1. **メール自動送信機能**: `src/gsuite/pta/informationDistribution.gs` 参考
2. **スプレッドシート操作**: `src/gsuite/pta/memberManagement.gs` 参考  
3. **VBA + AI連携**: `src/outlook/OutlookAI_Unified.bas` 参考
4. **設定管理**: `AI_ConfigManager.bas` パターン使用
5. **エラーハンドリング**: 既存の`ShowError`/`logError`パターン踏襲

## トラブルシューティングとFAQ

### よくある問題と解決法

#### Google Apps Script関連
**Q: 「参照エラー: xxx は定義されていません」**
```javascript
// A: 依存関数の定義順序を確認、または共通モジュールの読み込み
// 例: PTA_CONFIG, logInfo, logError などは共通関数として定義が必要
```

**Q: スプレッドシートアクセスエラー**  
```javascript
// A: 権限確認とシート存在チェック
function safeGetSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート '${sheetName}' が見つかりません`);
  }
  return sheet;
}
```

#### VBA + OpenAI API関連
**Q: API接続エラー (401 Unauthorized)**
```vba
' A: APIキーとエンドポイントの確認
Private Sub ValidateAPISettings()
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Then
        ShowError "API設定エラー", "APIキーが設定されていません"
        Exit Sub
    End If
End Sub
```

**Q: タイムアウトエラー**
```vba  
' A: タイムアウト値の調整とレート制限対策
http.setTimeouts 0, REQUEST_TIMEOUT * 1000, REQUEST_TIMEOUT * 1000, REQUEST_TIMEOUT * 1000
```

### パフォーマンス最適化

#### 大量データ処理
```javascript
// バッチ処理でAPI制限回避
function processLargeDataset(data, batchSize = 50) {
  const results = [];
  for (let i = 0; i < data.length; i += batchSize) {
    const batch = data.slice(i, i + batchSize);
    results.push(...processBatch(batch));
    
    // レート制限対策の待機
    if (i + batchSize < data.length) {
      Utilities.sleep(1000);
    }
  }
  return results;
}
```

## 開発ワークフロー

### 新機能開発手順
1. **要件分析**: 既存機能との整合性確認
2. **設計**: 該当するパターン（GAS/VBA）の選択  
3. **実装**: テンプレートコードの活用
4. **テスト**: 単体テスト + 統合テスト
5. **ドキュメント**: README.md更新

### コードレビューポイント
- ✅ 日本語コメント・メッセージの充実度
- ✅ エラーハンドリングの網羅性
- ✅ セキュリティ要件（機密情報漏洩防止）
- ✅ 既存パターンとの整合性
- ✅ パフォーマンス（特にAPI呼び出し回数）

## 禁止事項と注意点
- ❌ 英語でのコメントや説明
- ❌ 機密情報のハードコーディング  
- ❌ セキュリティリスクのあるコード
- ❌ 著作権に問題のあるコード
- ❌ 既存の動作を破壊する変更
- ❌ 不必要な依存関係の追加
