# Copilot 指示ファイル（効率化強化版）

## プロジェクト概要
このプロジェクトは **PTA情報配信システム** のための Google Workspace および Microsoft 365 自動化スクリプト集です。

### 主要技術スタック
- **Google Apps Script (GAS)**: Gmail自動化、スプレッドシート操作、カレンダー管理
- **VBA (Outlook)**: メール処理、AI連携（OpenAI API）、検索フォルダ分析
- **VSTO (C#)**: Outlook アドイン、M365展開対応、保守性向上
- **JavaScript/Node.js**: ビルドツール、テスト、CI/CD

### アーキテクチャ
```
PTA/
├── src/
│   ├── gsuite/          # Google Apps Script ファイル
│   │   ├── gmail/       # メール処理自動化
│   │   └── pta/         # PTA業務特化機能
│   ├── outlook/         # Outlook VBA マクロ
│   ├── outlook-vsto/    # Outlook VSTO アドイン (C#)
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

// ✅ VSTO推奨パターン (C#)
public async Task ProcessEmailWithAIAsync()
{
    try
    {
        _logger.LogInformation("メール解析処理を開始します");
        
        // 選択されたメールアイテムの取得
        var mailItem = GetSelectedMailItem();
        if (mailItem == null)
        {
            throw new InvalidOperationException("メールが選択されていません");
        }
        
        // AI処理の実行（非同期）
        var result = await _emailAnalysisService.AnalyzeEmailAsync(mailItem);
        
        // 結果の表示
        ShowAnalysisResult(result);
        
        _logger.LogInformation("メール解析処理が完了しました");
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "メール解析処理中にエラーが発生しました");
        MessageBox.Show($"メール解析中にエラーが発生しました:\n{ex.Message}", 
                       "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
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

### VSTO + OpenAI API パターン
```csharp
// API呼び出しの標準パターン（非同期）
public async Task<string> CallOpenAIAPIAsync(string content, string systemPrompt)
{
    try
    {
        _logger.LogDebug("OpenAI API呼び出しを開始します");
        
        // リクエストボディの構築
        var requestBody = new
        {
            model = _configService.GetOpenAIModel(),
            messages = new[]
            {
                new { role = "system", content = systemPrompt },
                new { role = "user", content = content }
            },
            max_tokens = _configService.GetMaxTokens(),
            temperature = 0.7
        };
        
        var json = JsonConvert.SerializeObject(requestBody);
        var httpContent = new StringContent(json, Encoding.UTF8, "application/json");
        
        // API設定の取得
        var apiKey = _configService.GetOpenAIApiKey();
        var endpoint = _configService.GetOpenAIEndpoint();
        
        _httpClient.DefaultRequestHeaders.Authorization = 
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
        
        // API呼び出し実行
        var response = await _httpClient.PostAsync(endpoint, httpContent);
        
        if (response.IsSuccessStatusCode)
        {
            var responseContent = await response.Content.ReadAsStringAsync();
            var result = ParseOpenAIResponse(responseContent);
            
            _logger.LogDebug("OpenAI API呼び出しが成功しました");
            return result;
        }
        else
        {
            var errorContent = await response.Content.ReadAsStringAsync();
            throw new OpenAIException($"API呼び出しエラー: {response.StatusCode} - {errorContent}");
        }
    }
    catch (HttpRequestException ex)
    {
        _logger.LogError(ex, "HTTP通信エラーが発生しました");
        throw new OpenAIException("API通信中にエラーが発生しました", ex);
    }
    catch (Exception ex)
    {
        _logger.LogError(ex, "OpenAI API呼び出し中に予期しないエラーが発生しました");
        throw;
    }
}

// 設定管理パターン（依存関係注入）
public class ConfigurationService
{
    private readonly ILogger<ConfigurationService> _logger;
    
    public string GetOpenAIApiKey()
    {
        try
        {
            // Windows Credential Manager からの安全な取得
            var credential = CredentialManager.ReadCredential("OutlookPTA_OpenAI");
            return credential?.Password ?? ConfigurationManager.AppSettings["OpenAIApiKey"];
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "OpenAI APIキーの取得中に警告が発生しました");
            return ConfigurationManager.AppSettings["OpenAIApiKey"];
        }
    }
    
    public string GetOpenAIEndpoint()
    {
        return ConfigurationManager.AppSettings["OpenAIEndpoint"] ?? 
               throw new ConfigurationException("OpenAIエンドポイントが設定されていません");
    }
}
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

## タスクベース開発のベストプラクティス

### タスク分解と進行の指針

#### 1. 効果的なタスク分解手法
```
大きなタスクを以下の段階で分解してください:

1. **要件分析フェーズ**
   - 機能要件の明確化
   - 既存システムへの影響分析
   - 技術的制約の確認

2. **設計フェーズ**
   - アーキテクチャレベルの設計
   - インターフェース設計
   - エラーハンドリング戦略

3. **実装フェーズ**
   - 小さな単位での段階的実装
   - 各段階でのテスト実行
   - 漸進的な機能追加

4. **検証フェーズ**
   - 単体テストの実行
   - 統合テストの確認
   - 既存機能への影響確認
```

#### 2. GitHub WorkflowでのCopilot活用

##### Pull Requestベースの開発
```
1. **ブランチ作成時**
   - 機能別の明確なブランチ名（feature/email-automation等）
   - 初期コミットでの作業計画コメント

2. **開発中**
   - 小さな変更での頻繁なコミット
   - コミットメッセージは日本語で具体的に記述
   - 段階的なコードレビュー要求

3. **PR作成時**
   - 変更内容の詳細な説明（日本語）
   - テスト結果の報告
   - 既存機能への影響説明
```

### コード生成時の必須要件
1. **日本語での応答とコメント**
2. **エラーメッセージも日本語で実装**  
3. **ログ出力も日本語で記録**
4. **関数や変数の説明も日本語で提供**
5. **既存のコードパターンとの整合性確保**
6. **段階的な実装アプローチの採用**

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
4. **VSTO + AI連携**: `src/outlook-vsto/Core/Services/EmailAnalysisService.cs` 参考
5. **VSTO UI統合**: `src/outlook-vsto/UI/Ribbon/PTARibbon.cs` 参考
6. **設定管理**: `AI_ConfigManager.bas` (VBA) / `ConfigurationService.cs` (VSTO) パターン使用
7. **エラーハンドリング**: 既存の`ShowError`/`logError`パターン踏襲

## テストとバリデーション戦略

### 段階的テストアプローチ

#### 1. 開発前のテスト設計
```javascript
// テスト設計の例：メール送信機能
function testEmailSending() {
  const testCases = [
    {
      description: "正常ケース：有効な会員リストでの送信",
      input: { memberList: validMembers, subject: "テスト件名", body: "テスト本文" },
      expected: { success: true, sentCount: validMembers.length }
    },
    {
      description: "異常ケース：空の会員リスト",
      input: { memberList: [], subject: "テスト件名", body: "テスト本文" },
      expected: { success: false, error: "会員リストが空です" }
    },
    {
      description: "異常ケース：無効なメールアドレス",
      input: { memberList: invalidMembers, subject: "テスト件名", body: "テスト本文" },
      expected: { success: false, error: "無効なメールアドレスが含まれています" }
    }
  ];
  
  testCases.forEach(testCase => {
    console.log(`テスト実行: ${testCase.description}`);
    // テスト実装
  });
}
```

#### 2. 実装中のバリデーション
```javascript
// 段階的な機能確認
function validateImplementation() {
  try {
    // Step 1: 基本機能の確認
    console.log("Step 1: 基本機能テスト開始");
    testBasicFunctionality();
    
    // Step 2: エラーハンドリングの確認
    console.log("Step 2: エラーハンドリングテスト開始");
    testErrorHandling();
    
    // Step 3: 統合テスト
    console.log("Step 3: 統合テスト開始");
    testIntegration();
    
    console.log("すべてのテストが完了しました");
  } catch (error) {
    console.error(`テスト失敗: ${error.message}`);
    throw error;
  }
}
```

#### 3. VBA向けテストパターン
```vba
' VBAテスト実装例
Private Sub TestAPIIntegration()
    On Error GoTo ErrorHandler
    
    ' テストケース1: 正常なAPI呼び出し
    Console.Print "テスト1: 正常なAPI呼び出し"
    Dim result As String
    result = CallOpenAIAPI("テストコンテンツ", "システムプロンプト")
    
    If result <> "" Then
        Console.Print "✓ API呼び出し成功"
    Else
        Console.Print "✗ API呼び出し失敗"
    End If
    
    ' テストケース2: エラーハンドリング
    Console.Print "テスト2: 無効なAPIキーでのエラーハンドリング"
    ' 実装...
    
    Exit Sub
    
ErrorHandler:
    Console.Print "テストエラー: " & Err.Description
End Sub
```

### コード品質チェックリスト

#### 実装前チェック
- [ ] タスクの明確な定義と分解
- [ ] 既存パターンとの整合性確認
- [ ] セキュリティ要件の把握
- [ ] テストケースの設計
- [ ] エラーケースの洗い出し

#### 実装中チェック
- [ ] 段階的な実装（小さな単位での確認）
- [ ] 日本語コメントの充実
- [ ] エラーハンドリングの実装
- [ ] ログ出力の追加
- [ ] 入力値検証の実装
- [ ] 各段階でのテスト実行

#### 実装後チェック
- [ ] 全テストケースの実行と合格確認
- [ ] 既存機能への影響なし確認
- [ ] パフォーマンス要件の満足
- [ ] セキュリティ要件の遵守確認
- [ ] ドキュメントの更新
- [ ] コードレビューの完了

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

### GitHub連携によるタスクベース開発

#### 1. Issue-Driven Development
```
1. **Issue作成**
   - 具体的なタスク定義（日本語）
   - 受け入れ条件の明確化
   - 関連する既存機能の特定
   - 推定工数と優先度の設定

2. **ブランチ戦略**
   - 機能別ブランチ: feature/[機能名]-[issue番号]
   - バグ修正: bugfix/[問題内容]-[issue番号]
   - hotfix: hotfix/[緊急修正内容]

3. **コミット規約**
   - 日本語での明確なコミットメッセージ
   - 変更内容の具体的説明
   - 関連Issueの参照（#issue番号）
```

#### 2. Pull Request ベストプラクティス
```
PR作成時のチェックリスト:
- [ ] 明確なPRタイトル（日本語）
- [ ] 変更内容の詳細説明
- [ ] テスト結果の報告
- [ ] スクリーンショット（UI変更時）
- [ ] 破壊的変更の有無確認
- [ ] セキュリティ影響の評価
- [ ] パフォーマンス影響の評価
- [ ] ドキュメント更新の確認
```

#### 3. コードレビュー指針
```
レビュー観点:
1. **機能性**: 要件通りの実装か
2. **保守性**: 理解しやすく変更しやすいか
3. **セキュリティ**: 脆弱性はないか
4. **パフォーマンス**: 性能要件を満たすか
5. **一貫性**: 既存パターンとの整合性
6. **テスト**: 適切なテストが含まれているか
```

### 新機能開発手順
1. **要件分析**: 既存機能との整合性確認
2. **設計**: 該当するパターン（GAS/VBA）の選択  
3. **実装**: テンプレートコードの活用
4. **テスト**: 単体テスト + 統合テスト
5. **ドキュメント**: README.md更新
6. **デプロイ**: 段階的なリリース計画

### Copilot活用のワークフロー統合

#### GitHub Copilot Chat の効果的な使用
```
// タスク開始時のコンテキスト設定
このPTAプロジェクトで「[具体的な機能名]」を実装します。

プロジェクト情報:
- 技術スタック: Google Apps Script / VBA
- 既存パターン: src/[関連ディレクトリ] 参照
- 要件: [具体的な要件]
- 制約条件: [技術的・業務的制約]

段階的に実装を進めてください。
```

#### コード生成時の品質確保
```
1. **初期実装**
   - 基本機能の骨格作成
   - エラーハンドリングの基本形実装
   - ログ出力の追加

2. **段階的改善**
   - 入力値検証の強化
   - エラーメッセージの日本語化
   - パフォーマンス最適化

3. **最終検証**
   - テストケース実行
   - セキュリティチェック
   - 既存機能への影響確認
```

### コードレビューポイント
- ✅ 日本語コメント・メッセージの充実度
- ✅ エラーハンドリングの網羅性
- ✅ セキュリティ要件（機密情報漏洩防止）
- ✅ 既存パターンとの整合性
- ✅ パフォーマンス（特にAPI呼び出し回数）

## 継続的改善とメンテナンス

### 品質メトリクスとモニタリング

#### 1. コード品質指標
```javascript
// 品質チェック自動化の例
function checkCodeQuality() {
  const metrics = {
    testCoverage: calculateTestCoverage(),
    codeComplexity: analyzeCyclomaticComplexity(),
    duplicateCode: findDuplicateBlocks(),
    securityIssues: scanSecurityVulnerabilities(),
    performanceIssues: analyzePerformanceBottlenecks()
  };
  
  logInfo(`品質メトリクス:`, metrics);
  
  if (metrics.testCoverage < 80) {
    logError('テストカバレッジが不足しています');
  }
  
  return metrics;
}
```

#### 2. 定期的なコードレビュー
```
月次レビュー項目:
- [ ] セキュリティパッチの適用状況
- [ ] パフォーマンス劣化の確認
- [ ] エラーログの傾向分析
- [ ] ユーザーフィードバックの反映
- [ ] 技術的負債の評価
- [ ] ドキュメントの更新状況
```

### 学習と知識共有

#### 1. ベストプラクティスの蓄積
```markdown
## 学習ログ テンプレート

### 実装した機能
- 機能名: [機能名]
- 実装日: [日付]
- 技術: [GAS/VBA/JavaScript]

### 学んだポイント
- 技術的課題: [課題内容]
- 解決方法: [解決策]
- 改善点: [次回への改善提案]

### 再利用可能パターン
```
[コードパターン]
```

### 今後の注意点
- [注意事項]
```

#### 2. トラブルシューティング知識ベース
```javascript
// 問題解決パターンの蓄積
const troubleshootingPatterns = {
  "API接続エラー": {
    symptoms: ["401 Unauthorized", "タイムアウト", "レート制限"],
    solutions: [
      "APIキーの確認",
      "エンドポイントURL検証",
      "リクエスト頻度の調整"
    ],
    prevention: [
      "定期的なAPI設定チェック",
      "レート制限監視の実装",
      "フォールバック機構の準備"
    ]
  }
  // 他のパターンも追加
};
```

## 禁止事項と注意点
- ❌ 英語でのコメントや説明
- ❌ 機密情報のハードコーディング  
- ❌ セキュリティリスクのあるコード
- ❌ 著作権に問題のあるコード
- ❌ 既存の動作を破壊する変更
- ❌ 不必要な依存関係の追加
- ❌ テストなしでの本番デプロイ
- ❌ ドキュメント更新の忘れ

## 緊急時対応手順

### 1. 障害発生時の対応
```
1. **初期対応** (5分以内)
   - 障害状況の把握
   - 影響範囲の特定
   - 一時的な回避策の実施

2. **詳細調査** (30分以内)
   - ログの詳細分析
   - 原因の特定
   - 修正方針の決定

3. **修正・復旧** (1時間以内)
   - 修正コードの実装
   - テスト実行
   - デプロイと動作確認

4. **事後対応** (24時間以内)
   - 根本原因分析
   - 再発防止策の検討
   - ドキュメント更新
```

### 2. セキュリティインシデント対応
```
1. **即座に実行**
   - 影響のあるシステムの隔離
   - アクセスログの保全
   - 関係者への連絡

2. **調査・対応**
   - 侵入経路の特定
   - 被害範囲の調査
   - セキュリティパッチの適用

3. **再発防止**
   - セキュリティ設定の見直し
   - 監視体制の強化
   - 教育・訓練の実施
```
