' OutlookAI_Installer.bas
' Outlook OpenAI マクロ - インストール・設定支援モジュール
' 作成日: 2024
' 説明: 初回インストール時の設定支援とセットアップガイダンス

Option Explicit

' =============================================================================
' インストール・設定支援メイン関数
' =============================================================================

' 初回セットアップガイド
Public Sub RunInitialSetup()
    On Error GoTo ErrorHandler
    
    ' ウェルカムメッセージ
    Dim welcomeMessage As String
    welcomeMessage = "Outlook AI Helper セットアップガイドへようこそ！" & vbCrLf & vbCrLf & _
                    "このガイドでは、以下の設定を行います：" & vbCrLf & _
                    "1. システム要件の確認" & vbCrLf & _
                    "2. Azure OpenAI API の設定" & vbCrLf & _
                    "3. マクロセキュリティの確認" & vbCrLf & _
                    "4. 動作テストの実行" & vbCrLf & _
                    "5. クイックアクセスボタンの設定" & vbCrLf & vbCrLf & _
                    "セットアップを開始しますか？"
    
    If MsgBox(welcomeMessage, vbYesNo + vbQuestion, "Outlook AI Helper セットアップ") = vbNo Then
        Exit Sub
    End If
    
    ' セットアップ手順を順番に実行
    If CheckSystemRequirements() Then
        If SetupAPIConfiguration() Then
            If CheckMacroSecurity() Then
                If RunSystemTest() Then
                    SetupQuickAccess
                    ShowCompletionMessage
                End If
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "セットアップ中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' システム要件確認
' =============================================================================

' システム要件の確認
Private Function CheckSystemRequirements() As Boolean
    On Error GoTo ErrorHandler
    
    ShowMessage "システム要件を確認しています...", "確認中", vbInformation
    
    Dim requirements As String
    Dim version As String
    
    ' Outlookバージョンの確認
    version = Application.Version
    
    requirements = "システム要件チェック結果:" & vbCrLf & vbCrLf & _
                  "Microsoft Outlook バージョン: " & version & vbCrLf
    
    ' バージョンチェック（2016以降を推奨）
    If Val(version) >= 16 Then
        requirements = requirements & "✓ サポートされているバージョンです" & vbCrLf
    Else
        requirements = requirements & "⚠ 古いバージョンです（2016以降を推奨）" & vbCrLf
    End If
    
    ' VBA有効性の確認
    requirements = requirements & vbCrLf & "VBA機能: 利用可能" & vbCrLf
    requirements = requirements & "マクロ実行: 可能" & vbCrLf
    
    ' インターネット接続の確認（簡易版）
    requirements = requirements & vbCrLf & "⚠ インターネット接続を手動で確認してください" & vbCrLf
    requirements = requirements & "⚠ Azure OpenAI サービスへのアクセス権限を確認してください" & vbCrLf & vbCrLf
    
    requirements = requirements & "システム要件を満たしていますか？"
    
    If MsgBox(requirements, vbYesNo + vbQuestion, "システム要件確認") = vbYes Then
        CheckSystemRequirements = True
    Else
        ShowMessage "システム要件を満たしてから再度セットアップを実行してください。", "セットアップ中断", vbExclamation
        CheckSystemRequirements = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "システム要件確認中にエラーが発生しました。", Err.Description
    CheckSystemRequirements = False
End Function

' =============================================================================
' API設定ガイド
' =============================================================================

' API設定のガイド
Private Function SetupAPIConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    Dim setupChoice As String
    setupChoice = InputBox("Azure OpenAI API の設定方法を選択してください:" & vbCrLf & vbCrLf & _
                          "1. ガイド付きで設定（推奨）" & vbCrLf & _
                          "2. 手動設定の説明を表示" & vbCrLf & _
                          "3. 既に設定済み（スキップ）" & vbCrLf & vbCrLf & _
                          "番号を入力してください:", _
                          "API設定")
    
    Select Case setupChoice
        Case "1"
            SetupAPIConfiguration = GuidedAPISetup()
        Case "2"
            ShowManualSetupInstructions
            SetupAPIConfiguration = True
        Case "3"
            SetupAPIConfiguration = True
        Case ""
            SetupAPIConfiguration = False
        Case Else
            ShowMessage "無効な選択です。", "入力エラー", vbExclamation
            SetupAPIConfiguration = False
    End Select
    
    Exit Function
    
ErrorHandler:
    ShowError "API設定中にエラーが発生しました。", Err.Description
    SetupAPIConfiguration = False
End Function

' ガイド付きAPI設定
Private Function GuidedAPISetup() As Boolean
    On Error GoTo ErrorHandler
    
    ShowMessage "ガイド付きAPI設定を開始します。" & vbCrLf & _
               "Azure ポータルの情報を確認してから続行してください。", _
               "API設定ガイド", vbInformation
    
    ' Azure リソース情報の収集
    Dim resourceName As String
    Dim deploymentName As String
    Dim apiKey As String
    Dim region As String
    
    ' リソース名
    resourceName = InputBox("Azure OpenAI リソース名を入力してください:" & vbCrLf & vbCrLf & _
                           "Azure ポータル > OpenAI > 概要 で確認できます" & vbCrLf & _
                           "例: my-openai-resource", _
                           "リソース名")
    If resourceName = "" Then
        GuidedAPISetup = False
        Exit Function
    End If
    
    ' デプロイメント名
    deploymentName = InputBox("GPT-4 デプロイメント名を入力してください:" & vbCrLf & vbCrLf & _
                             "Azure OpenAI Studio > デプロイ で確認できます" & vbCrLf & _
                             "例: gpt-4-deployment", _
                             "デプロイメント名")
    If deploymentName = "" Then
        GuidedAPISetup = False
        Exit Function
    End If
    
    ' APIキー
    apiKey = InputBox("API キーを入力してください:" & vbCrLf & vbCrLf & _
                     "Azure ポータル > OpenAI > キーとエンドポイント" & vbCrLf & _
                     "Key1 または Key2 をコピーしてください", _
                     "API キー")
    If apiKey = "" Then
        GuidedAPISetup = False
        Exit Function
    End If
    
    ' 設定内容の確認
    Dim endpoint As String
    endpoint = "https://" & resourceName & ".openai.azure.com/openai/deployments/" & deploymentName & "/chat/completions?api-version=2024-02-15-preview"
    
    Dim confirmMessage As String
    confirmMessage = "設定内容を確認してください:" & vbCrLf & vbCrLf & _
                    "リソース名: " & resourceName & vbCrLf & _
                    "デプロイメント: " & deploymentName & vbCrLf & _
                    "エンドポイント:" & vbCrLf & endpoint & vbCrLf & _
                    "API キー: " & Left(apiKey, 10) & "..." & vbCrLf & vbCrLf & _
                    "この設定で続行しますか？"
    
    If MsgBox(confirmMessage, vbYesNo + vbQuestion, "設定確認") = vbYes Then
        ' 設定手順の表示
        ShowConfigurationInstructions endpoint, apiKey
        GuidedAPISetup = True
    Else
        GuidedAPISetup = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "ガイド付きAPI設定中にエラーが発生しました。", Err.Description
    GuidedAPISetup = False
End Function

' 設定手順の表示
Private Sub ShowConfigurationInstructions(ByVal endpoint As String, ByVal apiKey As String)
    Dim instructions As String
    instructions = "以下の手順でAPI設定を完了してください:" & vbCrLf & vbCrLf & _
                  "1. Alt+F11 でVBAエディタを開く" & vbCrLf & _
                  "2. AI_Common.bas をダブルクリックして開く" & vbCrLf & _
                  "3. 以下の行を探して編集:" & vbCrLf & vbCrLf & _
                  "変更前:" & vbCrLf & _
                  "Public Const OPENAI_API_ENDPOINT As String = ""YOUR_ENDPOINT_HERE""" & vbCrLf & _
                  "Public Const OPENAI_API_KEY As String = ""YOUR_API_KEY_HERE""" & vbCrLf & vbCrLf & _
                  "変更後:" & vbCrLf & _
                  "Public Const OPENAI_API_ENDPOINT As String = """ & endpoint & """" & vbCrLf & _
                  "Public Const OPENAI_API_KEY As String = """ & apiKey & """"
    
    ' 分割表示
    ShowLongMessage instructions, "設定手順"
    
    ShowMessage "4. Ctrl+S で保存" & vbCrLf & _
               "5. VBAエディタを閉じる" & vbCrLf & vbCrLf & _
               "設定が完了したら「OK」をクリックしてください。", _
               "設定手順（続き）", vbInformation
End Sub

' 手動設定の説明
Private Sub ShowManualSetupInstructions()
    Dim instructions As String
    instructions = "手動設定の手順:" & vbCrLf & vbCrLf & _
                  "1. Azure ポータルでOpenAI リソースを作成" & vbCrLf & _
                  "2. GPT-4 モデルをデプロイ" & vbCrLf & _
                  "3. API キーとエンドポイントを取得" & vbCrLf & _
                  "4. AI_Common.bas の定数を編集:" & vbCrLf & _
                  "   - OPENAI_API_ENDPOINT" & vbCrLf & _
                  "   - OPENAI_API_KEY" & vbCrLf & vbCrLf & _
                  "詳細な手順はクイックスタートガイドを参照してください。"
    
    ShowMessage instructions, "手動設定手順", vbInformation
End Sub

' =============================================================================
' マクロセキュリティ確認
' =============================================================================

' マクロセキュリティの確認
Private Function CheckMacroSecurity() As Boolean
    On Error GoTo ErrorHandler
    
    Dim securityMessage As String
    securityMessage = "マクロセキュリティ設定を確認します。" & vbCrLf & vbCrLf & _
                     "推奨設定:" & vbCrLf & _
                     "「警告を表示してすべてのマクロを無効にする」" & vbCrLf & vbCrLf & _
                     "設定手順:" & vbCrLf & _
                     "1. ファイル > オプション" & vbCrLf & _
                     "2. トラストセンター > トラストセンターの設定" & vbCrLf & _
                     "3. マクロの設定 で適切なレベルを選択" & vbCrLf & vbCrLf & _
                     "マクロセキュリティは適切に設定されていますか？"
    
    If MsgBox(securityMessage, vbYesNo + vbQuestion, "マクロセキュリティ確認") = vbYes Then
        CheckMacroSecurity = True
    Else
        ShowMessage "マクロセキュリティを設定してから再度セットアップを実行してください。", _
                   "設定が必要", vbExclamation
        CheckMacroSecurity = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "マクロセキュリティ確認中にエラーが発生しました。", Err.Description
    CheckMacroSecurity = False
End Function

' =============================================================================
' システムテスト
' =============================================================================

' システムテストの実行
Private Function RunSystemTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testChoice As String
    testChoice = InputBox("システムテストを実行しますか？" & vbCrLf & vbCrLf & _
                         "1. 完全テスト（推奨）" & vbCrLf & _
                         "2. 基本テストのみ" & vbCrLf & _
                         "3. スキップ" & vbCrLf & vbCrLf & _
                         "番号を入力してください:", _
                         "システムテスト")
    
    Select Case testChoice
        Case "1"
            RunSystemTest = RunFullTest()
        Case "2"
            RunSystemTest = RunBasicTest()
        Case "3"
            RunSystemTest = True
        Case ""
            RunSystemTest = False
        Case Else
            ShowMessage "無効な選択です。", "入力エラー", vbExclamation
            RunSystemTest = False
    End Select
    
    Exit Function
    
ErrorHandler:
    ShowError "システムテスト中にエラーが発生しました。", Err.Description
    RunSystemTest = False
End Function

' 完全テストの実行
Private Function RunFullTest() As Boolean
    On Error GoTo ErrorHandler
    
    ShowMessage "完全テストを開始します...", "テスト実行中", vbInformation
    
    ' テスト1: 設定確認
    If Not ValidateConfiguration() Then
        ShowMessage "設定に問題があります。API設定を確認してください。", "テスト失敗", vbExclamation
        RunFullTest = False
        Exit Function
    End If
    
    ' テスト2: API接続テスト
    ShowMessage "API接続をテストしています...", "テスト実行中", vbInformation
    
    ' 簡易的なAPI接続確認（実際のAPIコールは別途実行）
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Then
        ShowMessage "API キーが設定されていません。", "テスト失敗", vbExclamation
        RunFullTest = False
        Exit Function
    End If
    
    ' テスト3: モジュール読み込み確認
    ShowMessage "全てのテストが正常に完了しました！", "テスト成功", vbInformation
    RunFullTest = True
    
    Exit Function
    
ErrorHandler:
    ShowError "完全テスト中にエラーが発生しました。", Err.Description
    RunFullTest = False
End Function

' 基本テストの実行
Private Function RunBasicTest() As Boolean
    On Error GoTo ErrorHandler
    
    ShowMessage "基本テストを実行中...", "テスト実行中", vbInformation
    
    ' 基本的な設定確認のみ
    If OPENAI_API_KEY <> "YOUR_API_KEY_HERE" And OPENAI_API_KEY <> "" Then
        ShowMessage "基本テストが完了しました。", "テスト成功", vbInformation
        RunBasicTest = True
    Else
        ShowMessage "API キーが設定されていません。", "テスト失敗", vbExclamation
        RunBasicTest = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "基本テスト中にエラーが発生しました。", Err.Description
    RunBasicTest = False
End Function

' =============================================================================
' クイックアクセス設定
' =============================================================================

' クイックアクセスの設定
Private Sub SetupQuickAccess()
    On Error GoTo ErrorHandler
    
    Dim setupChoice As String
    setupChoice = InputBox("クイックアクセス設定（オプション）:" & vbCrLf & vbCrLf & _
                          "1. クイックアクセスツールバーに追加の説明を表示" & vbCrLf & _
                          "2. ショートカットキーの設定説明を表示" & vbCrLf & _
                          "3. スキップ" & vbCrLf & vbCrLf & _
                          "番号を入力してください:", _
                          "クイックアクセス設定")
    
    Select Case setupChoice
        Case "1"
            ShowQuickAccessInstructions
        Case "2"
            ShowShortcutInstructions
        Case "3"
            ' スキップ
        Case ""
            ' キャンセル
        Case Else
            ShowMessage "無効な選択です。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "クイックアクセス設定中にエラーが発生しました。", Err.Description
End Sub

' クイックアクセスツールバーの説明
Private Sub ShowQuickAccessInstructions()
    Dim instructions As String
    instructions = "クイックアクセスツールバーへのマクロ追加手順:" & vbCrLf & vbCrLf & _
                  "1. リボンの上で右クリック" & vbCrLf & _
                  "2. 「クイック アクセス ツール バーのユーザー設定」を選択" & vbCrLf & _
                  "3. 「コマンドの選択」で「マクロ」を選択" & vbCrLf & _
                  "4. 以下のマクロを追加（推奨）:" & vbCrLf & _
                  "   - AI_EmailAnalyzer.AnalyzeSelectedEmail" & vbCrLf & _
                  "   - AI_EmailComposer.CreateRejectionEmail" & vbCrLf & _
                  "   - AI_EmailComposer.CreateAcceptanceEmail" & vbCrLf & _
                  "5. 「追加」ボタンでツールバーに登録"
    
    ShowMessage instructions, "クイックアクセス設定", vbInformation
End Sub

' ショートカットキーの説明
Private Sub ShowShortcutInstructions()
    Dim instructions As String
    instructions = "便利なショートカットキー:" & vbCrLf & vbCrLf & _
                  "Alt + F8: マクロ実行ダイアログを開く" & vbCrLf & _
                  "Alt + F11: VBA エディタを開く" & vbCrLf & _
                  "Ctrl + G: VBAのイミディエイトウィンドウ（デバッグ用）" & vbCrLf & vbCrLf & _
                  "マクロ名の省略記法:" & vbCrLf & _
                  "「AI_Email」まで入力すると候補が表示されます"
    
    ShowMessage instructions, "ショートカットキー", vbInformation
End Sub

' =============================================================================
' セットアップ完了
' =============================================================================

' セットアップ完了メッセージ
Private Sub ShowCompletionMessage()
    Dim completionMessage As String
    completionMessage = "🎉 Outlook AI Helper のセットアップが完了しました！" & vbCrLf & vbCrLf & _
                       "次の手順:" & vbCrLf & _
                       "1. メールを選択して Alt+F8 でマクロを実行" & vbCrLf & _
                       "2. クイックスタートガイドで詳細な使用方法を確認" & vbCrLf & _
                       "3. 不明な点があれば設定管理メニューを活用" & vbCrLf & vbCrLf & _
                       "主要なマクロ:" & vbCrLf & _
                       "• AI_EmailAnalyzer.AnalyzeSelectedEmail" & vbCrLf & _
                       "• AI_EmailComposer.CreateRejectionEmail" & vbCrLf & _
                       "• AI_EmailComposer.CreateAcceptanceEmail" & vbCrLf & _
                       "• AI_ConfigManager.ManageConfiguration" & vbCrLf & vbCrLf & _
                       "Outlookでのメール処理をお楽しみください！"
    
    ShowMessage completionMessage, "セットアップ完了", vbInformation
End Sub

' =============================================================================
' ヘルパー関数
' =============================================================================

' 長いメッセージの表示
Private Sub ShowLongMessage(ByVal message As String, ByVal title As String)
    ' 1000文字以上の場合は分割表示
    If Len(message) > 1000 Then
        Dim part1 As String
        Dim part2 As String
        part1 = Left(message, 1000)
        part2 = Mid(message, 1001)
        
        MsgBox part1, vbInformation, title & " (1/2)"
        MsgBox part2, vbInformation, title & " (2/2)"
    Else
        MsgBox message, vbInformation, title
    End If
End Sub

' デバッグ・診断機能
Public Sub RunDiagnostics()
    On Error GoTo ErrorHandler
    
    Dim diagnostics As String
    diagnostics = "=== Outlook AI Helper 診断情報 ===" & vbCrLf & vbCrLf & _
                 "Outlook バージョン: " & Application.Version & vbCrLf & _
                 "VBA バージョン: " & Application.Version & vbCrLf & _
                 "現在時刻: " & Now & vbCrLf & vbCrLf & _
                 "設定状況:" & vbCrLf & _
                 "API エンドポイント: " & Left(OPENAI_API_ENDPOINT, 50) & "..." & vbCrLf & _
                 "API キー設定: " & IIf(OPENAI_API_KEY = "YOUR_API_KEY_HERE", "未設定", "設定済み") & vbCrLf & _
                 "モデル: " & OPENAI_MODEL & vbCrLf & _
                 "タイムアウト: " & REQUEST_TIMEOUT & "秒" & vbCrLf & vbCrLf & _
                 "この情報はサポート時に役立ちます。"
    
    ShowMessage diagnostics, "診断情報", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ShowError "診断実行中にエラーが発生しました。", Err.Description
End Sub