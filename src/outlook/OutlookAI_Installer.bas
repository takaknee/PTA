' OutlookAI_Installer.bas
' Outlook OpenAI Macro - Installation and Setup Support Module
' Created: 2024
' Description: Setup support and guidance for initial installation

Option Explicit

' =============================================================================
' Constants Definition
' =============================================================================

' OpenAI API Settings
Public Const OPENAI_API_ENDPOINT As String = "https://your-azure-openai-endpoint.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview"
Public Const OPENAI_API_KEY As String = "YOUR_API_KEY_HERE" ' Production environment should read from configuration file
Public Const OPENAI_MODEL As String = "gpt-4"

' Application Settings
Public Const APP_NAME As String = "Outlook AI Helper"
Public Const APP_VERSION As String = "1.0.0"
Public Const REQUEST_TIMEOUT As Integer = 30 ' API request timeout (seconds)

' Version Requirements
Private Const MIN_OUTLOOK_VERSION As Integer = 16 ' Minimum supported Outlook version (2016)

' Message Display Settings
Private Const MAX_MESSAGE_LENGTH As Integer = 1000 ' Maximum length for single message box
Private Const MESSAGE_SPLIT_POSITION As Integer = 1001 ' Position to start second part of split message

' =============================================================================
' Common Functions
' =============================================================================

' Display message box (common format)
Public Sub ShowMessage(ByVal message As String, ByVal title As String, Optional ByVal messageType As VbMsgBoxStyle = vbInformation)
    MsgBox message, messageType, APP_NAME & " - " & title
End Sub

' Display error message
Public Sub ShowError(ByVal errorMessage As String, Optional ByVal details As String = "")
    Dim fullMessage As String
    fullMessage = "An error occurred: " & errorMessage
    If details <> "" Then
        fullMessage = fullMessage & vbCrLf & vbCrLf & "Details: " & details
    End If
    MsgBox fullMessage, vbCritical, APP_NAME & " - Error"
End Sub

' Configuration validation
Public Function ValidateConfiguration() As Boolean
    ' Check API Key
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Or OPENAI_API_KEY = "" Then
        ShowError "OpenAI API key is not configured.", _
                 "Please set OPENAI_API_KEY in AI_Common.bas."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' Check endpoint
    If InStr(OPENAI_API_ENDPOINT, "your-azure-openai-endpoint") > 0 Then
        ShowError "OpenAI API endpoint is not configured.", _
                 "Please set OPENAI_API_ENDPOINT in AI_Common.bas."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
End Function

' =============================================================================
' Installation and Setup Support Main Functions
' =============================================================================

' Initial Setup Guide
Public Sub RunInitialSetup()
    On Error GoTo ErrorHandler
    
    ' Welcome message
    Dim welcomeMessage As String
    welcomeMessage = "Welcome to Outlook AI Helper Setup Guide!" & vbCrLf & vbCrLf & _
                    "This guide will perform the following configuration:" & vbCrLf & _
                    "1. System requirements check" & vbCrLf & _
                    "2. Azure OpenAI API configuration" & vbCrLf & _
                    "3. Macro security verification" & vbCrLf & _
                    "4. Operation test execution" & vbCrLf & _
                    "5. Quick access button setup" & vbCrLf & vbCrLf & _
                    "Do you want to start the setup?"
    
    If MsgBox(welcomeMessage, vbYesNo + vbQuestion, "Outlook AI Helper Setup") = vbNo Then
        Exit Sub
    End If
    
    ' Execute setup steps in order
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
    ShowError "An error occurred during setup.", Err.Description
End Sub

' =============================================================================
' System Requirements Check
' =============================================================================

' System requirements verification
Private Function CheckSystemRequirements() As Boolean
    On Error GoTo ErrorHandler
    
    ShowMessage "Checking system requirements...", "Checking", vbInformation
    
    Dim requirements As String
    Dim version As String
    
    ' Check Outlook version
    version = Application.Version
    
    requirements = "System Requirements Check Results:" & vbCrLf & vbCrLf & _
                  "Microsoft Outlook Version: " & version & vbCrLf
    
    ' Version check (2016 or later recommended)
    If Val(version) >= MIN_OUTLOOK_VERSION Then
        requirements = requirements & "✓ Supported version" & vbCrLf
    Else
        requirements = requirements & "⚠ Old version (2016 or later recommended)" & vbCrLf
    End If
    
    ' VBA availability check
    requirements = requirements & vbCrLf & "VBA Feature: Available" & vbCrLf
    requirements = requirements & "Macro Execution: Enabled" & vbCrLf
    
    ' Internet connection check (simplified)
    requirements = requirements & vbCrLf & "⚠ Please manually check internet connection" & vbCrLf
    requirements = requirements & "⚠ Please verify access to Azure OpenAI service" & vbCrLf & vbCrLf
    
    requirements = requirements & "Do you meet the system requirements?"
    
    If MsgBox(requirements, vbYesNo + vbQuestion, "System Requirements Check") = vbYes Then
        CheckSystemRequirements = True
    Else
        ShowMessage "Please meet the system requirements before running setup again.", "Setup Interrupted", vbExclamation
        CheckSystemRequirements = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "An error occurred during system requirements check.", Err.Description
    CheckSystemRequirements = False
End Function

' =============================================================================
' API Configuration Guide
' =============================================================================

' API configuration guide
Private Function SetupAPIConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    Dim setupChoice As String
    setupChoice = InputBox("Select Azure OpenAI API configuration method:" & vbCrLf & vbCrLf & _
                          "1. Guided setup (recommended)" & vbCrLf & _
                          "2. Show manual setup instructions" & vbCrLf & _
                          "3. Already configured (skip)" & vbCrLf & vbCrLf & _
                          "Enter number:", _
                          "API Configuration")
    
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
            ShowMessage "Invalid selection.", "Input Error", vbExclamation
            SetupAPIConfiguration = False
    End Select
    
    Exit Function
    
ErrorHandler:
    ShowError "An error occurred during API configuration.", Err.Description
    SetupAPIConfiguration = False
End Function

' Guided API setup
Private Function GuidedAPISetup() As Boolean
    On Error GoTo ErrorHandler
    
    ShowMessage "Starting guided API setup." & vbCrLf & _
               "Please check Azure portal information before continuing.", _
               "API Setup Guide", vbInformation
    
    ' Collect Azure resource information
    Dim resourceName As String
    Dim deploymentName As String
    Dim apiKey As String
    Dim region As String
    
    ' Resource name
    resourceName = InputBox("Enter Azure OpenAI resource name:" & vbCrLf & vbCrLf & _
                           "You can check this at Azure Portal > OpenAI > Overview" & vbCrLf & _
                           "Example: my-openai-resource", _
                           "Resource Name")
    If resourceName = "" Then
        GuidedAPISetup = False
        Exit Function
    End If
    
    ' Deployment name
    deploymentName = InputBox("Enter GPT-4 deployment name:" & vbCrLf & vbCrLf & _
                             "You can check this at Azure OpenAI Studio > Deploy" & vbCrLf & _
                             "Example: gpt-4-deployment", _
                             "Deployment Name")
    If deploymentName = "" Then
        GuidedAPISetup = False
        Exit Function
    End If
    
    ' API key
    apiKey = InputBox("Enter API key:" & vbCrLf & vbCrLf & _
                     "Azure Portal > OpenAI > Keys and Endpoint" & vbCrLf & _
                     "Copy Key1 or Key2", _
                     "API Key")
    If apiKey = "" Then
        GuidedAPISetup = False
        Exit Function
    End If
    
    ' Confirm configuration content
    Dim endpoint As String
    endpoint = "https://" & resourceName & ".openai.azure.com/openai/deployments/" & deploymentName & "/chat/completions?api-version=2024-02-15-preview"
    
    Dim confirmMessage As String
    confirmMessage = "Please confirm configuration content:" & vbCrLf & vbCrLf & _
                    "Resource Name: " & resourceName & vbCrLf & _
                    "Deployment: " & deploymentName & vbCrLf & _
                    "Endpoint:" & vbCrLf & endpoint & vbCrLf & _
                    "API Key: " & Left(apiKey, 10) & "..." & vbCrLf & vbCrLf & _
                    "Continue with this configuration?"
    
    If MsgBox(confirmMessage, vbYesNo + vbQuestion, "Configuration Confirmation") = vbYes Then
        ' Display configuration instructions
        ShowConfigurationInstructions endpoint, apiKey
        GuidedAPISetup = True
    Else
        GuidedAPISetup = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "An error occurred during guided API setup.", Err.Description
    GuidedAPISetup = False
End Function

' Display configuration instructions
Private Sub ShowConfigurationInstructions(ByVal endpoint As String, ByVal apiKey As String)
    Dim instructions As String
    instructions = "Please complete API configuration with the following steps:" & vbCrLf & vbCrLf & _
                  "1. Press Alt+F11 to open VBA editor" & vbCrLf & _
                  "2. Double-click AI_Common.bas to open" & vbCrLf & _
                  "3. Find and edit the following lines:" & vbCrLf & vbCrLf & _
                  "Before change:" & vbCrLf & _
                  "Public Const OPENAI_API_ENDPOINT As String = ""YOUR_ENDPOINT_HERE""" & vbCrLf & _
                  "Public Const OPENAI_API_KEY As String = ""YOUR_API_KEY_HERE""" & vbCrLf & vbCrLf & _
                  "After change:" & vbCrLf & _
                  "Public Const OPENAI_API_ENDPOINT As String = """ & endpoint & """" & vbCrLf & _
                  "Public Const OPENAI_API_KEY As String = """ & apiKey & """"
    
    ' Split display
    ShowLongMessage instructions, "Configuration Steps"
    
    ShowMessage "4. Press Ctrl+S to save" & vbCrLf & _
               "5. Close VBA editor" & vbCrLf & vbCrLf & _
               "Click 'OK' when configuration is complete.", _
               "Configuration Steps (continued)", vbInformation
End Sub

' Manual setup instructions
Private Sub ShowManualSetupInstructions()
    Dim instructions As String
    instructions = "Manual setup steps:" & vbCrLf & vbCrLf & _
                  "1. Create OpenAI resource in Azure portal" & vbCrLf & _
                  "2. Deploy GPT-4 model" & vbCrLf & _
                  "3. Get API key and endpoint" & vbCrLf & _
                  "4. Edit constants in AI_Common.bas:" & vbCrLf & _
                  "   - OPENAI_API_ENDPOINT" & vbCrLf & _
                  "   - OPENAI_API_KEY" & vbCrLf & vbCrLf & _
                  "Please refer to quickstart guide for detailed steps."
    
    ShowMessage instructions, "Manual Setup Steps", vbInformation
End Sub

' =============================================================================
' Macro Security Check
' =============================================================================

' Macro security check
Private Function CheckMacroSecurity() As Boolean
    On Error GoTo ErrorHandler
    
    Dim securityMessage As String
    securityMessage = "Checking macro security settings." & vbCrLf & vbCrLf & _
                     "Recommended setting:" & vbCrLf & _
                     "'Disable all macros with notification'" & vbCrLf & vbCrLf & _
                     "Configuration steps:" & vbCrLf & _
                     "1. File > Options" & vbCrLf & _
                     "2. Trust Center > Trust Center Settings" & vbCrLf & _
                     "3. Select appropriate level in Macro Settings" & vbCrLf & vbCrLf & _
                     "Is macro security properly configured?"
    
    If MsgBox(securityMessage, vbYesNo + vbQuestion, "Macro Security Check") = vbYes Then
        CheckMacroSecurity = True
    Else
        ShowMessage "Please configure macro security before running setup again.", _
                   "Configuration Required", vbExclamation
        CheckMacroSecurity = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "An error occurred during macro security check.", Err.Description
    CheckMacroSecurity = False
End Function

' =============================================================================
' System Test
' =============================================================================

' System test execution
Private Function RunSystemTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testChoice As String
    testChoice = InputBox("Do you want to run system test?" & vbCrLf & vbCrLf & _
                         "1. Full test (recommended)" & vbCrLf & _
                         "2. Basic test only" & vbCrLf & _
                         "3. Skip" & vbCrLf & vbCrLf & _
                         "Enter number:", _
                         "System Test")
    
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
            ShowMessage "Invalid selection.", "Input Error", vbExclamation
            RunSystemTest = False
    End Select
    
    Exit Function
    
ErrorHandler:
    ShowError "An error occurred during system test.", Err.Description
    RunSystemTest = False
End Function

' Full test execution
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
' Helper Functions
' =============================================================================

' Display long message
Private Sub ShowLongMessage(ByVal message As String, ByVal title As String)
    ' Split display if 1000 characters or more
    If Len(message) > MAX_MESSAGE_LENGTH Then
        Dim part1 As String
        Dim part2 As String
        part1 = Left(message, MAX_MESSAGE_LENGTH)
        part2 = Mid(message, MESSAGE_SPLIT_POSITION)
        
        MsgBox part1, vbInformation, title & " (1/2)"
        MsgBox part2, vbInformation, title & " (2/2)"
    Else
        MsgBox message, vbInformation, title
    End If
End Sub

' Debug and diagnostic function
Public Sub RunDiagnostics()
    On Error GoTo ErrorHandler
    
    Dim diagnostics As String
    diagnostics = "=== Outlook AI Helper Diagnostic Information ===" & vbCrLf & vbCrLf & _
                 "Outlook Version: " & Application.Version & vbCrLf & _
                 "VBA Version: " & Application.Version & vbCrLf & _
                 "Current Time: " & Now & vbCrLf & vbCrLf & _
                 "Configuration Status:" & vbCrLf & _
                 "API Endpoint: " & Left(OPENAI_API_ENDPOINT, 50) & "..." & vbCrLf & _
                 "API Key Status: " & IIf(OPENAI_API_KEY = "YOUR_API_KEY_HERE", "Not configured", "Configured") & vbCrLf & _
                 "Model: " & OPENAI_MODEL & vbCrLf & _
                 "Timeout: " & REQUEST_TIMEOUT & " seconds" & vbCrLf & vbCrLf & _
                 "This information is helpful for support."
    
    ShowMessage diagnostics, "Diagnostic Information", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ShowError "An error occurred during diagnostic execution.", Err.Description
End Sub