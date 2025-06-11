' TestModernUI.bas
' Outlook AI Helper Modern UI 簡易テスト
' 作成日: 2024
' 説明: 新しいモダンUI機能の基本動作テスト

Option Explicit

' モダンUI機能の簡易テスト
Public Sub QuickTestModernUI()
    On Error GoTo ErrorHandler
    
    Dim testResults As String
    testResults = "Outlook AI Helper - モダンUI簡易テスト" & vbCrLf & vbCrLf
    
    ' 1. 日本語エイリアス関数の存在確認
    testResults = testResults & "📋 日本語エイリアス関数テスト:" & vbCrLf
    
    ' AIヘルパー_統合メニュー の存在確認
    On Error Resume Next
    Dim hasAIHelper As Boolean
    hasAIHelper = False
    
    ' 関数の存在をチェック
    Application.Run "AIヘルパー_統合メニュー"
    If Err.Number = 0 Then
        hasAIHelper = True
        testResults = testResults & "✅ AIヘルパー_統合メニュー: 利用可能" & vbCrLf
    ElseIf Err.Number = 1004 Then
        testResults = testResults & "❌ AIヘルパー_統合メニュー: 関数が見つかりません" & vbCrLf
    Else
        testResults = testResults & "⚠️ AIヘルパー_統合メニュー: " & Err.Description & vbCrLf
    End If
    
    On Error GoTo ErrorHandler
    
    ' 2. ShowEnhancedMainMenu の存在確認
    On Error Resume Next
    Application.Run "ShowEnhancedMainMenu"
    If Err.Number = 0 Then
        testResults = testResults & "✅ ShowEnhancedMainMenu: 利用可能" & vbCrLf
    ElseIf Err.Number = 1004 Then
        testResults = testResults & "❌ ShowEnhancedMainMenu: 関数が見つかりません" & vbCrLf
    Else
        testResults = testResults & "⚠️ ShowEnhancedMainMenu: " & Err.Description & vbCrLf
    End If
    
    On Error GoTo ErrorHandler
    
    ' 3. AnalyzeSearchFolders の存在確認
    On Error Resume Next
    Application.Run "AnalyzeSearchFolders"
    If Err.Number = 0 Then
        testResults = testResults & "✅ AnalyzeSearchFolders: 利用可能" & vbCrLf
    ElseIf Err.Number = 1004 Then
        testResults = testResults & "❌ AnalyzeSearchFolders: 関数が見つかりません" & vbCrLf
    Else
        testResults = testResults & "⚠️ AnalyzeSearchFolders: " & Err.Description & vbCrLf
    End If
    
    On Error GoTo ErrorHandler
    
    ' 4. 基本的な設定確認
    testResults = testResults & vbCrLf & "🔧 基本設定確認:" & vbCrLf
    
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Then
        testResults = testResults & "⚠️ API Key: 未設定（設定してください）" & vbCrLf
    Else
        testResults = testResults & "✅ API Key: 設定済み" & vbCrLf
    End If
    
    If InStr(OPENAI_API_ENDPOINT, "your-azure-openai-endpoint") > 0 Then
        testResults = testResults & "⚠️ API Endpoint: 未設定（設定してください）" & vbCrLf
    Else
        testResults = testResults & "✅ API Endpoint: 設定済み" & vbCrLf
    End If
    
    ' 5. 推奨する次のステップ
    testResults = testResults & vbCrLf & "🚀 推奨次ステップ:" & vbCrLf
    If hasAIHelper Then
        testResults = testResults & "1. AIヘルパー_統合メニュー を実行してUIをテスト" & vbCrLf
        testResults = testResults & "2. API設定が完了していれば、メール解析をテスト" & vbCrLf
    Else
        testResults = testResults & "1. OutlookAI_Unified.bas を最新版に更新" & vbCrLf
        testResults = testResults & "2. OutlookAI_MainForm.bas をインポート" & vbCrLf
        testResults = testResults & "3. VBAプロジェクトを保存" & vbCrLf
    End If
    
    testResults = testResults & vbCrLf & "テスト完了: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    MsgBox testResults, vbInformation, "モダンUI簡易テスト結果"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "テスト実行中にエラーが発生しました: " & Err.Description, vbCritical, "テストエラー"
End Sub

' インストール状況の確認
Public Sub CheckInstallationStatus()
    On Error GoTo ErrorHandler
    
    Dim statusReport As String
    statusReport = "Outlook AI Helper - インストール状況確認" & vbCrLf & vbCrLf
    
    ' VBAプロジェクト内のモジュール確認
    statusReport = statusReport & "📂 VBAモジュール確認:" & vbCrLf
    
    ' OutlookAI_Unified.bas の確認
    On Error Resume Next
    Application.Run "ShowMainMenu"
    If Err.Number = 0 Then
        statusReport = statusReport & "✅ OutlookAI_Unified.bas: インストール済み" & vbCrLf
    Else
        statusReport = statusReport & "❌ OutlookAI_Unified.bas: 未インストール" & vbCrLf
    End If
    
    ' OutlookAI_MainForm.bas の確認
    Application.Run "ShowMainForm"
    If Err.Number = 0 Then
        statusReport = statusReport & "✅ OutlookAI_MainForm.bas: インストール済み" & vbCrLf
    Else
        statusReport = statusReport & "❌ OutlookAI_MainForm.bas: 未インストール" & vbCrLf
    End If
    
    On Error GoTo ErrorHandler
    
    ' インストール手順の案内
    statusReport = statusReport & vbCrLf & "📋 インストール手順:" & vbCrLf
    statusReport = statusReport & "1. VBAエディタ（Alt+F11）を開く" & vbCrLf
    statusReport = statusReport & "2. ファイル > インポート で.basファイルを選択" & vbCrLf
    statusReport = statusReport & "3. 必要なファイル:" & vbCrLf
    statusReport = statusReport & "   - OutlookAI_Unified.bas（必須）" & vbCrLf
    statusReport = statusReport & "   - OutlookAI_MainForm.bas（モダンUI用）" & vbCrLf
    statusReport = statusReport & "4. API設定を完了" & vbCrLf
    statusReport = statusReport & "5. AIヘルパー_統合メニュー を実行"
    
    MsgBox statusReport, vbInformation, "インストール状況確認"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "インストール状況確認中にエラーが発生しました: " & Err.Description, vbCritical, "確認エラー"
End Sub