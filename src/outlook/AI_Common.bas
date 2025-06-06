' AI_Common.bas
' Outlook OpenAI マクロ - 共通関数・定数モジュール
' 作成日: 2024
' 説明: 全モジュールで使用する共通関数、定数、設定を管理

Option Explicit

' =============================================================================
' 定数定義
' =============================================================================

' OpenAI API 設定
Public Const OPENAI_API_ENDPOINT As String = "https://your-azure-openai-endpoint.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview"
Public Const OPENAI_API_KEY As String = "YOUR_API_KEY_HERE" ' 本番環境では設定ファイルから読み込み
Public Const OPENAI_MODEL As String = "gpt-4"

' アプリケーション設定
Public Const APP_NAME As String = "Outlook AI Helper"
Public Const APP_VERSION As String = "1.0.0"
Public Const MAX_CONTENT_LENGTH As Integer = 50000 ' 最大処理文字数
Public Const REQUEST_TIMEOUT As Integer = 30 ' APIリクエストタイムアウト（秒）

' プロンプトテンプレート
Public Const SYSTEM_PROMPT_ANALYZER As String = "あなたは日本語のビジネスメール分析の専門家です。メール内容を分析し、重要な情報を抽出してください。"
Public Const SYSTEM_PROMPT_COMPOSER As String = "あなたは日本語のビジネスメール作成の専門家です。適切で丁寧なビジネスメールを作成してください。"
Public Const SYSTEM_PROMPT_SEARCH As String = "あなたは日本語のメール整理・分析の専門家です。メールフォルダの分析と整理の提案をしてください。"

' =============================================================================
' 共通関数
' =============================================================================

' メッセージボックス表示（共通形式）
Public Sub ShowMessage(ByVal message As String, ByVal title As String, Optional ByVal messageType As VbMsgBoxStyle = vbInformation)
    MsgBox message, messageType, APP_NAME & " - " & title
End Sub

' エラーメッセージ表示
Public Sub ShowError(ByVal errorMessage As String, Optional ByVal details As String = "")
    Dim fullMessage As String
    fullMessage = "エラーが発生しました: " & errorMessage
    If details <> "" Then
        fullMessage = fullMessage & vbCrLf & vbCrLf & "詳細: " & details
    End If
    MsgBox fullMessage, vbCritical, APP_NAME & " - エラー"
End Sub

' 成功メッセージ表示
Public Sub ShowSuccess(ByVal message As String)
    MsgBox message, vbInformation, APP_NAME & " - 完了"
End Sub

' ログ出力（デバッグ用）
Public Sub WriteLog(ByVal message As String, Optional ByVal logLevel As String = "INFO")
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & logLevel & "] " & message
End Sub

' 文字列の清浄化（HTML タグ除去等）
Public Function CleanText(ByVal inputText As String) As String
    Dim cleanedText As String
    cleanedText = inputText
    
    ' HTMLタグの除去
    cleanedText = RemoveHtmlTags(cleanedText)
    
    ' 余分な空白や改行の除去
    cleanedText = Trim(cleanedText)
    cleanedText = Replace(cleanedText, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    
    CleanText = cleanedText
End Function

' HTMLタグ除去
Private Function RemoveHtmlTags(ByVal htmlText As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "<[^>]*>"
    
    RemoveHtmlTags = regex.Replace(htmlText, "")
End Function

' 文字数制限チェック
Public Function CheckContentLength(ByVal content As String) As Boolean
    If Len(content) > MAX_CONTENT_LENGTH Then
        ShowMessage "メール内容が大きすぎます（" & Len(content) & "文字）。" & vbCrLf & _
                   "最大" & MAX_CONTENT_LENGTH & "文字まで処理可能です。", "制限超過", vbExclamation
        CheckContentLength = False
    Else
        CheckContentLength = True
    End If
End Function

' 選択されたメールアイテムの取得
Public Function GetSelectedMailItem() As Object
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim olSelection As Object
    
    Set olApp = Application
    Set olSelection = olApp.ActiveExplorer.Selection
    
    If olSelection.Count = 0 Then
        ShowMessage "メールを選択してください。", "選択エラー", vbExclamation
        Set GetSelectedMailItem = Nothing
        Exit Function
    End If
    
    If olSelection.Count > 1 Then
        ShowMessage "複数のメールが選択されています。1つのメールを選択してください。", "選択エラー", vbExclamation
        Set GetSelectedMailItem = Nothing
        Exit Function
    End If
    
    ' メールアイテムかどうかをチェック
    If olSelection.Item(1).Class = olMail Then
        Set GetSelectedMailItem = olSelection.Item(1)
    Else
        ShowMessage "選択されたアイテムはメールではありません。", "選択エラー", vbExclamation
        Set GetSelectedMailItem = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "メールアイテムの取得中にエラーが発生しました。", Err.Description
    Set GetSelectedMailItem = Nothing
End Function

' プログレス表示（簡易版）
Public Sub ShowProgress(ByVal message As String)
    WriteLog "進行状況: " & message
    ' 実際の実装では、プログレスバーまたはステータス表示を行う
    DoEvents ' UIの応答性を保つ
End Sub

' JSON エスケープ処理
Public Function EscapeJsonString(ByVal inputString As String) As String
    Dim result As String
    result = inputString
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJsonString = result
End Function

' 設定値の検証
Public Function ValidateConfiguration() As Boolean
    ' API Key の確認
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Or OPENAI_API_KEY = "" Then
        ShowError "OpenAI API キーが設定されていません。", _
                 "AI_Common.bas の OPENAI_API_KEY を設定してください。"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' エンドポイントの確認
    If InStr(OPENAI_API_ENDPOINT, "your-azure-openai-endpoint") > 0 Then
        ShowError "OpenAI API エンドポイントが設定されていません。", _
                 "AI_Common.bas の OPENAI_API_ENDPOINT を設定してください。"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
End Function

' メインメニュー表示
Public Sub ShowMainMenu()
    Dim choice As String
    Dim menuText As String
    
    menuText = APP_NAME & " v" & APP_VERSION & vbCrLf & vbCrLf & _
               "利用可能な機能:" & vbCrLf & _
               "1. メール内容を解析" & vbCrLf & _
               "2. 営業断りメールを作成" & vbCrLf & _
               "3. 承諾メールを作成" & vbCrLf & _
               "4. 検索フォルダを分析" & vbCrLf & _
               "5. 設定確認" & vbCrLf & vbCrLf & _
               "実行したい機能の番号を入力してください:"
    
    choice = InputBox(menuText, APP_NAME & " - メインメニュー")
    
    Select Case choice
        Case "1"
            Call AI_EmailAnalyzer.AnalyzeSelectedEmail
        Case "2"
            Call AI_EmailComposer.CreateRejectionEmail
        Case "3"
            Call AI_EmailComposer.CreateAcceptanceEmail
        Case "4"
            Call AI_SearchAnalyzer.AnalyzeSearchFolders
        Case "5"
            Call ShowConfigurationInfo
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。1-5の番号を入力してください。", "入力エラー", vbExclamation
    End Select
End Sub

' 設定情報表示
Private Sub ShowConfigurationInfo()
    Dim configInfo As String
    configInfo = "現在の設定:" & vbCrLf & vbCrLf & _
                "アプリケーション名: " & APP_NAME & vbCrLf & _
                "バージョン: " & APP_VERSION & vbCrLf & _
                "API エンドポイント: " & OPENAI_API_ENDPOINT & vbCrLf & _
                "API キー: " & Left(OPENAI_API_KEY, 10) & "..." & vbCrLf & _
                "最大処理文字数: " & MAX_CONTENT_LENGTH & vbCrLf & _
                "タイムアウト: " & REQUEST_TIMEOUT & "秒"
    
    ShowMessage configInfo, "設定情報"
End Sub