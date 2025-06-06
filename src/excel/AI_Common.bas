' AI_Common.bas
' Excel AI Helper - 共通関数・定数モジュール
' 作成日: 2024
' 説明: 全モジュールで使用する共通関数、定数、設定を管理

Option Explicit

' =============================================================================
' 定数定義
' =============================================================================

' Azure OpenAI API 設定
Public Const OPENAI_API_ENDPOINT As String = "https://your-azure-openai-endpoint.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview"
Public Const OPENAI_API_KEY As String = "YOUR_API_KEY_HERE" ' 本番環境では設定から読み込み
Public Const OPENAI_MODEL As String = "gpt-4"

' アプリケーション設定
Public Const APP_NAME As String = "Excel AI Helper"
Public Const APP_VERSION As String = "1.0.0"
Public Const MAX_CONTENT_LENGTH As Integer = 50000 ' 最大処理文字数
Public Const REQUEST_TIMEOUT As Integer = 60 ' APIリクエストタイムアウト（秒）
Public Const MAX_ROWS_PER_BATCH As Integer = 1000 ' 1回の処理での最大行数

' プロンプトテンプレート
Public Const SYSTEM_PROMPT_TABLE_ANALYZER As String = "あなたは日本語のデータ分析の専門家です。Excelの表データを分析し、指定された観点で各行を評価してください。"
Public Const SYSTEM_PROMPT_CELL_PROCESSOR As String = "あなたは日本語のExcel関数とデータ処理の専門家です。ユーザーの要求に応じて適切な関数やデータ変換を提案してください。"

' メッセージテンプレート
Public Const MSG_API_KEY_NOT_SET As String = "Azure OpenAI APIキーが設定されていません。設定から APIキーを設定してください。"
Public Const MSG_PROCESSING_COMPLETE As String = "AI処理が完了しました。"
Public Const MSG_PROCESSING_ERROR As String = "AI処理中にエラーが発生しました："

' =============================================================================
' 共通ユーティリティ関数
' =============================================================================

' メッセージボックス表示（成功）
Public Sub ShowSuccess(ByVal message As String)
    MsgBox message, vbInformation + vbOKOnly, APP_NAME
End Sub

' メッセージボックス表示（警告）
Public Sub ShowWarning(ByVal message As String)
    MsgBox message, vbExclamation + vbOKOnly, APP_NAME
End Sub

' メッセージボックス表示（エラー）
Public Sub ShowError(ByVal title As String, ByVal message As String)
    MsgBox title & vbCrLf & vbCrLf & message, vbCritical + vbOKOnly, APP_NAME
End Sub

' 確認ダイアログ表示
Public Function ShowConfirm(ByVal message As String) As Boolean
    ShowConfirm = (MsgBox(message, vbQuestion + vbYesNo, APP_NAME) = vbYes)
End Function

' =============================================================================
' Azure OpenAI API 通信機能
' =============================================================================

' Azure OpenAI APIへのリクエスト送信
Public Function SendOpenAIRequest(ByVal systemPrompt As String, ByVal userMessage As String) As String
    On Error GoTo ErrorHandler
    
    ' 設定値検証
    If Not ValidateConfiguration() Then
        SendOpenAIRequest = ""
        Exit Function
    End If
    
    ' HTTP リクエストオブジェクトの作成
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' リクエストボディの作成
    Dim requestBody As String
    requestBody = CreateRequestBody(systemPrompt, userMessage)
    
    ' HTTPリクエストの設定
    http.Open "POST", GetAPIEndpoint(), False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "api-key", GetAPIKey()
    
    ' リクエスト送信
    http.send requestBody
    
    ' レスポンス処理
    If http.status = 200 Then
        SendOpenAIRequest = ParseResponse(http.responseText)
    Else
        ShowError "API通信エラー", "ステータスコード: " & http.status & vbCrLf & "レスポンス: " & http.responseText
        SendOpenAIRequest = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "API通信中にエラーが発生しました", Err.Description
    SendOpenAIRequest = ""
End Function

' リクエストボディの作成
Private Function CreateRequestBody(ByVal systemPrompt As String, ByVal userMessage As String) As String
    Dim requestBody As String
    requestBody = "{" & _
                 """messages"": [" & _
                 "{""role"": ""system"", ""content"": """ & EscapeJsonString(systemPrompt) & """}," & _
                 "{""role"": ""user"", ""content"": """ & EscapeJsonString(userMessage) & """}" & _
                 "]," & _
                 """max_tokens"": 2000," & _
                 """temperature"": 0.7" & _
                 "}"
    CreateRequestBody = requestBody
End Function

' JSONレスポンスの解析
Private Function ParseResponse(ByVal responseJson As String) As String
    On Error GoTo ErrorHandler
    
    ' 簡易JSON解析（本格実装では専用ライブラリを使用）
    Dim startPos As Integer
    Dim endPos As Integer
    
    ' "content"フィールドの値を抽出
    startPos = InStr(responseJson, """content"":""") + 11
    If startPos > 11 Then
        endPos = InStr(startPos, responseJson, """,""")
        If endPos = 0 Then
            endPos = InStr(startPos, responseJson, """}")
        End If
        
        If endPos > startPos Then
            ParseResponse = Mid(responseJson, startPos, endPos - startPos)
            ' エスケープ文字の処理
            ParseResponse = Replace(ParseResponse, "\\n", vbCrLf)
            ParseResponse = Replace(ParseResponse, "\\""", """")
            ParseResponse = Replace(ParseResponse, "\\\\", "\")
        Else
            ParseResponse = ""
        End If
    Else
        ParseResponse = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "レスポンス解析エラー", Err.Description
    ParseResponse = ""
End Function

' JSON文字列のエスケープ処理
Public Function EscapeJsonString(ByVal inputString As String) As String
    Dim result As String
    result = inputString
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\\n")
    result = Replace(result, vbCr, "\\n")
    result = Replace(result, vbLf, "\\n")
    result = Replace(result, vbTab, "\\t")
    EscapeJsonString = result
End Function

' =============================================================================
' 設定取得関数
' =============================================================================

' APIエンドポイント取得
Public Function GetAPIEndpoint() As String
    ' AI_ConfigManager.basの GetAPISetting を使用
    GetAPIEndpoint = GetAPISetting("Endpoint", OPENAI_API_ENDPOINT)
End Function

' APIキー取得
Public Function GetAPIKey() As String
    ' AI_ConfigManager.basの GetAPISetting を使用
    GetAPIKey = GetAPISetting("APIKey", OPENAI_API_KEY)
End Function

' レジストリアクセス用の設定取得関数（AI_ConfigManager.basからコピー）
Private Function GetAPISetting(ByVal key As String, ByVal defaultValue As String) As String
    On Error GoTo ErrorHandler
    
    Dim registryPath As String
    registryPath = "HKEY_CURRENT_USER\Software\ExcelAIHelper\"
    
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    GetAPISetting = shell.RegRead(registryPath & key)
    
    Exit Function
    
ErrorHandler:
    ' エラーの場合はデフォルト値を返す
    GetAPISetting = defaultValue
End Function

' 設定値の検証
Public Function ValidateConfiguration() As Boolean
    ' API Key の確認
    If GetAPIKey() = "YOUR_API_KEY_HERE" Or GetAPIKey() = "" Then
        ShowError "API設定エラー", MSG_API_KEY_NOT_SET
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' エンドポイントの確認
    If GetAPIEndpoint() = "" Then
        ShowError "API設定エラー", "APIエンドポイントが設定されていません。"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
End Function

' =============================================================================
' Excel関連ユーティリティ
' =============================================================================

' 選択範囲の検証
Public Function ValidateSelection() As Boolean
    On Error GoTo ErrorHandler
    
    If Selection Is Nothing Then
        ShowWarning "範囲が選択されていません。"
        ValidateSelection = False
        Exit Function
    End If
    
    If TypeName(Selection) <> "Range" Then
        ShowWarning "セル範囲を選択してください。"
        ValidateSelection = False
        Exit Function
    End If
    
    ValidateSelection = True
    Exit Function
    
ErrorHandler:
    ShowError "選択範囲の検証エラー", Err.Description
    ValidateSelection = False
End Function

' 処理進捗表示
Public Sub ShowProgress(ByVal currentRow As Integer, ByVal totalRows As Integer, ByVal message As String)
    Dim progressMsg As String
    progressMsg = message & vbCrLf & _
                 "進捗: " & currentRow & "/" & totalRows & " (" & Format(currentRow / totalRows, "0%") & ")"
    
    Application.StatusBar = progressMsg
    DoEvents
End Sub

' 進捗表示クリア
Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

' =============================================================================
' ログ機能
' =============================================================================

' ログメッセージの出力
Public Sub LogMessage(ByVal level As String, ByVal message As String)
    ' デバッグ用：イミディエイトウィンドウに出力
    Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " [" & level & "] " & message
End Sub

' エラーログの出力
Public Sub LogError(ByVal functionName As String, ByVal errorMessage As String)
    LogMessage "ERROR", functionName & ": " & errorMessage
End Sub

' 情報ログの出力
Public Sub LogInfo(ByVal message As String)
    LogMessage "INFO", message
End Sub