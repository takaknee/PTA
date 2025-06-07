' AI_ApiConnector.bas
' Outlook OpenAI Macro - OpenAI API Connection Module
' Created: 2024
' Description: Manages communication with Azure OpenAI API

Option Explicit

' =============================================================================
' Constants Definition (Copied from AI_Common.bas for independence)
' =============================================================================

' OpenAI API Settings
Public Const OPENAI_API_ENDPOINT As String = "https://your-azure-openai-endpoint.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview"
Public Const OPENAI_API_KEY As String = "YOUR_API_KEY_HERE" ' Production environment should read from configuration file
Public Const OPENAI_MODEL As String = "gpt-4"

' Application Settings
Public Const APP_NAME As String = "Outlook AI Helper"
Public Const REQUEST_TIMEOUT As Integer = 30 ' API request timeout (seconds)

' =============================================================================
' Common Functions (Required for this module)
' =============================================================================

' Display error message
Public Sub ShowError(ByVal errorMessage As String, Optional ByVal details As String = "")
    Dim fullMessage As String
    fullMessage = "An error occurred: " & errorMessage
    If details <> "" Then
        fullMessage = fullMessage & vbCrLf & vbCrLf & "Details: " & details
    End If
    MsgBox fullMessage, vbCritical, APP_NAME & " - Error"
End Sub

' Display message box (common format)
Public Sub ShowMessage(ByVal message As String, ByVal title As String, Optional ByVal messageType As VbMsgBoxStyle = vbInformation)
    MsgBox message, messageType, APP_NAME & " - " & title
End Sub

' Progress display (simplified)
Public Sub ShowProgress(ByVal message As String)
    ' For debugging purposes
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " [PROGRESS] " & message
    DoEvents ' Maintain UI responsiveness
End Sub

' Configuration validation
Public Function ValidateConfiguration() As Boolean
    ' Check API Key
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Or OPENAI_API_KEY = "" Then
        ShowError "OpenAI API key is not configured.", _
                 "Please set OPENAI_API_KEY constant."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' Check endpoint
    If InStr(OPENAI_API_ENDPOINT, "your-azure-openai-endpoint") > 0 Then
        ShowError "OpenAI API endpoint is not configured.", _
                 "Please set OPENAI_API_ENDPOINT constant."
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
End Function

' Log output (for debugging)
Public Sub WriteLog(ByVal message As String, Optional ByVal logLevel As String = "INFO")
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & logLevel & "] " & message
End Sub

' =============================================================================
' OpenAI API Connection Functions
' =============================================================================

' Send request to OpenAI API
Public Function SendOpenAIRequest(ByVal systemPrompt As String, ByVal userMessage As String, Optional ByVal maxTokens As Integer = 1000) As String
    On Error GoTo ErrorHandler
    
    ' Configuration validation
    If Not ValidateConfiguration() Then
        SendOpenAIRequest = ""
        Exit Function
    End If
    
    ShowProgress "Starting AI analysis..."
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Create request body
    Dim requestBody As String
    requestBody = CreateRequestBody(systemPrompt, userMessage, maxTokens)
    
    WriteLog "Sending API request: " & Left(userMessage, 100) & "..."
    
    ' Configure HTTP request
    http.Open "POST", OPENAI_API_ENDPOINT, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & OPENAI_API_KEY
    http.setRequestHeader "User-Agent", APP_NAME
    
    ShowProgress "Connecting to API..."
    
    ' リクエスト送信
    http.send requestBody
    
    ShowProgress "AI処理中..."
    
    ' レスポンスの処理
    If http.Status = 200 Then
        Dim response As String
        response = http.responseText
        WriteLog "API レスポンス受信成功"
        
        ' レスポンスからメッセージ内容を抽出
        Dim result As String
        result = ParseOpenAIResponse(response)
        
        If result <> "" Then
            ShowProgress "AI分析完了"
            SendOpenAIRequest = result
        Else
            ShowError "APIレスポンスの解析に失敗しました。", response
            SendOpenAIRequest = ""
        End If
    Else
        ShowError "API リクエストが失敗しました。", _
                 "ステータスコード: " & http.Status & vbCrLf & _
                 "レスポンス: " & http.responseText
        WriteLog "API エラー: " & http.Status & " - " & http.responseText, "ERROR"
        SendOpenAIRequest = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "API接続中にエラーが発生しました。", Err.Description
    WriteLog "API接続エラー: " & Err.Description, "ERROR"
    SendOpenAIRequest = ""
End Function

' リクエストボディの作成
Private Function CreateRequestBody(ByVal systemPrompt As String, ByVal userMessage As String, ByVal maxTokens As Integer) As String
    Dim body As String
    
    ' JSON形式でリクエストボディを作成
    body = "{" & _
           """model"": """ & OPENAI_MODEL & """," & _
           """messages"": [" & _
               "{""role"": ""system"", ""content"": """ & EscapeJsonString(systemPrompt) & """}," & _
               "{""role"": ""user"", ""content"": """ & EscapeJsonString(userMessage) & """}" & _
           "]," & _
           """max_tokens"": " & maxTokens & "," & _
           """temperature"": 0.3," & _
           """top_p"": 1.0," & _
           """frequency_penalty"": 0.0," & _
           """presence_penalty"": 0.0" & _
           "}"
    
    CreateRequestBody = body
End Function

' OpenAI APIレスポンスの解析
Private Function ParseOpenAIResponse(ByVal responseText As String) As String
    On Error GoTo ErrorHandler
    
    ' 簡易JSON解析（VBAの制限により正規表現を使用）
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    
    ' content フィールドを抽出
    regex.Pattern = """content"":\s*""([^""]*(?:\\.[^""]*)*)"""
    
    Dim matches As Object
    Set matches = regex.Execute(responseText)
    
    If matches.Count > 0 Then
        Dim content As String
        content = matches(0).SubMatches(0)
        
        ' エスケープ文字の復元
        content = Replace(content, "\n", vbCrLf)
        content = Replace(content, "\t", vbTab)
        content = Replace(content, "\""", """")
        content = Replace(content, "\\", "\")
        
        ParseOpenAIResponse = content
    Else
        ' エラーレスポンスのチェック
        regex.Pattern = """error"":\s*{[^}]*""message"":\s*""([^""]*(?:\\.[^""]*)*)"""
        Set matches = regex.Execute(responseText)
        
        If matches.Count > 0 Then
            ShowError "OpenAI APIエラー", matches(0).SubMatches(0)
        Else
            ShowError "不明なAPIレスポンス形式です。", responseText
        End If
        
        ParseOpenAIResponse = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "APIレスポンスの解析中にエラーが発生しました。", Err.Description
    ParseOpenAIResponse = ""
End Function

' API接続テスト
Public Sub TestAPIConnection()
    On Error GoTo ErrorHandler
    
    If Not ValidateConfiguration() Then
        Exit Sub
    End If
    
    ShowProgress "API接続テストを実行中..."
    
    Dim testPrompt As String
    testPrompt = "こんにちは"
    
    Dim result As String
    result = SendOpenAIRequest("あなたは親切なアシスタントです。", testPrompt, 50)
    
    If result <> "" Then
        ShowSuccess "API接続テスト成功!" & vbCrLf & vbCrLf & "レスポンス: " & result
        WriteLog "API接続テスト成功"
    Else
        ShowError "API接続テストに失敗しました。設定を確認してください。"
        WriteLog "API接続テスト失敗", "ERROR"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "API接続テスト中にエラーが発生しました。", Err.Description
    WriteLog "API接続テストエラー: " & Err.Description, "ERROR"
End Sub

' 複数回に分けてリクエストを送信（長いテキスト用）
Public Function SendOpenAIRequestChunked(ByVal systemPrompt As String, ByVal userMessage As String, Optional ByVal chunkSize As Integer = 3000) As String
    On Error GoTo ErrorHandler
    
    If Len(userMessage) <= chunkSize Then
        ' 分割不要
        SendOpenAIRequestChunked = SendOpenAIRequest(systemPrompt, userMessage)
        Exit Function
    End If
    
    ShowProgress "大きなテキストを分割して処理中..."
    
    Dim chunks() As String
    Dim chunkCount As Integer
    chunkCount = 0
    
    ' テキストを分割
    Dim i As Integer
    For i = 1 To Len(userMessage) Step chunkSize
        ReDim Preserve chunks(chunkCount)
        chunks(chunkCount) = Mid(userMessage, i, chunkSize)
        chunkCount = chunkCount + 1
    Next i
    
    ' 各チャンクを処理
    Dim results() As String
    ReDim results(chunkCount - 1)
    
    For i = 0 To chunkCount - 1
        ShowProgress "チャンク " & (i + 1) & "/" & chunkCount & " を処理中..."
        
        Dim chunkPrompt As String
        chunkPrompt = systemPrompt & " (この内容は " & chunkCount & " 個に分割されたテキストの " & (i + 1) & " 番目です。)"
        
        results(i) = SendOpenAIRequest(chunkPrompt, chunks(i))
        
        If results(i) = "" Then
            ShowError "チャンク " & (i + 1) & " の処理に失敗しました。"
            SendOpenAIRequestChunked = ""
            Exit Function
        End If
        
        ' APIレート制限を考慮した待機
        Application.Wait (Now + TimeValue("0:00:02"))
    Next i
    
    ' 結果を結合
    Dim combinedResult As String
    combinedResult = "=== 分割処理結果 ===" & vbCrLf & vbCrLf
    
    For i = 0 To chunkCount - 1
        combinedResult = combinedResult & "=== パート " & (i + 1) & " ===" & vbCrLf & _
                        results(i) & vbCrLf & vbCrLf
    Next i
    
    SendOpenAIRequestChunked = combinedResult
    
    Exit Function
    
ErrorHandler:
    ShowError "分割処理中にエラーが発生しました。", Err.Description
    SendOpenAIRequestChunked = ""
End Function

' APIレスポンス統計
Public Sub ShowAPIStatistics()
    ' 簡易的な統計情報表示
    Dim statsMessage As String
    statsMessage = "OpenAI API 統計情報:" & vbCrLf & vbCrLf & _
                  "接続エンドポイント: " & OPENAI_API_ENDPOINT & vbCrLf & _
                  "使用モデル: " & OPENAI_MODEL & vbCrLf & _
                  "デフォルトタイムアウト: " & REQUEST_TIMEOUT & "秒" & vbCrLf & _
                  "最大処理文字数: " & MAX_CONTENT_LENGTH & "文字"
    
    ShowMessage statsMessage, "API統計情報"
End Sub