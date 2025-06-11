' OutlookAI_Unified_Simplified.bas
' Outlook OpenAI マクロ - 簡素化統合版
' 作成日: 2024
' 説明: 重要機能に集約した簡潔な統合版
' 
' 主要機能:
' - メール内容解析
' - 営業断りメール作成  
' - 承諾メール作成
' - カスタムメール作成
' - 設定管理
' - API接続テスト

Option Explicit

' =============================================================================
' 定数定義
' =============================================================================

' OpenAI API 設定
Public Const OPENAI_API_ENDPOINT As String = "https://your-azure-openai-endpoint.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview"
Public Const OPENAI_API_KEY As String = "YOUR_API_KEY_HERE"
Public Const OPENAI_MODEL As String = "gpt-4"

' アプリケーション設定
Public Const APP_NAME As String = "Outlook AI Helper"
Public Const APP_VERSION As String = "1.0.0 Simplified"
Public Const MAX_CONTENT_LENGTH As Long = 50000
Public Const REQUEST_TIMEOUT As Integer = 30
Public Const MAX_TOKEN As Integer = 2000

' プロンプトテンプレート
Public Const SYSTEM_PROMPT_ANALYZER As String = "あなたは日本語のビジネスメール分析の専門家です。メール内容を分析し、重要な情報を抽出してください。"
Public Const SYSTEM_PROMPT_COMPOSER As String = "あなたは日本語のビジネスメール作成の専門家です。適切で丁寧なビジネスメールを作成してください。"

' =============================================================================
' メインメニュー関数
' =============================================================================

' メインメニュー表示
Public Sub ShowMainMenu()
    Dim choice As String
    Dim menuText As String
    
    menuText = "🤖 " & APP_NAME & " v" & APP_VERSION & vbCrLf & vbCrLf & "利用可能な機能:" & vbCrLf & "1. メール内容を解析" & vbCrLf & "2. 営業断りメール作成" & vbCrLf & "3. 承諾メール作成" & vbCrLf & "4. カスタムメール作成" & vbCrLf & "5. 設定管理" & vbCrLf & "6. API接続テスト" & vbCrLf & vbCrLf & "実行したい機能の番号を入力してください:"
    
    choice = InputBox(menuText, APP_NAME & " - メインメニュー")
    
    Select Case choice
        Case "1": Call AnalyzeSelectedEmail
        Case "2": Call CreateRejectionEmail
        Case "3": Call CreateAcceptanceEmail
        Case "4": Call CreateCustomBusinessEmail
        Case "5": Call ManageConfiguration
        Case "6": Call TestAPIConnection
        Case "": ' キャンセル時は何もしない
        Case Else: ShowMessage "無効な選択です。1-6の番号を入力してください。", "入力エラー", vbExclamation
    End Select
End Sub

' =============================================================================
' 共通関数
' =============================================================================

' メッセージボックス表示
Public Sub ShowMessage(ByVal message As String, ByVal title As String, Optional ByVal messageType As VbMsgBoxStyle = vbInformation)
    MsgBox message, messageType, APP_NAME & " - " & title
End Sub

' エラーメッセージ表示
Public Sub ShowError(ByVal errorMessage As String, Optional ByVal details As String = "")
    Dim fullMessage As String
    fullMessage = "エラーが発生しました: " & errorMessage
    If details <> "" Then fullMessage = fullMessage & vbCrLf & vbCrLf & "詳細: " & details
    MsgBox fullMessage, vbCritical, APP_NAME & " - エラー"
End Sub

' 成功メッセージ表示
Public Sub ShowSuccess(ByVal message As String)
    MsgBox message, vbInformation, APP_NAME & " - 完了"
End Sub

' プログレス表示（簡易版）
Public Sub ShowProgress(ByVal message As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " [進行状況] " & message
    DoEvents
End Sub

' 設定値の検証
Public Function ValidateConfiguration() As Boolean
    If OPENAI_API_KEY = "YOUR_API_KEY_HERE" Or OPENAI_API_KEY = "" Then
        ShowError "OpenAI API キーが設定されていません。", "このファイルの OPENAI_API_KEY 定数を設定してください。"
        ValidateConfiguration = False
        Exit Function
    End If
    
    If InStr(OPENAI_API_ENDPOINT, "your-azure-openai-endpoint") > 0 Then
        ShowError "OpenAI API エンドポイントが設定されていません。", "このファイルの OPENAI_API_ENDPOINT 定数を設定してください。"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
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

' メール本文テキストの取得
Private Function GetEmailBodyText(ByVal mailItem As Object) As String
    On Error GoTo ErrorHandler
    
    Dim bodyText As String
    
    If Len(mailItem.Body) > 0 Then
        bodyText = mailItem.Body
    ElseIf Len(mailItem.HTMLBody) > 0 Then
        bodyText = CleanText(mailItem.HTMLBody)
    Else
        bodyText = "（メール本文が空です）"
    End If
    
    GetEmailBodyText = bodyText
    Exit Function
    
ErrorHandler:
    ShowError "メール本文の取得中にエラーが発生しました。", Err.Description
    GetEmailBodyText = ""
End Function

' 文字列の清浄化（HTMLタグ除去等）
Public Function CleanText(ByVal inputText As String) As String
    Dim cleanedText As String
    cleanedText = inputText
    
    ' HTMLタグの除去（簡易版）
    Dim i As Integer
    Dim inTag As Boolean
    Dim result As String
    
    For i = 1 To Len(cleanedText)
        Dim char As String
        char = Mid(cleanedText, i, 1)
        
        If char = "<" Then
            inTag = True
        ElseIf char = ">" Then
            inTag = False
        ElseIf Not inTag Then
            result = result & char
        End If
    Next i
    
    ' 余分な空白や改行の除去
    result = Trim(result)
    Do While InStr(result, vbCrLf & vbCrLf & vbCrLf) > 0
        result = Replace(result, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Loop
    
    CleanText = result
End Function

' 文字数制限チェック
Public Function CheckContentLength(ByVal content As String) As Boolean
    If Len(content) > MAX_CONTENT_LENGTH Then
        ShowMessage "メール内容が大きすぎます（" & Len(content) & "文字）。最大" & MAX_CONTENT_LENGTH & "文字まで処理可能です。", "制限超過", vbExclamation
        CheckContentLength = False
    Else
        CheckContentLength = True
    End If
End Function

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

' =============================================================================
' OpenAI API 接続関数
' =============================================================================

' OpenAI API にリクエストを送信
Public Function SendOpenAIRequest(ByVal systemPrompt As String, ByVal userMessage As String, Optional ByVal maxTokens As Integer = 1000) As String
    On Error GoTo ErrorHandler
    
    If Not ValidateConfiguration() Then
        SendOpenAIRequest = ""
        Exit Function
    End If
    
    ShowProgress "AI分析を開始しています..."
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' リクエストボディの作成
    Dim requestBody As String
    requestBody = "{""model"": """ & OPENAI_MODEL & """, ""messages"": [{""role"": ""system"", ""content"": """ & EscapeJsonString(systemPrompt) & """}, {""role"": ""user"", ""content"": """ & EscapeJsonString(userMessage) & """}], ""max_tokens"": " & maxTokens & ", ""temperature"": 0.7}"
    
    ' HTTPリクエストの設定
    http.Open "POST", OPENAI_API_ENDPOINT, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & OPENAI_API_KEY
    http.setRequestHeader "User-Agent", APP_NAME & "/" & APP_VERSION
    
    ShowProgress "APIに接続中..."
    http.send requestBody
    ShowProgress "AI処理中..."
    
    ' レスポンスの処理
    If http.Status = 200 Then
        Dim response As String
        response = http.responseText
        
        ' JSONレスポンスの解析（簡易版）
        Dim result As String
        result = ParseOpenAIResponse(response)
        
        ShowProgress "AI分析完了"
        SendOpenAIRequest = result
    Else
        ShowError "OpenAI API接続エラーが発生しました。", "ステータスコード: " & http.Status & vbCrLf & "エラー内容: " & http.statusText & vbCrLf & "設定を確認してください。"
        SendOpenAIRequest = ""
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "API接続中にエラーが発生しました。", Err.Description
    SendOpenAIRequest = ""
End Function

' OpenAI レスポンスの解析（簡易版）
Private Function ParseOpenAIResponse(ByVal jsonResponse As String) As String
    On Error GoTo ErrorHandler
    
    Dim contentStart As Integer
    Dim contentEnd As Integer
    Dim content As String
    
    ' "content": "..." の部分を抽出
    contentStart = InStr(jsonResponse, """content"": """)
    If contentStart > 0 Then
        contentStart = contentStart + 12
        contentEnd = InStr(contentStart, jsonResponse, """, """)
        If contentEnd = 0 Then contentEnd = InStr(contentStart, jsonResponse, """}")
        If contentEnd = 0 Then contentEnd = InStr(contentStart, jsonResponse, """}]")
    End If
    
    If contentStart > 0 And contentEnd > contentStart Then
        content = Mid(jsonResponse, contentStart, contentEnd - contentStart)
        content = Replace(content, "\""", """")
        content = Replace(content, "\\", "\")
        content = Replace(content, "\n", vbCrLf)
        content = Replace(content, "\t", vbTab)
        ParseOpenAIResponse = content
    Else
        ParseOpenAIResponse = "コンテンツが見つかりません"
    End If
    
    Exit Function
    
ErrorHandler:
    ParseOpenAIResponse = "解析エラーが発生しました"
End Function

' API接続テスト
Public Sub TestAPIConnection()
    On Error GoTo ErrorHandler
    
    If Not ValidateConfiguration() Then Exit Sub
    
    ShowProgress "API接続をテスト中..."
    
    Dim result As String
    result = SendOpenAIRequest("簡潔に日本語で応答してください。", "こんにちは", 50)
    
    If result <> "" Then
        ShowSuccess "API接続テストが成功しました！" & vbCrLf & vbCrLf & "応答: " & result
    Else
        ShowError "API接続テストが失敗しました。", "設定を確認してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "API接続テスト中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' メール解析機能
' =============================================================================

' 選択されたメールを解析
Public Sub AnalyzeSelectedEmail()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Set mailItem = GetSelectedMailItem()
    
    If mailItem Is Nothing Then Exit Sub
    
    ShowProgress "メール内容を分析中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then Exit Sub
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & "以下のメールの内容を分析し、重要な情報を日本語で簡潔に要約してください。"
    
    Dim userMessage As String
    userMessage = "件名: " & mailItem.Subject & vbCrLf & "送信者: " & mailItem.SenderName & vbCrLf & "本文: " & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ShowMessage result, "メール分析結果"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "メール解析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' メール作成支援機能
' =============================================================================

' 営業断りメール作成
Public Sub CreateRejectionEmail()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Set mailItem = GetSelectedMailItem()
    
    If mailItem Is Nothing Then Exit Sub
    
    Dim rejectionReason As String
    rejectionReason = InputBox("断り理由を入力してください（省略可）:", APP_NAME & " - 断り理由", "")
    
    ShowProgress "営業断りメールを作成中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then Exit Sub
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_COMPOSER & vbCrLf & "営業メールに対する丁寧で適切な断りメールを日本語で作成してください。"
    
    Dim userMessage As String
    userMessage = "元メール件名: " & mailItem.Subject & vbCrLf & "元メール送信者: " & mailItem.SenderName & vbCrLf & "断り理由: " & rejectionReason & vbCrLf & "元メール本文: " & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        CreateReplyEmail mailItem, "Re: " & mailItem.Subject, result
        ShowSuccess "営業断りメールの返信ウィンドウを開きました。内容を確認してから送信してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "営業断りメール作成中にエラーが発生しました。", Err.Description
End Sub

' 承諾メール作成
Public Sub CreateAcceptanceEmail()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Set mailItem = GetSelectedMailItem()
    
    If mailItem Is Nothing Then Exit Sub
    
    Dim acceptanceDetails As String
    acceptanceDetails = InputBox("承諾内容の詳細を入力してください（省略可）:", APP_NAME & " - 承諾詳細", "")
    
    ShowProgress "承諾メールを作成中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then Exit Sub
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_COMPOSER & vbCrLf & "依頼や提案に対する前向きで適切な承諾メールを日本語で作成してください。"
    
    Dim userMessage As String
    userMessage = "元メール件名: " & mailItem.Subject & vbCrLf & "元メール送信者: " & mailItem.SenderName & vbCrLf & "承諾詳細: " & acceptanceDetails & vbCrLf & "元メール本文: " & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        CreateReplyEmail mailItem, "Re: " & mailItem.Subject, result
        ShowSuccess "承諾メールの返信ウィンドウを開きました。内容を確認してから送信してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "承諾メール作成中にエラーが発生しました。", Err.Description
End Sub

' カスタムメール作成
Public Sub CreateCustomBusinessEmail()
    On Error GoTo ErrorHandler
    
    Dim emailType As String
    emailType = InputBox("作成するメールの種類を選択してください：" & vbCrLf & "1. お礼メール" & vbCrLf & "2. 謝罪メール" & vbCrLf & "3. 問い合わせメール" & vbCrLf & "4. 提案・依頼メール" & vbCrLf & "5. その他" & vbCrLf & "番号を入力してください:", APP_NAME & " - メール作成")
    
    If emailType = "" Then Exit Sub
    
    Dim emailDetails As String
    emailDetails = InputBox("メールの詳細情報を入力してください：" & vbCrLf & "・宛先（相手の名前、役職等）" & vbCrLf & "・目的・内容" & vbCrLf & "・背景情報", APP_NAME & " - 詳細情報")
    
    If emailDetails = "" Then
        ShowMessage "詳細情報が入力されていません。", "入力エラー", vbExclamation
        Exit Sub
    End If
    
    ShowProgress "カスタムメールを作成中..."
    
    Dim emailTypeText As String
    Select Case emailType
        Case "1": emailTypeText = "お礼メール"
        Case "2": emailTypeText = "謝罪メール"
        Case "3": emailTypeText = "問い合わせメール"
        Case "4": emailTypeText = "提案・依頼メール"
        Case Else: emailTypeText = "ビジネスメール"
    End Select
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_COMPOSER & vbCrLf & "適切なビジネス敬語を使用して" & emailTypeText & "を日本語で作成してください。"
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, emailDetails, MAX_TOKEN)
    
    If result <> "" Then
        CreateNewEmail "（AI生成）" & emailTypeText, result
        ShowSuccess "カスタムメールの編集ウィンドウを開きました。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムメール作成中にエラーが発生しました。", Err.Description
End Sub

' 返信メールの作成
Private Sub CreateReplyEmail(ByVal originalMail As Object, ByVal subject As String, ByVal body As String)
    On Error GoTo ErrorHandler
    
    Dim replyMail As Object
    Set replyMail = originalMail.Reply
    
    replyMail.Subject = subject
    replyMail.Body = body
    replyMail.Display
    
    Exit Sub
    
ErrorHandler:
    ShowError "返信メール作成中にエラーが発生しました。", Err.Description
End Sub

' 新規メールの作成
Private Sub CreateNewEmail(ByVal subject As String, ByVal body As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim newMail As Object
    
    Set olApp = Application
    Set newMail = olApp.CreateItem(olMailItem)
    
    newMail.Subject = subject
    newMail.Body = body
    newMail.Display
    
    Exit Sub
    
ErrorHandler:
    ShowError "新規メール作成中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 設定管理機能
' =============================================================================

' 設定管理
Public Sub ManageConfiguration()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    choice = InputBox("設定管理メニュー:" & vbCrLf & "1. API設定確認" & vbCrLf & "2. API設定変更ガイド" & vbCrLf & "3. バージョン情報" & vbCrLf & "番号を入力してください:", APP_NAME & " - 設定管理")
    
    Select Case choice
        Case "1": Call ShowConfigurationInfo
        Case "2": Call ShowAPISettingsGuide
        Case "3": Call ShowVersionInfo
        Case "": ' キャンセル時は何もしない
        Case Else: ShowMessage "無効な選択です。1-3の番号を入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "設定管理中にエラーが発生しました。", Err.Description
End Sub

' 設定情報表示
Private Sub ShowConfigurationInfo()
    Dim configInfo As String
    configInfo = "現在の設定:" & vbCrLf & vbCrLf & "アプリケーション名: " & APP_NAME & vbCrLf & "バージョン: " & APP_VERSION & vbCrLf & "API エンドポイント: " & Left(OPENAI_API_ENDPOINT, 50) & "..." & vbCrLf & "API キー: " & Left(OPENAI_API_KEY, 10) & "..." & vbCrLf & "最大処理文字数: " & MAX_CONTENT_LENGTH & vbCrLf & "タイムアウト: " & REQUEST_TIMEOUT & "秒"
    
    ShowMessage configInfo, "設定情報"
End Sub

' API設定変更ガイド
Private Sub ShowAPISettingsGuide()
    Dim guideMessage As String
    guideMessage = "API設定を変更するには、以下の手順を実行してください：" & vbCrLf & vbCrLf & "1. VBAエディタでこのファイルを開く" & vbCrLf & "2. ファイル上部の定数セクションを探す" & vbCrLf & "3. OPENAI_API_ENDPOINT と OPENAI_API_KEY を編集" & vbCrLf & "4. ファイルを保存" & vbCrLf & "5. 「API接続テスト」で動作確認"
    
    ShowMessage guideMessage, "API設定変更方法"
End Sub

' バージョン情報表示
Private Sub ShowVersionInfo()
    Dim versionInfo As String
    versionInfo = "Outlook AI Helper - 簡素化統合版" & vbCrLf & vbCrLf & "バージョン: " & APP_VERSION & vbCrLf & "作成日: 2024年" & vbCrLf & vbCrLf & "主要機能:" & vbCrLf & "・メール内容解析" & vbCrLf & "・営業断りメール作成" & vbCrLf & "・承諾メール作成" & vbCrLf & "・カスタムメール作成" & vbCrLf & "・設定管理" & vbCrLf & "・API接続テスト"
    
    ShowMessage versionInfo, "バージョン情報"
End Sub

' =============================================================================
' 日本語関数エイリアス（利便性向上のため）
' =============================================================================

' メインメニュー表示
Public Sub メインメニュー()
    Call ShowMainMenu
End Sub

' メール内容解析
Public Sub メール解析()
    Call AnalyzeSelectedEmail
End Sub

' 営業断りメール作成
Public Sub 営業断り()
    Call CreateRejectionEmail
End Sub

' 承諾メール作成
Public Sub 承諾メール()
    Call CreateAcceptanceEmail
End Sub

' カスタムメール作成
Public Sub カスタムメール()
    Call CreateCustomBusinessEmail
End Sub

' 設定管理
Public Sub 設定()
    Call ManageConfiguration
End Sub

' API接続テスト
Public Sub API接続テスト()
    Call TestAPIConnection
End Sub