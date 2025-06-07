' OutlookAI_Unified.bas
' Outlook OpenAI マクロ - 統合版
' 作成日: 2024
' 説明: 全機能を1つのファイルに統合した利用しやすい版
' 
' 利用方法: このファイルをOutlookのVBAプロジェクトにインポートするだけで全機能が利用可能
' 
' 主要機能:
' - メール内容解析
' - 営業断りメール作成
' - 承諾メール作成
' - カスタムメール作成
' - 検索フォルダ分析
' - 設定管理
' - API接続管理

Option Explicit

' =============================================================================
' 型定義（共通）
' =============================================================================

' メール内容を格納する型
Private Type EmailContent
    Subject As String
    Body As String
End Type

' 検索フォルダ情報を格納する型
Private Type SearchFolderInfo
    Name As String
    ItemCount As Long
    LastAccessDate As Date
    Description As String
End Type

' =============================================================================
' 定数定義
' =============================================================================

' OpenAI API 設定
Public Const OPENAI_API_ENDPOINT As String = "https://your-azure-openai-endpoint.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview"
Public Const OPENAI_API_KEY As String = "YOUR_API_KEY_HERE" ' 本番環境では設定ファイルから読み込み
Public Const OPENAI_MODEL As String = "gpt-4"

' アプリケーション設定
Public Const APP_NAME As String = "Outlook AI Helper"
Public Const APP_VERSION As String = "1.0.0 Unified"
Public Const MAX_CONTENT_LENGTH As Integer = 50000 ' 最大処理文字数
Public Const REQUEST_TIMEOUT As Integer = 30 ' APIリクエストタイムアウト（秒）

' プロンプトテンプレート
Public Const SYSTEM_PROMPT_ANALYZER As String = "あなたは日本語のビジネスメール分析の専門家です。メール内容を分析し、重要な情報を抽出してください。"
Public Const SYSTEM_PROMPT_COMPOSER As String = "あなたは日本語のビジネスメール作成の専門家です。適切で丁寧なビジネスメールを作成してください。"
Public Const SYSTEM_PROMPT_SEARCH As String = "あなたは日本語のメール整理・分析の専門家です。メールフォルダの分析と整理の提案をしてください。"

' エラーメッセージ
Public Const MSG_API_KEY_NOT_SET As String = "OpenAI API キーが設定されていません。OPENAI_API_KEY定数を設定してください。"
Public Const MSG_ENDPOINT_NOT_SET As String = "OpenAI API エンドポイントが設定されていません。OPENAI_API_ENDPOINT定数を設定してください。"

' =============================================================================
' メインメニュー関数（エントリーポイント）
' =============================================================================

' メインメニュー表示
Public Sub ShowMainMenu()
    Dim choice As String
    Dim menuText As String
    
    menuText = APP_NAME & " v" & APP_VERSION & vbCrLf & vbCrLf & _
               "利用可能な機能:" & vbCrLf & _
               "1. メール内容を解析" & vbCrLf & _
               "2. 営業断りメールを作成" & vbCrLf & _
               "3. 承諾メールを作成" & vbCrLf & _
               "4. カスタムメールを作成" & vbCrLf & _
               "5. 検索フォルダを分析" & vbCrLf & _
               "6. 設定管理" & vbCrLf & _
               "7. API接続テスト" & vbCrLf & vbCrLf & _
               "実行したい機能の番号を入力してください:"
    
    choice = InputBox(menuText, APP_NAME & " - メインメニュー")
    
    Select Case choice
        Case "1"
            Call AnalyzeSelectedEmail
        Case "2"
            Call CreateRejectionEmail
        Case "3"
            Call CreateAcceptanceEmail
        Case "4"
            Call CreateCustomBusinessEmail
        Case "5"
            Call AnalyzeSearchFolders
        Case "6"
            Call ManageConfiguration
        Case "7"
            Call TestAPIConnection
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。1-7の番号を入力してください。", "入力エラー", vbExclamation
    End Select
End Sub

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
                 "このファイルの OPENAI_API_KEY 定数を設定してください。"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ' エンドポイントの確認
    If InStr(OPENAI_API_ENDPOINT, "your-azure-openai-endpoint") > 0 Then
        ShowError "OpenAI API エンドポイントが設定されていません。", _
                 "このファイルの OPENAI_API_ENDPOINT 定数を設定してください。"
        ValidateConfiguration = False
        Exit Function
    End If
    
    ValidateConfiguration = True
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

' =============================================================================
' OpenAI API 接続関数
' =============================================================================

' OpenAI API にリクエストを送信
Public Function SendOpenAIRequest(ByVal systemPrompt As String, ByVal userMessage As String, Optional ByVal maxTokens As Integer = 1000) As String
    On Error GoTo ErrorHandler
    
    ' 設定の検証
    If Not ValidateConfiguration() Then
        SendOpenAIRequest = ""
        Exit Function
    End If
    
    ShowProgress "AI分析を開始しています..."
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' リクエストボディの作成
    Dim requestBody As String
    requestBody = CreateRequestBody(systemPrompt, userMessage, maxTokens)
    
    WriteLog "API リクエスト送信: " & Left(userMessage, 100) & "..."
    
    ' HTTPリクエストの設定
    http.Open "POST", OPENAI_API_ENDPOINT, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & OPENAI_API_KEY
    http.setRequestHeader "User-Agent", APP_NAME & "/" & APP_VERSION
    
    ShowProgress "APIに接続中..."
    
    ' リクエスト送信
    http.send requestBody
    
    ShowProgress "AI処理中..."
    
    ' レスポンスの処理
    If http.Status = 200 Then
        Dim response As String
        response = http.responseText
        WriteLog "API レスポンス受信成功"
        
        ' JSONレスポンスの解析
        Dim result As String
        result = ParseOpenAIResponse(response)
        
        WriteLog "AI分析完了: " & Len(result) & "文字"
        ShowProgress "AI分析完了"
        
        SendOpenAIRequest = result
    Else
        WriteLog "API エラー: " & http.Status & " - " & http.statusText, "ERROR"
        ShowError "OpenAI API接続エラーが発生しました。", _
                 "ステータスコード: " & http.Status & vbCrLf & _
                 "エラー内容: " & http.statusText & vbCrLf & vbCrLf & _
                 "設定を確認してください。"
        SendOpenAIRequest = ""
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog "API接続エラー: " & Err.Description, "ERROR"
    ShowError "API接続中にエラーが発生しました。", Err.Description
    SendOpenAIRequest = ""
End Function

' API接続テスト
Public Sub TestAPIConnection()
    On Error GoTo ErrorHandler
    
    If Not ValidateConfiguration() Then
        Exit Sub
    End If
    
    ShowProgress "API接続をテスト中..."
    
    Dim testPrompt As String
    testPrompt = "こんにちは"
    
    Dim result As String
    result = SendOpenAIRequest("簡潔に日本語で応答してください。", testPrompt, 50)
    
    If result <> "" Then
        ShowSuccess "API接続テストが成功しました！" & vbCrLf & vbCrLf & _
                   "応答: " & result
    Else
        ShowError "API接続テストが失敗しました。", "設定を確認してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "API接続テスト中にエラーが発生しました。", Err.Description
End Sub

' リクエストボディの作成
Private Function CreateRequestBody(ByVal systemPrompt As String, ByVal userMessage As String, ByVal maxTokens As Integer) As String
    Dim requestBody As String
    
    requestBody = "{" & _
                  """model"": """ & OPENAI_MODEL & """," & _
                  """messages"": [" & _
                  "{""role"": ""system"", ""content"": """ & EscapeJsonString(systemPrompt) & """}," & _
                  "{""role"": ""user"", ""content"": """ & EscapeJsonString(userMessage) & """}" & _
                  "]," & _
                  """max_tokens"": " & maxTokens & "," & _
                  """temperature"": 0.7" & _
                  "}"
    
    CreateRequestBody = requestBody
End Function

' OpenAI レスポンスの解析
Private Function ParseOpenAIResponse(ByVal jsonResponse As String) As String
    On Error GoTo ErrorHandler
    
    ' 簡易的なJSON解析（choices[0].message.contentを抽出）
    Dim contentStart As Integer
    Dim contentEnd As Integer
    Dim content As String
    
    ' "content": "..." の部分を抽出
    contentStart = InStr(jsonResponse, """content"": """)
    If contentStart > 0 Then
        contentStart = contentStart + 12 ' "content": " の長さ
        contentEnd = InStr(contentStart, jsonResponse, """, """)
        If contentEnd = 0 Then
            contentEnd = InStr(contentStart, jsonResponse, """}")
        End If
        If contentEnd = 0 Then
            contentEnd = InStr(contentStart, jsonResponse, """}]")
        End If
        
        If contentEnd > contentStart Then
            content = Mid(jsonResponse, contentStart, contentEnd - contentStart)
            ' エスケープ文字の復元
            content = Replace(content, "\""", """")
            content = Replace(content, "\\", "\")
            content = Replace(content, "\n", vbCrLf)
            content = Replace(content, "\t", vbTab)
            ParseOpenAIResponse = content
        Else
            ParseOpenAIResponse = "レスポンス解析エラー"
        End If
    Else
        ParseOpenAIResponse = "コンテンツが見つかりません"
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog "レスポンス解析エラー: " & Err.Description, "ERROR"
    ParseOpenAIResponse = "解析エラーが発生しました"
End Function

' =============================================================================
' メール解析機能
' =============================================================================

' 選択されたメールを解析（メニューから呼び出される）
Public Sub AnalyzeSelectedEmail()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Set mailItem = GetSelectedMailItem()
    
    If mailItem Is Nothing Then
        Exit Sub
    End If
    
    ' 解析タイプの選択
    Dim analysisType As String
    analysisType = InputBox("実行する解析を選択してください:" & vbCrLf & vbCrLf & _
                           "1. 基本情報分析" & vbCrLf & _
                           "2. 重要度・緊急度分析" & vbCrLf & _
                           "3. 感情・トーン分析" & vbCrLf & _
                           "4. 要約作成" & vbCrLf & _
                           "5. 内容明確化" & vbCrLf & vbCrLf & _
                           "番号を入力してください:", _
                           APP_NAME & " - 解析タイプ選択")
    
    If analysisType = "" Then Exit Sub
    
    Select Case analysisType
        Case "1"
            Call AnalyzeBasicInfo(mailItem)
        Case "2"
            Call AnalyzePriorityUrgency(mailItem)
        Case "3"
            Call AnalyzeEmotionTone(mailItem)
        Case "4"
            Call CreateSummary(mailItem)
        Case "5"
            Call ClarifyEmailContent(mailItem)
        Case Else
            ShowMessage "無効な選択です。1-5の番号を入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "メール解析中にエラーが発生しました。", Err.Description
End Sub

' 基本情報分析
Private Sub AnalyzeBasicInfo(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "基本情報を分析中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "以下のメールの基本情報を分析し、以下の形式で出力してください：" & vbCrLf & _
                   "1. 目的・内容要約" & vbCrLf & _
                   "2. 送信者の意図" & vbCrLf & _
                   "3. 求められているアクション" & vbCrLf & _
                   "4. 期限・締切" & vbCrLf & _
                   "5. 重要なキーワード"
    
    Dim userMessage As String
    userMessage = "以下のメールを分析してください：" & vbCrLf & vbCrLf & _
                  "【メール情報】" & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & " <" & mailItem.SenderEmailAddress & ">" & vbCrLf & _
                  "受信日時: " & mailItem.ReceivedTime & vbCrLf & vbCrLf & _
                  "【本文】" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, 1500)
    
    If result <> "" Then
        ShowAnalysisResult "基本情報分析結果", result
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "基本情報分析中にエラーが発生しました。", Err.Description
End Sub

' 分析結果表示（長いテキスト用）
Private Sub ShowAnalysisResult(ByVal title As String, ByVal content As String)
    Const maxLength As Integer = 1500
    
    If Len(content) <= maxLength Then
        ShowMessage content, title
    Else
        ' 長いコンテンツを分割して表示
        Dim currentPos As Integer
        Dim pageNum As Integer
        currentPos = 1
        pageNum = 1
        
        While currentPos <= Len(content)
            Dim pageContent As String
            Dim nextPos As Integer
            nextPos = currentPos + maxLength
            
            If nextPos > Len(content) Then
                pageContent = Mid(content, currentPos)
            Else
                ' 区切りの良い位置で分割
                Dim breakPos As Integer
                breakPos = InStrRev(content, vbCrLf, nextPos)
                If breakPos > currentPos Then
                    nextPos = breakPos
                End If
                pageContent = Mid(content, currentPos, nextPos - currentPos)
            End If
            
            Dim pageTitle As String
            pageTitle = title & " (" & pageNum & "/" & Int((Len(content) / maxLength) + 1) & ")"
            ShowMessage pageContent, pageTitle
            
            currentPos = nextPos + 1
            pageNum = pageNum + 1
        Wend
    End If
End Sub

' =============================================================================
' メール作成支援機能
' =============================================================================

' 営業メール断りメール作成（メニューから呼び出される）
Public Sub CreateRejectionEmail()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Set mailItem = GetSelectedMailItem()
    
    If mailItem Is Nothing Then
        Exit Sub
    End If
    
    ' 営業メールかどうかの確認
    If Not IsCommercialEmail(mailItem) Then
        If MsgBox("選択されたメールは営業メールのようには見えませんが、断りメールを作成しますか？", _
                  vbYesNo + vbQuestion, APP_NAME) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' 断り理由の入力
    Dim rejectionReason As String
    rejectionReason = InputBox("断り理由を入力してください（省略可）:" & vbCrLf & vbCrLf & _
                              "例:" & vbCrLf & _
                              "・予算の都合上" & vbCrLf & _
                              "・現在のサービスに満足している" & vbCrLf & _
                              "・組織の方針により" & vbCrLf & _
                              "・検討時期ではない", _
                              APP_NAME & " - 断り理由", "")
    
    ' 作成実行
    Call GenerateRejectionEmail(mailItem, rejectionReason)
    
    Exit Sub
    
ErrorHandler:
    ShowError "営業断りメール作成中にエラーが発生しました。", Err.Description
End Sub

' 承諾メール作成（メニューから呼び出される）
Public Sub CreateAcceptanceEmail()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Set mailItem = GetSelectedMailItem()
    
    If mailItem Is Nothing Then
        Exit Sub
    End If
    
    ' 承諾内容の詳細入力
    Dim acceptanceDetails As String
    acceptanceDetails = InputBox("承諾内容の詳細を入力してください（省略可）:" & vbCrLf & vbCrLf & _
                               "例:" & vbCrLf & _
                               "・喜んで参加させていただきます" & vbCrLf & _
                               "・条件を確認の上、前向きに検討いたします" & vbCrLf & _
                               "・詳細を教えていただければ対応可能です", _
                               APP_NAME & " - 承諾詳細", "")
    
    ' 作成実行
    Call GenerateAcceptanceEmail(mailItem, acceptanceDetails)
    
    Exit Sub
    
ErrorHandler:
    ShowError "承諾メール作成中にエラーが発生しました。", Err.Description
End Sub

' カスタムメール作成
Public Sub CreateCustomBusinessEmail()
    On Error GoTo ErrorHandler
    
    ' メールタイプの選択
    Dim emailType As String
    emailType = InputBox("作成するメールの種類を選択してください：" & vbCrLf & vbCrLf & _
                        "1. お礼メール" & vbCrLf & _
                        "2. 謝罪メール" & vbCrLf & _
                        "3. 問い合わせメール" & vbCrLf & _
                        "4. 提案・依頼メール" & vbCrLf & _
                        "5. フォローアップメール" & vbCrLf & _
                        "6. その他（カスタム）" & vbCrLf & vbCrLf & _
                        "番号を入力してください:", _
                        APP_NAME & " - メール作成")
    
    If emailType = "" Then Exit Sub
    
    ' 詳細情報の入力
    Dim emailDetails As String
    emailDetails = InputBox("メールの詳細情報を入力してください：" & vbCrLf & vbCrLf & _
                           "・宛先（相手の名前、役職等）" & vbCrLf & _
                           "・目的・内容" & vbCrLf & _
                           "・背景情報" & vbCrLf & _
                           "・期待する結果", _
                           APP_NAME & " - 詳細情報")
    
    If emailDetails = "" Then
        ShowMessage "詳細情報が入力されていません。", "入力エラー", vbExclamation
        Exit Sub
    End If
    
    ' メール作成実行
    Call GenerateCustomEmail(emailType, emailDetails)
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムメール作成中にエラーが発生しました。", Err.Description
End Sub

' 営業メール判定
Private Function IsCommercialEmail(ByVal mailItem As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim subject As String
    Dim body As String
    Dim combinedText As String
    
    subject = LCase(mailItem.Subject)
    body = LCase(GetEmailBodyText(mailItem))
    combinedText = subject & " " & body
    
    ' 営業メールのキーワード判定
    Dim commercialKeywords As Variant
    commercialKeywords = Array("営業", "セールス", "販売", "提案", "サービス", "商品", "料金", "価格", _
                              "無料", "キャンペーン", "お得", "割引", "特別", "限定", "資料請求", _
                              "デモ", "体験", "トライアル", "導入", "効果", "改善", "解決")
    
    Dim i As Integer
    For i = 0 To UBound(commercialKeywords)
        If InStr(combinedText, commercialKeywords(i)) > 0 Then
            IsCommercialEmail = True
            Exit Function
        End If
    Next i
    
    IsCommercialEmail = False
    Exit Function
    
ErrorHandler:
    IsCommercialEmail = False
End Function

' 生成されたメールの解析
Private Function ParseGeneratedEmail(ByVal generatedText As String) As EmailContent
    Dim result As EmailContent
    
    ' 件名の抽出
    Dim subjectStart As Integer
    Dim subjectEnd As Integer
    
    subjectStart = InStr(generatedText, "件名:")
    If subjectStart = 0 Then subjectStart = InStr(generatedText, "Subject:")
    If subjectStart = 0 Then subjectStart = InStr(generatedText, "【件名】")
    
    If subjectStart > 0 Then
        subjectStart = subjectStart + 3
        subjectEnd = InStr(subjectStart, generatedText, vbCrLf)
        If subjectEnd = 0 Then subjectEnd = Len(generatedText)
        
        result.Subject = Trim(Mid(generatedText, subjectStart, subjectEnd - subjectStart))
        ' 不要な文字の除去
        result.Subject = Replace(result.Subject, "【", "")
        result.Subject = Replace(result.Subject, "】", "")
        result.Subject = Replace(result.Subject, "「", "")
        result.Subject = Replace(result.Subject, "」", "")
    Else
        result.Subject = "Re: （AI生成メール）"
    End If
    
    ' 本文の抽出
    Dim bodyStart As Integer
    bodyStart = InStr(generatedText, "本文:")
    If bodyStart = 0 Then bodyStart = InStr(generatedText, "Body:")
    If bodyStart = 0 Then bodyStart = InStr(generatedText, "【本文】")
    If bodyStart = 0 Then bodyStart = subjectEnd
    
    If bodyStart > 0 Then
        result.Body = Trim(Mid(generatedText, bodyStart + 3))
        ' 不要な文字の除去
        result.Body = Replace(result.Body, "【本文】", "")
        result.Body = Replace(result.Body, "本文:", "")
    Else
        result.Body = generatedText
    End If
    
    ParseGeneratedEmail = result
End Function

' 営業メール断りメールを生成
Private Sub GenerateRejectionEmail(ByVal originalMail As Object, Optional ByVal rejectionReason As String = "")
    On Error GoTo ErrorHandler
    
    ShowProgress "営業断りメールを作成中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(originalMail)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 営業断りメール作成のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_COMPOSER & vbCrLf & _
                   "営業メールに対する丁寧で適切な断りメールを作成してください。" & vbCrLf & _
                   "以下の点に注意してください：" & vbCrLf & _
                   "1. 相手に敬意を示す丁寧な表現" & vbCrLf & _
                   "2. 断る理由を簡潔かつ適切に伝える" & vbCrLf & _
                   "3. 今後の関係を損なわない配慮" & vbCrLf & _
                   "4. 日本のビジネスマナーに適した文面" & vbCrLf & _
                   "5. 適切な長さ（長すぎず短すぎず）" & vbCrLf & vbCrLf & _
                   "出力形式：件名と本文を含む完全なメール"
    
    Dim userMessage As String
    userMessage = "以下の営業メールに対する断りメールを作成してください：" & vbCrLf & vbCrLf & _
                  "【元メール情報】" & vbCrLf & _
                  "件名: " & originalMail.Subject & vbCrLf & _
                  "送信者: " & originalMail.SenderName & vbCrLf & _
                  "送信者メールアドレス: " & originalMail.SenderEmailAddress & vbCrLf & vbCrLf & _
                  "【元メール本文】" & vbCrLf & emailBody & vbCrLf & vbCrLf
    
    If rejectionReason <> "" Then
        userMessage = userMessage & "【断り理由】" & vbCrLf & rejectionReason & vbCrLf & vbCrLf
    End If
    
    userMessage = userMessage & "適切で丁寧な断りメールを作成してください。"
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, 1500)
    
    If result <> "" Then
        ' 結果から件名と本文を分離
        Dim parsedEmail As EmailContent
        parsedEmail = ParseGeneratedEmail(result)
        
        ' 返信メールを作成
        CreateReplyEmail originalMail, parsedEmail.Subject, parsedEmail.Body
        
        ShowSuccess "営業断りメールの下書きを作成しました。内容を確認してから送信してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "営業断りメール生成中にエラーが発生しました。", Err.Description
End Sub

' 承諾メールを生成
Private Sub GenerateAcceptanceEmail(ByVal originalMail As Object, Optional ByVal acceptanceDetails As String = "")
    On Error GoTo ErrorHandler
    
    ShowProgress "承諾メールを作成中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(originalMail)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 承諾メール作成のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_COMPOSER & vbCrLf & _
                   "依頼や提案に対する前向きで適切な承諾メールを作成してください。" & vbCrLf & _
                   "以下の点に注意してください：" & vbCrLf & _
                   "1. 感謝の気持ちを示す" & vbCrLf & _
                   "2. 承諾の意思を明確に伝える" & vbCrLf & _
                   "3. 必要に応じて条件や確認事項を含める" & vbCrLf & _
                   "4. 次のステップを提案する" & vbCrLf & _
                   "5. 日本のビジネスマナーに適した文面" & vbCrLf & _
                   "6. 建設的で前向きなトーン" & vbCrLf & vbCrLf & _
                   "出力形式：件名と本文を含む完全なメール"
    
    Dim userMessage As String
    userMessage = "以下のメールに対する承諾・同意の返信メールを作成してください：" & vbCrLf & vbCrLf & _
                  "【元メール情報】" & vbCrLf & _
                  "件名: " & originalMail.Subject & vbCrLf & _
                  "送信者: " & originalMail.SenderName & vbCrLf & vbCrLf & _
                  "【元メール本文】" & vbCrLf & emailBody & vbCrLf & vbCrLf
    
    If acceptanceDetails <> "" Then
        userMessage = userMessage & "【承諾詳細・追加情報】" & vbCrLf & acceptanceDetails & vbCrLf & vbCrLf
    End If
    
    userMessage = userMessage & "適切で前向きな承諾メールを作成してください。"
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, 1500)
    
    If result <> "" Then
        ' 結果から件名と本文を分離
        Dim parsedEmail As EmailContent
        parsedEmail = ParseGeneratedEmail(result)
        
        ' 返信メールを作成
        CreateReplyEmail originalMail, parsedEmail.Subject, parsedEmail.Body
        
        ShowSuccess "承諾メールの下書きを作成しました。内容を確認してから送信してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "承諾メール生成中にエラーが発生しました。", Err.Description
End Sub

' カスタムメール生成
Private Sub GenerateCustomEmail(ByVal emailType As String, ByVal details As String)
    On Error GoTo ErrorHandler
    
    ShowProgress "カスタムメールを作成中..."
    
    ' メールタイプに応じたプロンプト設定
    Dim emailTypeText As String
    Select Case emailType
        Case "1": emailTypeText = "お礼メール"
        Case "2": emailTypeText = "謝罪メール"
        Case "3": emailTypeText = "問い合わせメール"
        Case "4": emailTypeText = "提案・依頼メール"
        Case "5": emailTypeText = "フォローアップメール"
        Case Else: emailTypeText = "ビジネスメール"
    End Select
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_COMPOSER & vbCrLf & _
                   "以下の種類のビジネスメールを作成してください：" & emailTypeText & vbCrLf & vbCrLf & _
                   "注意点：" & vbCrLf & _
                   "1. 適切なビジネス敬語を使用" & vbCrLf & _
                   "2. 相手に配慮した丁寧な表現" & vbCrLf & _
                   "3. 明確で簡潔な内容" & vbCrLf & _
                   "4. 適切な件名を設定" & vbCrLf & _
                   "5. 必要に応じて次のアクションを提案" & vbCrLf & vbCrLf & _
                   "出力形式：件名と本文を含む完全なメール"
    
    Dim userMessage As String
    userMessage = "以下の情報を基に" & emailTypeText & "を作成してください：" & vbCrLf & vbCrLf & _
                  details
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, 1500)
    
    If result <> "" Then
        ' 結果から件名と本文を分離
        Dim parsedEmail As EmailContent
        parsedEmail = ParseGeneratedEmail(result)
        
        ' 新規メールを作成
        CreateNewEmail parsedEmail.Subject, parsedEmail.Body
        
        ShowSuccess "カスタムメールの下書きを作成しました。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムメール生成中にエラーが発生しました。", Err.Description
End Sub

' 返信メールの作成
Private Sub CreateReplyEmail(ByVal originalMail As Object, ByVal subject As String, ByVal body As String)
    On Error GoTo ErrorHandler
    
    Dim replyMail As Object
    Set replyMail = originalMail.Reply
    
    ' 件名の設定（Reを除去して新しい件名を設定）
    If Left(subject, 3) <> "Re:" And Left(subject, 3) <> "RE:" Then
        replyMail.Subject = "Re: " & subject
    Else
        replyMail.Subject = subject
    End If
    
    replyMail.Body = body
    replyMail.Save
    
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
    newMail.Save
    
    Exit Sub
    
ErrorHandler:
    ShowError "新規メール作成中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 検索フォルダ分析機能
' =============================================================================

' 検索フォルダ分析（メニューから呼び出される）
Public Sub AnalyzeSearchFolders()
    On Error GoTo ErrorHandler
    
    ' 分析タイプの選択
    Dim analysisType As String
    analysisType = InputBox("実行する分析を選択してください:" & vbCrLf & vbCrLf & _
                           "1. 検索フォルダ使用状況分析" & vbCrLf & _
                           "2. 不要フォルダ検出" & vbCrLf & _
                           "3. フォルダ整理提案" & vbCrLf & _
                           "4. 検索条件分析" & vbCrLf & vbCrLf & _
                           "番号を入力してください:", _
                           APP_NAME & " - 分析タイプ選択")
    
    If analysisType = "" Then Exit Sub
    
    Select Case analysisType
        Case "1"
            Call AnalyzeSearchFolderUsage
        Case "2"
            Call DetectUnusedSearchFolders
        Case "3"
            Call SuggestFolderOrganization
        Case "4"
            Call AnalyzeSearchConditions
        Case Else
            ShowMessage "無効な選択です。1-4の番号を入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "検索フォルダ分析中にエラーが発生しました。", Err.Description
End Sub

' 検索フォルダ使用状況分析
Private Sub AnalyzeSearchFolderUsage()
    On Error GoTo ErrorHandler
    
    ShowProgress "検索フォルダ使用状況を分析中..."
    
    Dim searchFolders As Object
    Dim folderInfo As String
    Dim i As Integer
    
    Set searchFolders = Application.Session.DefaultStore.GetSearchFolders
    
    folderInfo = "【検索フォルダ使用状況分析】" & vbCrLf & vbCrLf
    folderInfo = folderInfo & "総検索フォルダ数: " & searchFolders.Count & vbCrLf & vbCrLf
    
    For i = 1 To searchFolders.Count
        Dim folder As Object
        Set folder = searchFolders.Item(i)
        
        folderInfo = folderInfo & "■ " & folder.Name & vbCrLf
        folderInfo = folderInfo & "  アイテム数: " & folder.Items.Count & vbCrLf
        folderInfo = folderInfo & "  作成日: " & folder.CreationTime & vbCrLf
        
        ' フォルダの検索条件を取得（可能な場合）
        If folder.Items.Count > 0 Then
            folderInfo = folderInfo & "  状態: アクティブ" & vbCrLf
        Else
            folderInfo = folderInfo & "  状態: 空" & vbCrLf
        End If
        
        folderInfo = folderInfo & vbCrLf
    Next i
    
    ' AI分析の実行
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_SEARCH & vbCrLf & _
                   "検索フォルダの使用状況を分析し、以下の観点で評価・提案してください：" & vbCrLf & _
                   "1. 使用頻度の評価" & vbCrLf & _
                   "2. 効率性の評価" & vbCrLf & _
                   "3. 改善提案" & vbCrLf & _
                   "4. 統合・削除候補"
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, folderInfo, 2000)
    
    If result <> "" Then
        ShowAnalysisResult "検索フォルダ使用状況分析", folderInfo & vbCrLf & "【AI分析結果】" & vbCrLf & result
    Else
        ShowAnalysisResult "検索フォルダ使用状況分析", folderInfo
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "検索フォルダ使用状況分析中にエラーが発生しました。", Err.Description
End Sub

' 不要フォルダ検出
Private Sub DetectUnusedSearchFolders()
    On Error GoTo ErrorHandler
    
    ShowProgress "不要な検索フォルダを検出中..."
    
    Dim searchFolders As Object
    Dim unusedFolders As String
    Dim i As Integer
    Dim unusedCount As Integer
    
    Set searchFolders = Application.Session.DefaultStore.GetSearchFolders
    unusedCount = 0
    unusedFolders = "【不要検索フォルダ検出結果】" & vbCrLf & vbCrLf
    
    For i = 1 To searchFolders.Count
        Dim folder As Object
        Set folder = searchFolders.Item(i)
        
        ' 空のフォルダまたは長期間未使用のフォルダを検出
        If folder.Items.Count = 0 Then
            unusedCount = unusedCount + 1
            unusedFolders = unusedFolders & "■ " & folder.Name & " (空フォルダ)" & vbCrLf
            unusedFolders = unusedFolders & "  作成日: " & folder.CreationTime & vbCrLf
            unusedFolders = unusedFolders & "  推奨: 削除候補" & vbCrLf & vbCrLf
        ElseIf DateDiff("d", folder.CreationTime, Now) > 90 And folder.Items.Count < 5 Then
            unusedCount = unusedCount + 1
            unusedFolders = unusedFolders & "■ " & folder.Name & " (使用頻度低)" & vbCrLf
            unusedFolders = unusedFolders & "  作成日: " & folder.CreationTime & vbCrLf
            unusedFolders = unusedFolders & "  アイテム数: " & folder.Items.Count & vbCrLf
            unusedFolders = unusedFolders & "  推奨: 統合または削除を検討" & vbCrLf & vbCrLf
        End If
    Next i
    
    If unusedCount = 0 Then
        unusedFolders = unusedFolders & "不要な検索フォルダは検出されませんでした。" & vbCrLf & _
                                       "すべての検索フォルダが有効に使用されています。"
    Else
        unusedFolders = unusedFolders & "検出された不要フォルダ数: " & unusedCount & vbCrLf & vbCrLf & _
                                       "※削除前に内容を確認してください。"
    End If
    
    ShowAnalysisResult "不要検索フォルダ検出", unusedFolders
    
    Exit Sub
    
ErrorHandler:
    ShowError "不要フォルダ検出中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 設定管理機能
' =============================================================================

' 設定管理メニュー
Public Sub ManageConfiguration()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    choice = InputBox("設定管理メニュー:" & vbCrLf & vbCrLf & _
                     "1. API設定確認" & vbCrLf & _
                     "2. API設定変更" & vbCrLf & _
                     "3. 設定情報表示" & vbCrLf & _
                     "4. 初期化" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     APP_NAME & " - 設定管理")
    
    Select Case choice
        Case "1"
            Call ShowConfigurationInfo
        Case "2"
            Call ChangeAPISettings
        Case "3"
            Call ShowDetailedConfiguration
        Case "4"
            Call ResetConfiguration
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。1-4の番号を入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "設定管理中にエラーが発生しました。", Err.Description
End Sub

' 設定情報表示
Private Sub ShowConfigurationInfo()
    Dim configInfo As String
    configInfo = "現在の設定:" & vbCrLf & vbCrLf & _
                "アプリケーション名: " & APP_NAME & vbCrLf & _
                "バージョン: " & APP_VERSION & vbCrLf & _
                "API エンドポイント: " & Left(OPENAI_API_ENDPOINT, 50) & "..." & vbCrLf & _
                "API キー: " & Left(OPENAI_API_KEY, 10) & "..." & vbCrLf & _
                "最大処理文字数: " & MAX_CONTENT_LENGTH & vbCrLf & _
                "タイムアウト: " & REQUEST_TIMEOUT & "秒" & vbCrLf & vbCrLf & _
                "※設定を変更するには、このファイルの定数を編集してください。"
    
    ShowMessage configInfo, "設定情報"
End Sub

' API設定変更ガイド
Private Sub ChangeAPISettings()
    Dim guideMessage As String
    guideMessage = "API設定を変更するには、以下の手順を実行してください：" & vbCrLf & vbCrLf & _
                   "1. VBAエディタでこのファイルを開く" & vbCrLf & _
                   "2. ファイル上部の定数セクションを探す" & vbCrLf & _
                   "3. 以下の定数を編集:" & vbCrLf & _
                   "   - OPENAI_API_ENDPOINT" & vbCrLf & _
                   "   - OPENAI_API_KEY" & vbCrLf & _
                   "4. ファイルを保存" & vbCrLf & _
                   "5. 「API接続テスト」で動作確認" & vbCrLf & vbCrLf & _
                   "※設定後は必ずテストを実行してください。"
    
    ShowMessage guideMessage, "API設定変更方法"
End Sub

' 詳細設定情報表示
Private Sub ShowDetailedConfiguration()
    Dim detailedInfo As String
    detailedInfo = "詳細設定情報:" & vbCrLf & vbCrLf & _
                   "【アプリケーション情報】" & vbCrLf & _
                   "名前: " & APP_NAME & vbCrLf & _
                   "バージョン: " & APP_VERSION & vbCrLf & _
                   "モード: 統合版" & vbCrLf & vbCrLf & _
                   "【API設定】" & vbCrLf & _
                   "エンドポイント: " & OPENAI_API_ENDPOINT & vbCrLf & _
                   "モデル: " & OPENAI_MODEL & vbCrLf & _
                   "最大トークン: 1000-2000（機能により可変）" & vbCrLf & _
                   "タイムアウト: " & REQUEST_TIMEOUT & "秒" & vbCrLf & vbCrLf & _
                   "【制限事項】" & vbCrLf & _
                   "最大文字数: " & MAX_CONTENT_LENGTH & "文字" & vbCrLf & _
                   "同時処理: 1リクエスト" & vbCrLf & vbCrLf & _
                   "【利用可能機能】" & vbCrLf & _
                   "・メール解析" & vbCrLf & _
                   "・メール作成支援" & vbCrLf & _
                   "・検索フォルダ分析" & vbCrLf & _
                   "・API接続テスト"
    
    ShowAnalysisResult "詳細設定情報", detailedInfo
End Sub

' 設定初期化
Private Sub ResetConfiguration()
    If MsgBox("設定を初期状態に戻しますか？" & vbCrLf & vbCrLf & _
              "※この操作はファイルの定数を手動で編集する必要があります。", _
              vbYesNo + vbQuestion, "設定初期化") = vbYes Then
        
        Dim resetGuide As String
        resetGuide = "設定を初期化するには：" & vbCrLf & vbCrLf & _
                     "1. VBAエディタでこのファイルを開く" & vbCrLf & _
                     "2. 以下の定数を初期値に戻す:" & vbCrLf & _
                     "   OPENAI_API_ENDPOINT = ""https://your-azure-openai-endpoint...""" & vbCrLf & _
                     "   OPENAI_API_KEY = ""YOUR_API_KEY_HERE""" & vbCrLf & _
                     "3. ファイルを保存" & vbCrLf & vbCrLf & _
                     "初期化後は再度API設定を行ってください。"
        
        ShowMessage resetGuide, "設定初期化手順"
    End If
End Sub