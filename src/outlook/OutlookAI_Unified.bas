' OutlookAI_Unified.bas
' Outlook OpenAI マクロ - 統合版
' 作成日: 2024
' 説明: 全機能を1つのファイルに統合した利用しやすい版
' 
' 利用方法: このファイルをOutlookのVBAプロジェクトにインポートするだけで全機能が利用可能
' 
' 注意点:
' - InputBoxの表示文字数は254文字までのため、表示テキストは簡潔にする
' - VBAでは比較演算子は = を使用（== ではなく）
' - 文字列結合の行継続には _ を使用する
' - APIレスポンスのJSON解析は簡易処理のため、形式が変わると修正が必要
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
Public Const MAX_CONTENT_LENGTH As Long = 50000 ' 最大処理文字数
Public Const REQUEST_TIMEOUT As Integer = 30 ' APIリクエストタイムアウト（秒）
Public Const MAX_TOKEN As Integer = 15000 ' 最大トークン数

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
    
    ' "content": "..." の部分を抽出（新しいAPIレスポンス形式に対応）
    contentStart = InStr(jsonResponse, """content"": """)
    If contentStart > 0 Then
        ' 従来の形式
        contentStart = contentStart + 12 ' "content": " の長さ
        contentEnd = InStr(contentStart, jsonResponse, """, """)
        If contentEnd = 0 Then
            contentEnd = InStr(contentStart, jsonResponse, """}")
        End If
        If contentEnd = 0 Then
            contentEnd = InStr(contentStart, jsonResponse, """}]")
        End If
    Else
        ' 新しい形式：message.content を検索
        contentStart = InStr(jsonResponse, """message"":{")
        If contentStart > 0 Then
            contentStart = InStr(contentStart, jsonResponse, """content"":""")
            If contentStart > 0 Then
                contentStart = contentStart + 11 ' "content":"" の長さ
                contentEnd = InStr(contentStart, jsonResponse, """,""")
                If contentEnd = 0 Then
                    contentEnd = InStr(contentStart, jsonResponse, """}")
                End If
            End If
        End If
    End If
    
    If contentStart > 0 And contentEnd > contentStart Then
        content = Mid(jsonResponse, contentStart, contentEnd - contentStart)
        ' エスケープ文字の復元
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
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ShowAnalysisResult "基本情報分析結果", result
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "基本情報分析中にエラーが発生しました。", Err.Description
End Sub

' 重要度・緊急度分析
Private Sub AnalyzePriorityUrgency(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "重要度・緊急度を分析中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "以下のメールの重要度と緊急度を分析し、以下の形式で出力してください：" & vbCrLf & _
                   "1. 重要度評価（5段階：非常に重要、重要、中程度、低い、非常に低い）" & vbCrLf & _
                   "2. 緊急度評価（5段階：緊急、早急に対応、通常対応、余裕あり、緊急性なし）" & vbCrLf & _
                   "3. 重要度・緊急度の根拠" & vbCrLf & _
                   "4. 推奨される対応時間（即時、当日中、数日以内、1週間以内、特になし）" & vbCrLf & _
                   "5. 注意すべき点や優先事項"
    
    Dim userMessage As String
    userMessage = "以下のメールを重要度・緊急度の観点で分析してください：" & vbCrLf & vbCrLf & _
                  "【メール情報】" & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & " <" & mailItem.SenderEmailAddress & ">" & vbCrLf & _
                  "受信日時: " & mailItem.ReceivedTime & vbCrLf & vbCrLf & _
                  "【本文】" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ShowAnalysisResult "重要度・緊急度分析結果", result
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "重要度・緊急度分析中にエラーが発生しました。", Err.Description
End Sub

' 感情・トーン分析
Private Sub AnalyzeEmotionTone(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "感情とトーンを分析中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "以下のメールの感情とトーンを分析し、以下の形式で出力してください：" & vbCrLf & _
                   "1. 全体的な感情傾向（ポジティブ、ニュートラル、ネガティブなど）" & vbCrLf & _
                   "2. 検出された具体的な感情（喜び、怒り、期待、不満、感謝など）" & vbCrLf & _
                   "3. 文章のトーン（フォーマル、カジュアル、丁寧、強制的など）" & vbCrLf & _
                   "4. 文化的・ビジネス的背景の考慮事項" & vbCrLf & _
                   "5. 送信者の意図と期待する返信の推測" & vbCrLf & _
                   "6. 返信時に注意すべき感情やトーンの推奨"
    
    Dim userMessage As String
    userMessage = "以下のメールの感情とトーンを分析してください：" & vbCrLf & vbCrLf & _
                  "【メール情報】" & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & " <" & mailItem.SenderEmailAddress & ">" & vbCrLf & _
                  "受信日時: " & mailItem.ReceivedTime & vbCrLf & vbCrLf & _
                  "【本文】" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ShowAnalysisResult "感情・トーン分析結果", result
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "感情・トーン分析中にエラーが発生しました。", Err.Description
End Sub

' 要約作成
Private Sub CreateSummary(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "メールの要約を作成中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 要約の長さ選択
    Dim summaryLengthChoice As String
    summaryLengthChoice = InputBox("要約の長さを選択してください:" & vbCrLf & vbCrLf & _
                                  "1. 短い（1-2文）" & vbCrLf & _
                                  "2. 中程度（3-5文）" & vbCrLf & _
                                  "3. 詳細（段落形式）" & vbCrLf & vbCrLf & _
                                  "番号を入力してください:", _
                                  APP_NAME & " - 要約作成", "2")
    
    Dim summaryLength As String
    Select Case summaryLengthChoice
        Case "1"
            summaryLength = "短い（1-2文）"
        Case "3"
            summaryLength = "詳細（段落形式）"
        Case Else
            summaryLength = "中程度（3-5文）"
    End Select
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "以下のメールを" & summaryLength & "の長さで要約してください。" & vbCrLf & _
                   "要約には以下を含めてください：" & vbCrLf & _
                   "1. メインテーマ" & vbCrLf & _
                   "2. 重要なポイント" & vbCrLf & _
                   "3. 送信者の意図" & vbCrLf & _
                   "4. 必要なアクション（もしあれば）" & vbCrLf & _
                   "5. 締切や日時（もしあれば）"
    
    Dim userMessage As String
    userMessage = "以下のメールを" & summaryLength & "で要約してください：" & vbCrLf & vbCrLf & _
                  "【メール情報】" & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & " <" & mailItem.SenderEmailAddress & ">" & vbCrLf & _
                  "受信日時: " & mailItem.ReceivedTime & vbCrLf & vbCrLf & _
                  "【本文】" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ShowAnalysisResult "メール要約", result
        
        ' 要約のクリップボードへのコピーオプション
        If MsgBox("この要約をクリップボードにコピーしますか？", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            CopyToClipboard result
            ShowMessage "要約をクリップボードにコピーしました。", "コピー完了"
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "メール要約作成中にエラーが発生しました。", Err.Description
End Sub

' クリップボードにテキストをコピー
Private Sub CopyToClipboard(ByVal text As String)
    On Error GoTo ErrorHandler
    
    Dim dataObj As Object
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText text
    dataObj.PutInClipboard
    
    Exit Sub
    
ErrorHandler:
    WriteLog "クリップボードへのコピーエラー: " & Err.Description, "ERROR"
End Sub

' 分析結果表示（長いテキスト用）
Private Sub ShowAnalysisResult(ByVal title As String, ByVal content As String)
    Const maxLength As Integer = MAX_TOKEN
    
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

' 内容明確化
Private Sub ClarifyEmailContent(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "メール内容を明確化中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 明確化タイプの選択
    Dim clarificationType As String
    clarificationType = InputBox("実行する明確化タイプを選択してください:" & vbCrLf & vbCrLf & _
                               "1. あいまいな表現の明確化" & vbCrLf & _
                               "2. 難解な専門用語の説明" & vbCrLf & _
                               "3. 隠れた意図の分析" & vbCrLf & _
                               "4. 文化的・言語的背景の解説" & vbCrLf & _
                               "5. 論点の整理" & vbCrLf & vbCrLf & _
                               "番号を入力してください:", _
                               APP_NAME & " - 明確化タイプ", "1")
    
    Dim clarificationTypeText As String
    Select Case clarificationType
        Case "2"
            clarificationTypeText = "難解な専門用語の説明"
        Case "3"
            clarificationTypeText = "隠れた意図の分析"
        Case "4"
            clarificationTypeText = "文化的・言語的背景の解説"
        Case "5"
            clarificationTypeText = "論点の整理"
        Case Else
            clarificationTypeText = "あいまいな表現の明確化"
    End Select
    
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "以下のメール内容の「" & clarificationTypeText & "」を行い、より明確にしてください。" & vbCrLf & _
                   "特に以下の点に注目して分析してください：" & vbCrLf
    
    ' 明確化タイプに応じたプロンプト追加
    Select Case clarificationType
        Case "1" ' あいまいな表現の明確化
            systemPrompt = systemPrompt & _
                   "1. あいまいな表現や婉曲表現の特定" & vbCrLf & _
                   "2. それらの表現の可能性のある意味の解釈" & vbCrLf & _
                   "3. 明確化のための言い換え提案" & vbCrLf & _
                   "4. 明確化を求めるべき質問例"
        Case "2" ' 難解な専門用語の説明
            systemPrompt = systemPrompt & _
                   "1. メール内の専門用語や特殊な表現の特定" & vbCrLf & _
                   "2. それらの用語の平易な説明" & vbCrLf & _
                   "3. 用語の背景知識や文脈" & vbCrLf & _
                   "4. 関連する基本概念の解説"
        Case "3" ' 隠れた意図の分析
            systemPrompt = systemPrompt & _
                   "1. 明示されていない送信者の意図や期待の推測" & vbCrLf & _
                   "2. 言外の要求や期待の特定" & vbCrLf & _
                   "3. 送信者の立場や状況の考慮" & vbCrLf & _
                   "4. 適切な対応方法の提案"
        Case "4" ' 文化的・言語的背景の解説
            systemPrompt = systemPrompt & _
                   "1. 文化的な表現や慣用句の特定" & vbCrLf & _
                   "2. ビジネスコミュニケーションの文化的背景" & vbCrLf & _
                   "3. フォーマル/インフォーマル表現の解釈" & vbCrLf & _
                   "4. 言葉の選択に表れる敬意や関係性の考察"
        Case "5" ' 論点の整理
            systemPrompt = systemPrompt & _
                   "1. メール内の主要な論点の特定と整理" & vbCrLf & _
                   "2. 論理展開の構造化" & vbCrLf & _
                   "3. 複雑な内容の要素分解" & vbCrLf & _
                   "4. 理解しやすい順序での再構成"
    End Select
    
    Dim userMessage As String
    userMessage = "以下のメールの「" & clarificationTypeText & "」を行ってください：" & vbCrLf & vbCrLf & _
                  "【メール情報】" & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & " <" & mailItem.SenderEmailAddress & ">" & vbCrLf & _
                  "受信日時: " & mailItem.ReceivedTime & vbCrLf & vbCrLf & _
                  "【本文】" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ShowAnalysisResult clarificationTypeText & " 結果", result
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "内容明確化中にエラーが発生しました。", Err.Description
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
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ' 結果から件名と本文を分離
        Dim parsedEmail As EmailContent
        parsedEmail = ParseGeneratedEmail(result)
        
        ' 返信メールを作成
        CreateReplyEmail originalMail, parsedEmail.Subject, parsedEmail.Body
        
        ShowSuccess "営業断りメールの返信ウィンドウを開きました。内容を確認してから送信してください。"
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
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ' 結果から件名と本文を分離
        Dim parsedEmail As EmailContent
        parsedEmail = ParseGeneratedEmail(result)
        
        ' 返信メールを作成
        CreateReplyEmail originalMail, parsedEmail.Subject, parsedEmail.Body
        
        ShowSuccess "承諾メールの返信ウィンドウを開きました。内容を確認してから送信してください。"
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
    result = SendOpenAIRequest(systemPrompt, userMessage, MAX_TOKEN)
    
    If result <> "" Then
        ' 結果から件名と本文を分離
        Dim parsedEmail As EmailContent
        parsedEmail = ParseGeneratedEmail(result)
        
        ' 新規メールを作成
        CreateNewEmail parsedEmail.Subject, parsedEmail.Body
        
        ShowSuccess "カスタムメールの編集ウィンドウを開きました。"
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

' フォルダ整理提案
Private Sub SuggestFolderOrganization()
    On Error GoTo ErrorHandler
    
    ShowProgress "フォルダ整理案を作成中..."
    
    Dim searchFolders As Object
    Dim folderInfo As String
    Dim i As Integer
    
    Set searchFolders = Application.Session.DefaultStore.GetSearchFolders
    
    folderInfo = "【現在の検索フォルダ一覧】" & vbCrLf & vbCrLf
    folderInfo = folderInfo & "総検索フォルダ数: " & searchFolders.Count & vbCrLf & vbCrLf
    
    ' フォルダ一覧の取得
    For i = 1 To searchFolders.Count
        Dim folder As Object
        Set folder = searchFolders.Item(i)
        
        folderInfo = folderInfo & i & ". " & folder.Name & vbCrLf
        folderInfo = folderInfo & "   アイテム数: " & folder.Items.Count & vbCrLf
        folderInfo = folderInfo & "   作成日: " & folder.CreationTime & vbCrLf
        
        ' アクティブ状態の確認
        If folder.Items.Count > 0 Then
            folderInfo = folderInfo & "   状態: アクティブ" & vbCrLf
        Else
            folderInfo = folderInfo & "   状態: 空" & vbCrLf
        End If
        
        ' 検索フォルダ名からカテゴリや目的を推測
        Dim folderNameLower As String
        folderNameLower = LCase(folder.Name)
        
        If InStr(folderNameLower, "プロジェクト") > 0 Or InStr(folderNameLower, "案件") > 0 Then
            folderInfo = folderInfo & "   カテゴリ: プロジェクト関連" & vbCrLf
        ElseIf InStr(folderNameLower, "会議") > 0 Or InStr(folderNameLower, "ミーティング") > 0 Then
            folderInfo = folderInfo & "   カテゴリ: 会議関連" & vbCrLf
        ElseIf InStr(folderNameLower, "重要") > 0 Or InStr(folderNameLower, "緊急") > 0 Then
            folderInfo = folderInfo & "   カテゴリ: 重要/緊急" & vbCrLf
        ElseIf InStr(folderNameLower, "顧客") > 0 Or InStr(folderNameLower, "クライアント") > 0 Then
            folderInfo = folderInfo & "   カテゴリ: 顧客関連" & vbCrLf
        ElseIf InStr(folderNameLower, "社内") > 0 Or InStr(folderNameLower, "部門") > 0 Then
            folderInfo = folderInfo & "   カテゴリ: 社内連絡" & vbCrLf
        Else
            folderInfo = folderInfo & "   カテゴリ: その他/分類不明" & vbCrLf
        End If
        
        folderInfo = folderInfo & vbCrLf
    Next i
    
    ' AI分析の実行
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_SEARCH & vbCrLf & _
                   "現在の検索フォルダ構成を分析し、効率的な整理方法を提案してください。" & vbCrLf & _
                   "以下の点を考慮して具体的な提案をしてください：" & vbCrLf & _
                   "1. フォルダの名称統一化（命名規則の提案）" & vbCrLf & _
                   "2. カテゴリ別の整理方法" & vbCrLf & _
                   "3. 重複フォルダの統合" & vbCrLf & _
                   "4. フォルダ階層の整理" & vbCrLf & _
                   "5. 検索フォルダの作成基準" & vbCrLf & _
                   "6. アーカイブ/定期メンテナンス方法"
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, folderInfo, 2500)
    
    If result <> "" Then
        ShowAnalysisResult "フォルダ整理提案", folderInfo & vbCrLf & vbCrLf & _
                           "【AI提案内容】" & vbCrLf & vbCrLf & result
    Else
        ShowAnalysisResult "フォルダ整理提案", folderInfo & vbCrLf & vbCrLf & _
                           "AI分析を実行できませんでした。API設定を確認してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "フォルダ整理提案中にエラーが発生しました。", Err.Description
End Sub

' 検索条件分析
Private Sub AnalyzeSearchConditions()
    On Error GoTo ErrorHandler
    
    ShowProgress "検索条件を分析中..."
    
    Dim searchFolders As Object
    Set searchFolders = Application.Session.DefaultStore.GetSearchFolders
    
    If searchFolders.Count = 0 Then
        ShowMessage "検索フォルダが存在しません。検索フォルダを作成してから実行してください。", "フォルダなし", vbExclamation
        Exit Sub
    End If
    
    ' 分析対象フォルダの選択
    Dim folderList As String
    Dim i As Integer
    
    folderList = "分析する検索フォルダを選択してください:" & vbCrLf & vbCrLf
    
    For i = 1 To searchFolders.Count
        folderList = folderList & i & ". " & searchFolders.Item(i).Name & vbCrLf
    Next i
    
    folderList = folderList & vbCrLf & "番号を入力してください:"
    
    Dim folderChoice As String
    folderChoice = InputBox(folderList, APP_NAME & " - フォルダ選択")
    
    If folderChoice = "" Then Exit Sub
    
    ' 選択された番号の検証
    Dim folderIndex As Integer
    On Error Resume Next
    folderIndex = CInt(folderChoice)
    On Error GoTo ErrorHandler
    
    If folderIndex < 1 Or folderIndex > searchFolders.Count Then
        ShowMessage "無効な選択です。1-" & searchFolders.Count & "の番号を入力してください。", "入力エラー", vbExclamation
        Exit Sub
    End If
    
    ' 選択されたフォルダの情報取得
    Dim selectedFolder As Object
    Set selectedFolder = searchFolders.Item(folderIndex)
    
    ShowProgress "フォルダ「" & selectedFolder.Name & "」の検索条件を分析中..."
    
    ' フォルダの詳細情報収集
    Dim folderInfo As String
    folderInfo = "【検索フォルダ情報】" & vbCrLf & vbCrLf & _
                "フォルダ名: " & selectedFolder.Name & vbCrLf & _
                "作成日時: " & selectedFolder.CreationTime & vbCrLf & _
                "アイテム数: " & selectedFolder.Items.Count & vbCrLf
                
    ' フォルダに含まれるメールのサンプルを収集
    Dim sampleMail As String
    sampleMail = ""
    
    If selectedFolder.Items.Count > 0 Then
        Dim mailCount As Integer
        mailCount = 0
        
        ' 最大5件のメールをサンプルとして収集
        For i = 1 To selectedFolder.Items.Count
            If mailCount >= 5 Then Exit For
            
            On Error Resume Next
            Dim item As Object
            Set item = selectedFolder.Items.Item(i)
            
            If Not item Is Nothing Then
                If item.Class = olMail Then
                    mailCount = mailCount + 1
                    sampleMail = sampleMail & "■ サンプルメール " & mailCount & ":" & vbCrLf & _
                                "件名: " & item.Subject & vbCrLf & _
                                "送信者: " & item.SenderName & vbCrLf & _
                                "受信日: " & item.ReceivedTime & vbCrLf & _
                                "カテゴリ: " & item.Categories & vbCrLf & _
                                "添付ファイル: " & (item.Attachments.Count > 0) & vbCrLf & vbCrLf
                End If
            End If
            On Error GoTo ErrorHandler
        Next i
    End If
    
    If sampleMail = "" Then
        sampleMail = "フォルダ内にメールアイテムが見つかりませんでした。"
    End If
    
    ' AI分析の実行
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_SEARCH & vbCrLf & _
                   "以下の検索フォルダとそのサンプルメールを分析し、このフォルダの検索条件を推測してください。" & vbCrLf & _
                   "以下の観点で分析してください：" & vbCrLf & _
                   "1. 推測される検索条件（送信者、キーワード、日付範囲など）" & vbCrLf & _
                   "2. 検索フォルダの目的と用途" & vbCrLf & _
                   "3. より効果的にするための検索条件の改善提案" & vbCrLf & _
                   "4. サンプルメールに共通する特徴" & vbCrLf & _
                   "5. 検索条件の再作成例（具体的な条件式）"
    
    Dim analysisContent As String
    analysisContent = folderInfo & vbCrLf & _
                     "【サンプルメール】" & vbCrLf & vbCrLf & _
                     sampleMail
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, analysisContent, 2500)
    
    If result <> "" Then
        ShowAnalysisResult "検索条件分析結果: " & selectedFolder.Name, result
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "検索条件分析中にエラーが発生しました。", Err.Description
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

' =============================================================================
' ユーザビリティ向上機能（2024年追加）
' =============================================================================

' 新しい統合UIのエントリーポイント
Public Sub AIヘルパー_統合メニュー()
    ' OutlookAI_MainForm.bas が利用可能な場合は新しいUIを使用
    ' フォールバック: 改良版クラシックメニュー
    Call ShowEnhancedMainMenu
End Sub

' 改良版メインメニュー（日本語エイリアス含む）
Public Sub ShowEnhancedMainMenu()
    Dim choice As String
    Dim menuText As String
    
    menuText = "🤖 " & APP_NAME & " v" & APP_VERSION & vbCrLf & vbCrLf & _
               "📊 メール解析:" & vbCrLf & _
               "  1️⃣ メール内容解析" & vbCrLf & _
               "  2️⃣ 検索フォルダ分析" & vbCrLf & vbCrLf & _
               "✉️ メール作成支援:" & vbCrLf & _
               "  3️⃣ 営業断りメール作成" & vbCrLf & _
               "  4️⃣ 承諾メール作成" & vbCrLf & _
               "  5️⃣ カスタムメール作成" & vbCrLf & vbCrLf & _
               "⚙️ システム管理:" & vbCrLf & _
               "  6️⃣ 設定管理" & vbCrLf & _
               "  7️⃣ API接続テスト" & vbCrLf & vbCrLf & _
               "💡 ヒント: 番号入力の代わりに日本語関数名でも実行可能" & vbCrLf & _
               "   例: 「メール内容解析」関数を直接実行" & vbCrLf & vbCrLf & _
               "実行したい機能の番号を入力してください:"
    
    choice = InputBox(menuText, APP_NAME & " - 統合メニュー")
    
    Select Case choice
        Case "1"
            Call AnalyzeSelectedEmail
        Case "2"
            Call AnalyzeSearchFolders
        Case "3"
            Call CreateRejectionEmail
        Case "4"
            Call CreateAcceptanceEmail
        Case "5"
            Call CreateCustomBusinessEmail
        Case "6"
            Call ManageConfiguration
        Case "7"
            Call TestAPIConnection
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。1-7の番号を入力してください。" & vbCrLf & vbCrLf & _
                       "💡 ヒント: 各機能は日本語関数名でも直接実行できます：" & vbCrLf & _
                       "• メール内容解析" & vbCrLf & _
                       "• 営業断りメール作成" & vbCrLf & _
                       "• 承諾メール作成 など", "入力エラー", vbExclamation
    End Select
End Sub

' =============================================================================
' 日本語関数エイリアス（利便性向上のため）
' =============================================================================

' 📧 メール内容解析：選択されたメールの内容をAIで分析
Public Sub メール内容解析()
    Call AnalyzeSelectedEmail
End Sub

' 📁 検索フォルダ分析：検索フォルダの内容と分類状況を分析
Public Sub 検索フォルダ分析()
    Call AnalyzeSearchFolders
End Sub

' ❌ 営業断りメール：営業メールに対する丁寧な断りメールを作成
Public Sub 営業断りメール作成()
    Call CreateRejectionEmail
End Sub

' ✅ 承諾メール：ビジネス提案への承諾メールを作成
Public Sub 承諾メール作成()
    Call CreateAcceptanceEmail
End Sub

' ✏️ カスタムメール：カスタムプロンプトでビジネスメールを作成
Public Sub カスタムメール作成()
    Call CreateCustomBusinessEmail
End Sub

' 🔧 設定管理：API設定や各種設定の管理
Public Sub 設定管理()
    Call ManageConfiguration
End Sub

' 🔌 API接続テスト：OpenAI APIとの接続状態をテスト
Public Sub API接続テスト()
    Call TestAPIConnection
End Sub

' 🤖 統合メニュー表示：新しい使いやすいメニューを表示
Public Sub 統合メニュー表示()
    Call AIヘルパー_統合メニュー
End Sub

' =============================================================================
' 後方互換性のための関数エイリアス
' =============================================================================

' 従来のメインメニュー（後方互換性のため保持）
' 注意: 新規利用者は「AIヘルパー_統合メニュー」または日本語関数名を推奨
Public Sub メインメニュー表示()
    Call ShowMainMenu
End Sub