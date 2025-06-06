' AI_EmailAnalyzer.bas
' Outlook OpenAI マクロ - メール解析モジュール
' 作成日: 2024
' 説明: メール内容の解析、要約、重要情報抽出機能

Option Explicit

' =============================================================================
' メール解析メイン関数
' =============================================================================

' 選択されたメールを解析する（メニューから呼び出される）
Public Sub AnalyzeSelectedEmail()
    On Error GoTo ErrorHandler
    
    Dim mailItem As Object
    Set mailItem = GetSelectedMailItem()
    
    If mailItem Is Nothing Then
        Exit Sub
    End If
    
    ' 解析メニューの表示
    Dim choice As String
    choice = InputBox("実行する解析を選択してください:" & vbCrLf & vbCrLf & _
                     "1. 転送メールの元内容抽出" & vbCrLf & _
                     "2. メール内容の明確化・要約" & vbCrLf & _
                     "3. 重要情報抽出（期限・対象者・アクション）" & vbCrLf & _
                     "4. 全ての解析を実行" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     APP_NAME & " - メール解析")
    
    Select Case choice
        Case "1"
            Call ExtractOriginalContent(mailItem)
        Case "2"
            Call ClarifyEmailContent(mailItem)
        Case "3"
            Call ExtractImportantInfo(mailItem)
        Case "4"
            Call ComprehensiveAnalysis(mailItem)
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。1-4の番号を入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "メール解析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 転送メール内容抽出
' =============================================================================

' 転送されたメールから元の内容を抽出
Public Sub ExtractOriginalContent(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "転送メールの元内容を抽出中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 転送メール検出のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "このメールが転送メールの場合、以下の作業を行ってください：" & vbCrLf & _
                   "1. 転送履歴を識別" & vbCrLf & _
                   "2. 最初の送信者と送信日時を特定" & vbCrLf & _
                   "3. 元のメール本文を抽出" & vbCrLf & _
                   "4. 元の内容を要約" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【転送状況】" & vbCrLf & _
                   "【元の送信者】" & vbCrLf & _
                   "【元の送信日時】" & vbCrLf & _
                   "【元の件名】" & vbCrLf & _
                   "【元の内容要約】"
    
    Dim userMessage As String
    userMessage = "以下のメールを解析してください：" & vbCrLf & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & vbCrLf & _
                  "送信日時: " & mailItem.SentOn & vbCrLf & vbCrLf & _
                  "本文:" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequestChunked(systemPrompt, userMessage, 3000)
    
    If result <> "" Then
        ' 結果をメッセージボックスで表示
        DisplayAnalysisResult "転送メール分析結果", result
        
        ' 結果を新規メールとして作成するかの確認
        If MsgBox("この結果を新規メールとして下書き作成しますか？", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            CreateDraftEmail "転送メール分析結果 - " & mailItem.Subject, result
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "転送メール解析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' メール内容明確化
' =============================================================================

' メール内容を明確化・要約
Public Sub ClarifyEmailContent(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "メール内容を明確化中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 内容明確化のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "以下の作業を行ってください：" & vbCrLf & _
                   "1. 挨拶文、枕詞、定型文を識別" & vbCrLf & _
                   "2. 核心的な内容を抽出" & vbCrLf & _
                   "3. 明確で簡潔な要約を作成" & vbCrLf & _
                   "4. 次に取るべきアクションを提案" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【要点】" & vbCrLf & _
                   "【詳細内容】" & vbCrLf & _
                   "【推奨アクション】"
    
    Dim userMessage As String
    userMessage = "以下のメールの内容を明確化してください：" & vbCrLf & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & vbCrLf & vbCrLf & _
                  "本文:" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequestChunked(systemPrompt, userMessage, 3000)
    
    If result <> "" Then
        DisplayAnalysisResult "メール内容明確化結果", result
        
        ' Outlookのノートとして保存するかの確認
        If MsgBox("この要約をOutlookのノートとして保存しますか？", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            CreateNote "メール要約 - " & mailItem.Subject, result
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "メール内容明確化中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 重要情報抽出
' =============================================================================

' 期限、対象者、アクション項目を抽出
Public Sub ExtractImportantInfo(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "重要情報を抽出中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 重要情報抽出のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "メールから以下の重要情報を抽出してください：" & vbCrLf & _
                   "1. 期限・締切日（具体的な日時）" & vbCrLf & _
                   "2. 対象者・関係者（名前、役職、連絡先）" & vbCrLf & _
                   "3. 必要なアクション項目" & vbCrLf & _
                   "4. 重要度・緊急度" & vbCrLf & _
                   "5. 場所・イベント情報" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【期限・締切】" & vbCrLf & _
                   "【対象者・関係者】" & vbCrLf & _
                   "【必要なアクション】" & vbCrLf & _
                   "【重要度・緊急度】" & vbCrLf & _
                   "【場所・イベント】" & vbCrLf & _
                   "【その他の重要情報】"
    
    Dim userMessage As String
    userMessage = "以下のメールから重要情報を抽出してください：" & vbCrLf & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & vbCrLf & _
                  "送信日時: " & mailItem.SentOn & vbCrLf & vbCrLf & _
                  "本文:" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequestChunked(systemPrompt, userMessage, 3000)
    
    If result <> "" Then
        DisplayAnalysisResult "重要情報抽出結果", result
        
        ' カレンダーアイテムとして作成するかの確認
        If InStr(result, "期限") > 0 Or InStr(result, "締切") > 0 Then
            If MsgBox("期限情報が検出されました。カレンダーイベントを作成しますか？", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
                CreateCalendarEvent mailItem.Subject, result
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "重要情報抽出中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 包括的分析
' =============================================================================

' 全ての解析を実行
Public Sub ComprehensiveAnalysis(ByVal mailItem As Object)
    On Error GoTo ErrorHandler
    
    ShowProgress "包括的なメール分析を開始中..."
    
    Dim emailBody As String
    emailBody = GetEmailBodyText(mailItem)
    
    If Not CheckContentLength(emailBody) Then
        Exit Sub
    End If
    
    ' 包括的分析のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_ANALYZER & vbCrLf & _
                   "このメールの包括的分析を行ってください：" & vbCrLf & _
                   "1. メールの種類・目的の分類" & vbCrLf & _
                   "2. 内容の要約と明確化" & vbCrLf & _
                   "3. 重要情報の抽出" & vbCrLf & _
                   "4. 感情的なトーンの分析" & vbCrLf & _
                   "5. 推奨される返信方針" & vbCrLf & _
                   "6. 優先度の評価" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【メール分類】" & vbCrLf & _
                   "【内容要約】" & vbCrLf & _
                   "【重要情報】" & vbCrLf & _
                   "【感情・トーン】" & vbCrLf & _
                   "【返信方針】" & vbCrLf & _
                   "【優先度評価】" & vbCrLf & _
                   "【推奨アクション】"
    
    Dim userMessage As String
    userMessage = "以下のメールの包括的分析を行ってください：" & vbCrLf & vbCrLf & _
                  "件名: " & mailItem.Subject & vbCrLf & _
                  "送信者: " & mailItem.SenderName & vbCrLf & _
                  "送信日時: " & mailItem.SentOn & vbCrLf & _
                  "重要度: " & GetImportanceText(mailItem.Importance) & vbCrLf & vbCrLf & _
                  "本文:" & vbCrLf & emailBody
    
    Dim result As String
    result = SendOpenAIRequestChunked(systemPrompt, userMessage, 2500)
    
    If result <> "" Then
        DisplayAnalysisResult "包括的メール分析結果", result
        
        ' 結果保存のオプション提示
        Dim saveChoice As Integer
        saveChoice = MsgBox("この分析結果をどのように保存しますか？" & vbCrLf & vbCrLf & _
                           "はい: 新規メールとして下書き保存" & vbCrLf & _
                           "いいえ: Outlookノートとして保存" & vbCrLf & _
                           "キャンセル: 保存しない", _
                           vbYesNoCancel + vbQuestion, APP_NAME)
        
        Select Case saveChoice
            Case vbYes
                CreateDraftEmail "メール分析結果 - " & mailItem.Subject, result
            Case vbNo
                CreateNote "メール分析 - " & mailItem.Subject, result
        End Select
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "包括的メール分析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' ヘルパー関数
' =============================================================================

' メール本文テキストの取得
Private Function GetEmailBodyText(ByVal mailItem As Object) As String
    On Error GoTo ErrorHandler
    
    Dim bodyText As String
    
    ' プレーンテキストがある場合はそれを使用、なければHTMLからテキストを抽出
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

' 重要度のテキスト表示
Private Function GetImportanceText(ByVal importance As Integer) As String
    Select Case importance
        Case 0: GetImportanceText = "低"
        Case 1: GetImportanceText = "標準"
        Case 2: GetImportanceText = "高"
        Case Else: GetImportanceText = "不明"
    End Select
End Function

' 分析結果の表示
Private Sub DisplayAnalysisResult(ByVal title As String, ByVal result As String)
    ' 長いテキストの場合は、専用のフォームを使用するか、分割表示
    If Len(result) > 1000 Then
        ' 長いテキストの場合は、最初の部分を表示
        Dim shortResult As String
        shortResult = Left(result, 800) & vbCrLf & vbCrLf & "（結果が長いため、一部のみ表示）"
        ShowMessage shortResult, title
        
        ' 全文を見るかの確認
        If MsgBox("全文を確認しますか？", vbYesNo + vbQuestion, title) = vbYes Then
            ' 全文表示（複数のメッセージボックスに分割）
            ShowLongText result, title
        End If
    Else
        ShowMessage result, title
    End If
End Sub

' 長いテキストの分割表示
Private Sub ShowLongText(ByVal text As String, ByVal title As String)
    Const maxLength As Integer = 1500
    Dim currentPos As Integer
    Dim pageNum As Integer
    
    currentPos = 1
    pageNum = 1
    
    While currentPos <= Len(text)
        Dim pageText As String
        pageText = Mid(text, currentPos, maxLength)
        
        Dim totalPages As Integer
        totalPages = Int((Len(text) - 1) / maxLength) + 1
        
        MsgBox pageText, vbInformation, title & " (" & pageNum & "/" & totalPages & ")"
        
        currentPos = currentPos + maxLength
        pageNum = pageNum + 1
    Wend
End Sub

' 下書きメールの作成
Private Sub CreateDraftEmail(ByVal subject As String, ByVal body As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim newMail As Object
    
    Set olApp = Application
    Set newMail = olApp.CreateItem(olMailItem)
    
    newMail.Subject = subject
    newMail.Body = body
    newMail.Save
    
    ShowSuccess "下書きメールを作成しました。"
    
    Exit Sub
    
ErrorHandler:
    ShowError "下書きメール作成中にエラーが発生しました。", Err.Description
End Sub

' Outlookノートの作成
Private Sub CreateNote(ByVal subject As String, ByVal body As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim newNote As Object
    
    Set olApp = Application
    Set newNote = olApp.CreateItem(olNoteItem)
    
    newNote.Body = subject & vbCrLf & vbCrLf & body
    newNote.Save
    
    ShowSuccess "Outlookノートを作成しました。"
    
    Exit Sub
    
ErrorHandler:
    ShowError "ノート作成中にエラーが発生しました。", Err.Description
End Sub

' カレンダーイベントの作成
Private Sub CreateCalendarEvent(ByVal subject As String, ByVal details As String)
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim newEvent As Object
    
    Set olApp = Application
    Set newEvent = olApp.CreateItem(olAppointmentItem)
    
    newEvent.Subject = subject
    newEvent.Body = details
    ' デフォルト設定
    newEvent.Start = Date + 1 ' 明日
    newEvent.Duration = 60 ' 1時間
    newEvent.Save
    
    ShowSuccess "カレンダーイベントを作成しました。日時を調整してください。"
    
    Exit Sub
    
ErrorHandler:
    ShowError "カレンダーイベント作成中にエラーが発生しました。", Err.Description
End Sub