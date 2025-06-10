' AI_EmailComposer.bas
' Outlook OpenAI マクロ - メール作成支援モジュール
' 作成日: 2024
' 説明: 営業メール断り、承諾メール等の自動作成機能

Option Explicit

' =============================================================================
' 型定義
' =============================================================================

' メール内容を格納する型
Private Type EmailContent
    Subject As String
    Body As String
End Type

' =============================================================================
' メール作成支援メイン関数
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

' =============================================================================
' 営業メール断り作成
' =============================================================================

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
        
        ShowSuccess "営業断りメールの返信ウィンドウを開きました。内容を確認してから送信してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "営業断りメール生成中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 承諾メール作成
' =============================================================================

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
        
        ShowSuccess "承諾メールの返信ウィンドウを開きました。内容を確認してから送信してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "承諾メール生成中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' ビジネスメールテンプレート作成
' =============================================================================

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
        
        ShowSuccess "カスタムメールの編集ウィンドウを開きました。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムメール生成中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' ヘルパー関数
' =============================================================================

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

' メール本文テキストの取得（EmailAnalyzerと同じ関数）
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