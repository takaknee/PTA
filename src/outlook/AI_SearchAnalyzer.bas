' AI_SearchAnalyzer.bas
' Outlook OpenAI マクロ - 検索フォルダ分析モジュール
' 作成日: 2024
' 説明: 検索フォルダの使用状況分析と新規作成支援機能

Option Explicit

' =============================================================================
' 検索フォルダ分析メイン関数
' =============================================================================

' 検索フォルダ分析（メニューから呼び出される）
Public Sub AnalyzeSearchFolders()
    On Error GoTo ErrorHandler
    
    ' 分析メニューの表示
    Dim choice As String
    choice = InputBox("実行する分析を選択してください:" & vbCrLf & vbCrLf & _
                     "1. 既存検索フォルダの使用状況分析" & vbCrLf & _
                     "2. フォルダ内メール分析（新規検索フォルダ作成支援）" & vbCrLf & _
                     "3. 重複・不要検索フォルダの検出" & vbCrLf & _
                     "4. 包括的な検索フォルダ最適化提案" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     APP_NAME & " - 検索フォルダ分析")
    
    Select Case choice
        Case "1"
            Call AnalyzeExistingSearchFolders
        Case "2"
            Call AnalyzeFolderForNewSearch
        Case "3"
            Call DetectDuplicateSearchFolders
        Case "4"
            Call ComprehensiveSearchFolderAnalysis
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。1-4の番号を入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "検索フォルダ分析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 既存検索フォルダ分析
' =============================================================================

' 既存の検索フォルダの使用状況を分析
Private Sub AnalyzeExistingSearchFolders()
    On Error GoTo ErrorHandler
    
    ShowProgress "検索フォルダの使用状況を分析中..."
    
    ' 検索フォルダの情報を収集
    Dim searchFolderInfo As String
    searchFolderInfo = CollectSearchFolderInfo()
    
    If searchFolderInfo = "" Then
        ShowMessage "検索フォルダが見つかりませんでした。", "情報", vbInformation
        Exit Sub
    End If
    
    ' AI分析のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_SEARCH & vbCrLf & _
                   "検索フォルダの使用状況を分析し、以下の観点から評価してください：" & vbCrLf & _
                   "1. 各検索フォルダの有効性（メール数、最終更新日から判断）" & vbCrLf & _
                   "2. 不要と思われる検索フォルダの特定" & vbCrLf & _
                   "3. 改善提案（統合、削除、条件変更等）" & vbCrLf & _
                   "4. 使用頻度に基づく優先度設定" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【分析サマリー】" & vbCrLf & _
                   "【削除推奨フォルダ】" & vbCrLf & _
                   "【統合推奨フォルダ】" & vbCrLf & _
                   "【改善提案】" & vbCrLf & _
                   "【アクションプラン】"
    
    Dim userMessage As String
    userMessage = "以下の検索フォルダ情報を分析してください：" & vbCrLf & vbCrLf & searchFolderInfo
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, 2000)
    
    If result <> "" Then
        DisplayAnalysisResult "検索フォルダ使用状況分析", result
        
        ' 分析結果をExcelファイルに出力するかの確認
        If MsgBox("この分析結果をExcelファイルとして保存しますか？", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            SaveAnalysisToExcel "検索フォルダ分析", result
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "検索フォルダ分析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' フォルダ内メール分析
' =============================================================================

' 指定フォルダのメール分析（新規検索フォルダ作成支援）
Private Sub AnalyzeFolderForNewSearch()
    On Error GoTo ErrorHandler
    
    ' フォルダ選択
    Dim selectedFolder As Object
    Set selectedFolder = Application.GetNamespace("MAPI").PickFolder
    
    If selectedFolder Is Nothing Then
        ShowMessage "フォルダが選択されませんでした。", "キャンセル", vbInformation
        Exit Sub
    End If
    
    ShowProgress "フォルダ「" & selectedFolder.Name & "」のメールを分析中..."
    
    ' フォルダ内メール情報を収集
    Dim folderAnalysis As String
    folderAnalysis = AnalyzeFolderContents(selectedFolder)
    
    If folderAnalysis = "" Then
        ShowMessage "フォルダにメールが見つかりませんでした。", "情報", vbInformation
        Exit Sub
    End If
    
    ' AI分析のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_SEARCH & vbCrLf & _
                   "フォルダ内のメール分析結果から、効果的な検索フォルダの作成を提案してください：" & vbCrLf & _
                   "1. メールの分類傾向の特定" & vbCrLf & _
                   "2. 推奨検索条件の提案" & vbCrLf & _
                   "3. 複数の検索フォルダの組み合わせ提案" & vbCrLf & _
                   "4. 検索フォルダ名の提案" & vbCrLf & _
                   "5. 実用性の高い条件設定" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【フォルダ分析サマリー】" & vbCrLf & _
                   "【推奨検索フォルダ1】" & vbCrLf & _
                   "【推奨検索フォルダ2】" & vbCrLf & _
                   "【推奨検索フォルダ3】" & vbCrLf & _
                   "【実装手順】"
    
    Dim userMessage As String
    userMessage = "以下のフォルダ「" & selectedFolder.Name & "」の分析結果から、検索フォルダ作成を提案してください：" & vbCrLf & vbCrLf & folderAnalysis
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, 2000)
    
    If result <> "" Then
        DisplayAnalysisResult "検索フォルダ作成提案", result
        
        ' 自動作成の提案
        If MsgBox("提案された検索フォルダを自動作成しますか？" & vbCrLf & _
                  "（手動で条件を確認・調整することをお勧めします）", _
                  vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            CreateSearchFoldersFromAnalysis result, selectedFolder
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "フォルダ分析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 重複検索フォルダ検出
' =============================================================================

' 重複・類似検索フォルダの検出
Private Sub DetectDuplicateSearchFolders()
    On Error GoTo ErrorHandler
    
    ShowProgress "重複・類似検索フォルダを検出中..."
    
    ' 検索フォルダの詳細情報を収集
    Dim detailedInfo As String
    detailedInfo = CollectDetailedSearchFolderInfo()
    
    If detailedInfo = "" Then
        ShowMessage "検索フォルダが見つかりませんでした。", "情報", vbInformation
        Exit Sub
    End If
    
    ' AI分析のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_SEARCH & vbCrLf & _
                   "検索フォルダの情報から重複や類似条件を検出し、統合提案を行ってください：" & vbCrLf & _
                   "1. 完全に重複している検索フォルダの特定" & vbCrLf & _
                   "2. 類似条件で統合可能なフォルダの特定" & vbCrLf & _
                   "3. 統合方法の具体的提案" & vbCrLf & _
                   "4. 統合によるメリットの説明" & vbCrLf & _
                   "5. 統合時の注意事項" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【重複検出結果】" & vbCrLf & _
                   "【統合推奨グループ】" & vbCrLf & _
                   "【統合手順】" & vbCrLf & _
                   "【期待効果】"
    
    Dim userMessage As String
    userMessage = "以下の検索フォルダ情報から重複・類似フォルダを検出してください：" & vbCrLf & vbCrLf & detailedInfo
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userMessage, 2000)
    
    If result <> "" Then
        DisplayAnalysisResult "重複検索フォルダ検出結果", result
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "重複検索フォルダ検出中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' 包括的分析
' =============================================================================

' 包括的な検索フォルダ最適化提案
Private Sub ComprehensiveSearchFolderAnalysis()
    On Error GoTo ErrorHandler
    
    ShowProgress "包括的な検索フォルダ分析を実行中..."
    
    ' 全ての情報を収集
    Dim searchFolderInfo As String
    Dim mailboxInfo As String
    
    searchFolderInfo = CollectDetailedSearchFolderInfo()
    mailboxInfo = CollectMailboxStatistics()
    
    ' AI分析のプロンプト
    Dim systemPrompt As String
    systemPrompt = SYSTEM_PROMPT_SEARCH & vbCrLf & _
                   "メールボックス全体の包括的分析を行い、最適な検索フォルダ戦略を提案してください：" & vbCrLf & _
                   "1. 現状の問題点と改善点の特定" & vbCrLf & _
                   "2. ユーザーのメール使用パターンの分析" & vbCrLf & _
                   "3. 効率的な検索フォルダ構成の提案" & vbCrLf & _
                   "4. 段階的な実装計画の作成" & vbCrLf & _
                   "5. 継続的なメンテナンス方法の提案" & vbCrLf & vbCrLf & _
                   "出力形式：" & vbCrLf & _
                   "【現状分析】" & vbCrLf & _
                   "【問題点と改善点】" & vbCrLf & _
                   "【最適化戦略】" & vbCrLf & _
                   "【実装計画】" & vbCrLf & _
                   "【メンテナンス方法】"
    
    Dim userMessage As String
    userMessage = "以下の情報から包括的な検索フォルダ最適化を提案してください：" & vbCrLf & vbCrLf & _
                  "【検索フォルダ情報】" & vbCrLf & searchFolderInfo & vbCrLf & vbCrLf & _
                  "【メールボックス統計】" & vbCrLf & mailboxInfo
    
    Dim result As String
    result = SendOpenAIRequestChunked(systemPrompt, userMessage, 2000)
    
    If result <> "" Then
        DisplayAnalysisResult "包括的検索フォルダ最適化提案", result
        
        ' レポート保存の提案
        Dim saveChoice As Integer
        saveChoice = MsgBox("この最適化提案をどのように保存しますか？" & vbCrLf & vbCrLf & _
                           "はい: Excelファイルとして保存" & vbCrLf & _
                           "いいえ: テキストファイルとして保存" & vbCrLf & _
                           "キャンセル: 保存しない", _
                           vbYesNoCancel + vbQuestion, APP_NAME)
        
        Select Case saveChoice
            Case vbYes
                SaveAnalysisToExcel "検索フォルダ最適化提案", result
            Case vbNo
                SaveAnalysisToText "検索フォルダ最適化提案", result
        End Select
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "包括的分析中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' データ収集関数
' =============================================================================

' 検索フォルダ情報の収集
Private Function CollectSearchFolderInfo() As String
    On Error GoTo ErrorHandler
    
    Dim info As String
    Dim olNamespace As Object
    Dim searchFolders As Object
    Dim folder As Object
    
    Set olNamespace = Application.GetNamespace("MAPI")
    Set searchFolders = olNamespace.GetDefaultFolder(olFolderInbox).Parent.Folders("検索フォルダー")
    
    info = "=== 検索フォルダ一覧 ===" & vbCrLf & vbCrLf
    
    Dim folderCount As Integer
    folderCount = 0
    
    For Each folder In searchFolders.Folders
        folderCount = folderCount + 1
        info = info & "フォルダ" & folderCount & ": " & folder.Name & vbCrLf
        info = info & "  メール数: " & folder.Items.Count & vbCrLf
        
        ' 最新メールの日付を取得
        If folder.Items.Count > 0 Then
            Dim latestDate As Date
            latestDate = GetLatestEmailDate(folder)
            info = info & "  最新メール: " & Format(latestDate, "yyyy/mm/dd") & vbCrLf
        Else
            info = info & "  最新メール: なし" & vbCrLf
        End If
        
        info = info & vbCrLf
    Next folder
    
    If folderCount = 0 Then
        CollectSearchFolderInfo = ""
    Else
        CollectSearchFolderInfo = info
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog "検索フォルダ情報収集エラー: " & Err.Description, "ERROR"
    CollectSearchFolderInfo = ""
End Function

' 詳細検索フォルダ情報の収集
Private Function CollectDetailedSearchFolderInfo() As String
    On Error GoTo ErrorHandler
    
    Dim info As String
    info = CollectSearchFolderInfo()
    
    ' 今回はシンプルな実装とし、基本情報のみ返す
    CollectDetailedSearchFolderInfo = info
    
    Exit Function
    
ErrorHandler:
    WriteLog "詳細検索フォルダ情報収集エラー: " & Err.Description, "ERROR"
    CollectDetailedSearchFolderInfo = ""
End Function

' フォルダ内容の分析
Private Function AnalyzeFolderContents(ByVal targetFolder As Object) As String
    On Error GoTo ErrorHandler
    
    Dim analysis As String
    Dim items As Object
    Dim mailItem As Object
    Dim senderStats As Object
    Dim subjectStats As Object
    
    Set items = targetFolder.Items
    Set senderStats = CreateObject("Scripting.Dictionary")
    Set subjectStats = CreateObject("Scripting.Dictionary")
    
    analysis = "=== フォルダ分析: " & targetFolder.Name & " ===" & vbCrLf & vbCrLf
    analysis = analysis & "総メール数: " & items.Count & vbCrLf & vbCrLf
    
    If items.Count = 0 Then
        AnalyzeFolderContents = analysis & "メールがありません。"
        Exit Function
    End If
    
    ' サンプル分析（最大100件）
    Dim analyzeCount As Integer
    analyzeCount = IIf(items.Count > 100, 100, items.Count)
    
    Dim i As Integer
    For i = 1 To analyzeCount
        Set mailItem = items.Item(i)
        
        ' 送信者統計
        Dim sender As String
        sender = ExtractEmailAddress(mailItem.SenderEmailAddress)
        If senderStats.Exists(sender) Then
            senderStats(sender) = senderStats(sender) + 1
        Else
            senderStats.Add sender, 1
        End If
        
        ' 件名パターン統計（簡易版）
        Dim subjectPattern As String
        subjectPattern = AnalyzeSubjectPattern(mailItem.Subject)
        If subjectStats.Exists(subjectPattern) Then
            subjectStats(subjectPattern) = subjectStats(subjectPattern) + 1
        Else
            subjectStats.Add subjectPattern, 1
        End If
    Next i
    
    ' 統計結果を追加
    analysis = analysis & "=== 主要送信者 TOP5 ===" & vbCrLf
    analysis = analysis & GetTopSenders(senderStats, 5) & vbCrLf
    
    analysis = analysis & "=== 件名パターン TOP5 ===" & vbCrLf
    analysis = analysis & GetTopSubjectPatterns(subjectStats, 5) & vbCrLf
    
    AnalyzeFolderContents = analysis
    
    Exit Function
    
ErrorHandler:
    WriteLog "フォルダ内容分析エラー: " & Err.Description, "ERROR"
    AnalyzeFolderContents = ""
End Function

' メールボックス統計の収集
Private Function CollectMailboxStatistics() As String
    On Error GoTo ErrorHandler
    
    Dim stats As String
    Dim olNamespace As Object
    Dim inbox As Object
    Dim sentItems As Object
    
    Set olNamespace = Application.GetNamespace("MAPI")
    Set inbox = olNamespace.GetDefaultFolder(olFolderInbox)
    Set sentItems = olNamespace.GetDefaultFolder(olFolderSentMail)
    
    stats = "=== メールボックス統計 ===" & vbCrLf & vbCrLf
    stats = stats & "受信トレイ: " & inbox.Items.Count & " 件" & vbCrLf
    stats = stats & "送信済みアイテム: " & sentItems.Items.Count & " 件" & vbCrLf
    
    CollectMailboxStatistics = stats
    
    Exit Function
    
ErrorHandler:
    WriteLog "メールボックス統計収集エラー: " & Err.Description, "ERROR"
    CollectMailboxStatistics = ""
End Function

' =============================================================================
' ヘルパー関数
' =============================================================================

' フォルダ内の最新メール日付を取得
Private Function GetLatestEmailDate(ByVal folder As Object) As Date
    On Error GoTo ErrorHandler
    
    If folder.Items.Count = 0 Then
        GetLatestEmailDate = CDate("1900/01/01")
        Exit Function
    End If
    
    ' 最初のアイテムの日付を取得（既に日付順でソートされていると仮定）
    GetLatestEmailDate = folder.Items.Item(1).ReceivedTime
    
    Exit Function
    
ErrorHandler:
    GetLatestEmailDate = CDate("1900/01/01")
End Function

' メールアドレスの抽出
Private Function ExtractEmailAddress(ByVal fullAddress As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    
    startPos = InStr(fullAddress, "<")
    endPos = InStr(fullAddress, ">")
    
    If startPos > 0 And endPos > startPos Then
        ExtractEmailAddress = Mid(fullAddress, startPos + 1, endPos - startPos - 1)
    Else
        ExtractEmailAddress = fullAddress
    End If
End Function

' 件名パターンの分析
Private Function AnalyzeSubjectPattern(ByVal subject As String) As String
    ' 簡易的なパターン分析
    If InStr(subject, "Re:") > 0 Then
        AnalyzeSubjectPattern = "返信メール"
    ElseIf InStr(subject, "Fwd:") > 0 Or InStr(subject, "転送:") > 0 Then
        AnalyzeSubjectPattern = "転送メール"
    ElseIf InStr(subject, "会議") > 0 Or InStr(subject, "ミーティング") > 0 Then
        AnalyzeSubjectPattern = "会議関連"
    ElseIf InStr(subject, "お疲れ") > 0 Then
        AnalyzeSubjectPattern = "日常業務"
    Else
        AnalyzeSubjectPattern = "その他"
    End If
End Function

' TOP送信者の取得
Private Function GetTopSenders(ByVal senderStats As Object, ByVal topCount As Integer) As String
    ' 簡易実装：辞書の内容を表示
    Dim result As String
    Dim keys As Variant
    Dim i As Integer
    
    keys = senderStats.Keys
    
    result = ""
    For i = 0 To UBound(keys)
        If i >= topCount Then Exit For
        result = result & (i + 1) & ". " & keys(i) & " (" & senderStats(keys(i)) & "件)" & vbCrLf
    Next i
    
    GetTopSenders = result
End Function

' TOP件名パターンの取得
Private Function GetTopSubjectPatterns(ByVal subjectStats As Object, ByVal topCount As Integer) As String
    ' 簡易実装：辞書の内容を表示
    Dim result As String
    Dim keys As Variant
    Dim i As Integer
    
    keys = subjectStats.Keys
    
    result = ""
    For i = 0 To UBound(keys)
        If i >= topCount Then Exit For
        result = result & (i + 1) & ". " & keys(i) & " (" & subjectStats(keys(i)) & "件)" & vbCrLf
    Next i
    
    GetTopSubjectPatterns = result
End Function

' 分析結果の表示（EmailAnalyzerからコピー）
Private Sub DisplayAnalysisResult(ByVal title As String, ByVal result As String)
    If Len(result) > 1000 Then
        Dim shortResult As String
        shortResult = Left(result, 800) & vbCrLf & vbCrLf & "（結果が長いため、一部のみ表示）"
        ShowMessage shortResult, title
        
        If MsgBox("全文を確認しますか？", vbYesNo + vbQuestion, title) = vbYes Then
            ShowLongText result, title
        End If
    Else
        ShowMessage result, title
    End If
End Sub

' 長いテキストの分割表示（EmailAnalyzerからコピー）
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

' Excel保存機能
Private Sub SaveAnalysisToExcel(ByVal analysisType As String, ByVal content As String)
    On Error GoTo ErrorHandler
    
    ' 簡易実装：Excelアプリケーションを起動してデータを保存
    Dim xlApp As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlWorksheet = xlWorkbook.Worksheets(1)
    
    xlWorksheet.Cells(1, 1).Value = analysisType
    xlWorksheet.Cells(2, 1).Value = "作成日時: " & Now
    xlWorksheet.Cells(4, 1).Value = content
    
    xlApp.Visible = True
    
    ShowSuccess "分析結果をExcelファイルに出力しました。"
    
    Exit Sub
    
ErrorHandler:
    ShowError "Excel保存中にエラーが発生しました。", Err.Description
End Sub

' テキスト保存機能
Private Sub SaveAnalysisToText(ByVal analysisType As String, ByVal content As String)
    On Error GoTo ErrorHandler
    
    ' ファイル保存ダイアログの代わりに固定パスに保存
    Dim fileName As String
    fileName = Environ("USERPROFILE") & "\Desktop\" & analysisType & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open fileName For Output As #fileNum
    Print #fileNum, analysisType
    Print #fileNum, "作成日時: " & Now
    Print #fileNum, ""
    Print #fileNum, content
    Close #fileNum
    
    ShowSuccess "分析結果をテキストファイルに保存しました。" & vbCrLf & fileName
    
    Exit Sub
    
ErrorHandler:
    ShowError "テキストファイル保存中にエラーが発生しました。", Err.Description
End Sub

' AI分析結果から検索フォルダを作成
Private Sub CreateSearchFoldersFromAnalysis(ByVal analysis As String, ByVal baseFolder As Object)
    On Error GoTo ErrorHandler
    
    ShowMessage "自動検索フォルダ作成機能は現在準備中です。" & vbCrLf & _
               "手動でOutlookの検索フォルダ機能を使用して作成してください。", _
               "機能準備中", vbInformation
    
    Exit Sub
    
ErrorHandler:
    ShowError "検索フォルダ作成中にエラーが発生しました。", Err.Description
End Sub