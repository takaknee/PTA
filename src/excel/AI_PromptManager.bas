' AI_PromptManager.bas
' Excel AI Helper - プロンプト管理モジュール
' 作成日: 2024
' 説明: プロンプトテンプレートの管理、カスタムプロンプトの保存・読み込み

Option Explicit

' =============================================================================
' プロンプト管理メイン関数
' =============================================================================

' プロンプト管理メインメニュー
Public Sub ShowPromptManager()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    Dim menuText As String
    
    menuText = "プロンプト管理メニュー" & vbCrLf & vbCrLf & _
               "1. プリセットプロンプト一覧" & vbCrLf & _
               "2. カスタムプロンプト管理" & vbCrLf & _
               "3. プロンプトテンプレート編集" & vbCrLf & _
               "4. プロンプト効果測定" & vbCrLf & _
               "5. プロンプトのインポート/エクスポート" & vbCrLf & vbCrLf & _
               "実行したい項目の番号を入力してください:"
    
    choice = InputBox(menuText, APP_NAME & " - プロンプト管理")
    
    Select Case choice
        Case "1"
            Call ShowPresetPrompts
        Case "2"
            Call ManageCustomPrompts
        Case "3"
            Call EditPromptTemplates
        Case "4"
            Call MeasurePromptEffectiveness
        Case "5"
            Call ImportExportPrompts
        Case ""
            ' キャンセル
        Case Else
            ShowWarning "無効な選択です。1-5の番号を入力してください。"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "プロンプト管理中にエラーが発生しました", Err.Description
    LogError "ShowPromptManager", Err.Description
End Sub

' =============================================================================
' プリセットプロンプト管理
' =============================================================================

' プリセットプロンプト一覧表示
Private Sub ShowPresetPrompts()
    On Error GoTo ErrorHandler
    
    Dim promptList As String
    promptList = "【表分析用プリセットプロンプト】" & vbCrLf & vbCrLf & _
                "1. 学習意欲判定" & vbCrLf & _
                GetPromptTemplate("LEARNING_MOTIVATION") & vbCrLf & vbCrLf & _
                "2. 感情分析" & vbCrLf & _
                GetPromptTemplate("SENTIMENT_ANALYSIS") & vbCrLf & vbCrLf & _
                "3. カテゴリ分類" & vbCrLf & _
                GetPromptTemplate("CATEGORY_CLASSIFICATION") & vbCrLf & vbCrLf & _
                "4. 要約生成" & vbCrLf & _
                GetPromptTemplate("SUMMARIZATION") & vbCrLf & vbCrLf & _
                "【セル処理用プリセットプロンプト】" & vbCrLf & vbCrLf & _
                "5. Excel関数生成" & vbCrLf & _
                GetPromptTemplate("EXCEL_FUNCTION") & vbCrLf & vbCrLf & _
                "6. データ変換" & vbCrLf & _
                GetPromptTemplate("DATA_TRANSFORMATION")
    
    ' 結果表示（長いテキストなので分割表示）
    Call DisplayLongText("プリセットプロンプト一覧", promptList)
    
    Exit Sub
    
ErrorHandler:
    ShowError "プリセットプロンプト表示中にエラーが発生しました", Err.Description
    LogError "ShowPresetPrompts", Err.Description
End Sub

' プロンプトテンプレート取得
Private Function GetPromptTemplate(ByVal templateType As String) As String
    Select Case templateType
        Case "LEARNING_MOTIVATION"
            GetPromptTemplate = "以下の行データを分析し、学習意欲の高さを0-10のスコアで評価してください。" & vbCrLf & _
                              "評価基準:" & vbCrLf & _
                              "- 参加動機の積極性 (0-3点)" & vbCrLf & _
                              "- 学習内容への関心度 (0-3点)" & vbCrLf & _
                              "- 継続意志の強さ (0-4点)" & vbCrLf & _
                              "スコアのみを数値で回答してください。"
                              
        Case "SENTIMENT_ANALYSIS"
            GetPromptTemplate = "以下の行データを分析し、感情を判定してください。" & vbCrLf & _
                              "判定基準:" & vbCrLf & _
                              "- ポジティブ: 前向き、喜び、満足、希望などの感情" & vbCrLf & _
                              "- ネガティブ: 不満、怒り、悲しみ、不安などの感情" & vbCrLf & _
                              "- 中性: 客観的、事実的、感情が読み取れない" & vbCrLf & _
                              "「ポジティブ」「ネガティブ」「中性」のいずれかで回答してください。"
                              
        Case "CATEGORY_CLASSIFICATION"
            GetPromptTemplate = "以下の行データを分析し、最も適切なカテゴリを判定してください。" & vbCrLf & _
                              "共通カテゴリ例:" & vbCrLf & _
                              "- 教育・学習 - ビジネス・仕事 - 個人・プライベート" & vbCrLf & _
                              "- 技術・IT - 健康・医療 - その他" & vbCrLf & _
                              "最も適切なカテゴリ名のみを回答してください。"
                              
        Case "SUMMARIZATION"
            GetPromptTemplate = "以下の行データを分析し、50文字以内で要約してください。" & vbCrLf & _
                              "要約のポイント:" & vbCrLf & _
                              "- 最も重要な情報を抽出" & vbCrLf & _
                              "- 簡潔で分かりやすい表現" & vbCrLf & _
                              "- 文脈を保持" & vbCrLf & _
                              "要約のみを回答してください。"
                              
        Case "EXCEL_FUNCTION"
            GetPromptTemplate = "以下のデータまたは要求を基に、適切なExcel関数を生成してください。" & vbCrLf & _
                              "関数生成の指針:" & vbCrLf & _
                              "- 実用的で効率的な関数" & vbCrLf & _
                              "- エラーハンドリングを含む" & vbCrLf & _
                              "- 可読性の高い構造" & vbCrLf & _
                              "関数のみを回答してください（=で始まる）。"
                              
        Case "DATA_TRANSFORMATION"
            GetPromptTemplate = "以下のデータを指定された形式に変換してください。" & vbCrLf & _
                              "変換の指針:" & vbCrLf & _
                              "- データの整合性を保持" & vbCrLf & _
                              "- 適切な形式とフォーマット" & vbCrLf & _
                              "- 情報の損失を最小化" & vbCrLf & _
                              "変換後のデータのみを回答してください。"
                              
        Case Else
            GetPromptTemplate = "テンプレートが見つかりません: " & templateType
    End Select
End Function

' =============================================================================
' カスタムプロンプト管理
' =============================================================================

' カスタムプロンプト管理メニュー
Private Sub ManageCustomPrompts()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    choice = InputBox("カスタムプロンプト管理" & vbCrLf & vbCrLf & _
                     "1. 新規カスタムプロンプト作成" & vbCrLf & _
                     "2. カスタムプロンプト一覧" & vbCrLf & _
                     "3. カスタムプロンプト編集" & vbCrLf & _
                     "4. カスタムプロンプト削除" & vbCrLf & _
                     "5. カスタムプロンプト使用" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     "カスタムプロンプト管理")
    
    Select Case choice
        Case "1"
            Call CreateCustomPrompt
        Case "2"
            Call ShowCustomPromptList
        Case "3"
            Call EditCustomPrompt
        Case "4"
            Call DeleteCustomPrompt
        Case "5"
            Call UseCustomPrompt
        Case ""
            ' キャンセル
        Case Else
            ShowWarning "無効な選択です。1-5の番号を入力してください。"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムプロンプト管理中にエラーが発生しました", Err.Description
    LogError "ManageCustomPrompts", Err.Description
End Sub

' 新規カスタムプロンプト作成
Private Sub CreateCustomPrompt()
    On Error GoTo ErrorHandler
    
    ' プロンプト名の入力
    Dim promptName As String
    promptName = InputBox("カスタムプロンプトの名前を入力してください:" & vbCrLf & vbCrLf & _
                         "例: 営業成績分析、商品レビュー評価、など", _
                         "プロンプト名", "新規プロンプト")
    
    If promptName = "" Then Exit Sub
    
    ' 既存名前の重複チェック
    If IsCustomPromptExists(promptName) Then
        If Not ShowConfirm("同じ名前のプロンプトが既に存在します。上書きしますか？") Then
            Exit Sub
        End If
    End If
    
    ' カテゴリの選択
    Dim category As String
    Dim categoryChoice As String
    categoryChoice = InputBox("プロンプトのカテゴリを選択してください:" & vbCrLf & vbCrLf & _
                             "1. 表分析用" & vbCrLf & _
                             "2. セル処理用" & vbCrLf & _
                             "3. 汎用" & vbCrLf & vbCrLf & _
                             "番号を入力してください:", _
                             "カテゴリ選択", "1")
    
    Select Case categoryChoice
        Case "1"
            category = "TABLE_ANALYSIS"
        Case "2"
            category = "CELL_PROCESSING"
        Case "3"
            category = "GENERAL"
        Case Else
            category = "GENERAL"
    End Select
    
    ' プロンプト内容の入力
    Dim promptContent As String
    promptContent = InputBox("プロンプトの内容を入力してください:" & vbCrLf & vbCrLf & _
                            "※ データは自動的に追加されるため、分析指示のみ記述してください。", _
                            "プロンプト内容")
    
    If promptContent = "" Then Exit Sub
    
    ' 説明の入力
    Dim description As String
    description = InputBox("プロンプトの説明を入力してください（省略可能）:", _
                          "説明", "")
    
    ' カスタムプロンプトの保存
    Call SaveCustomPrompt(promptName, category, promptContent, description)
    
    ShowSuccess "カスタムプロンプト「" & promptName & "」を保存しました。"
    LogInfo "カスタムプロンプト作成: " & promptName
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムプロンプト作成中にエラーが発生しました", Err.Description
    LogError "CreateCustomPrompt", Err.Description
End Sub

' カスタムプロンプトの保存
Private Sub SaveCustomPrompt(ByVal promptName As String, ByVal category As String, ByVal content As String, ByVal description As String)
    On Error GoTo ErrorHandler
    
    ' カスタムプロンプトをレジストリに保存
    Dim promptKey As String
    promptKey = "CustomPrompts\" & promptName & "\"
    
    Call SaveAPISetting(promptKey & "Category", category)
    Call SaveAPISetting(promptKey & "Content", content)
    Call SaveAPISetting(promptKey & "Description", description)
    Call SaveAPISetting(promptKey & "CreatedDate", CStr(Now))
    
    ' プロンプト一覧にも追加
    Dim promptList As String
    promptList = GetAPISetting("CustomPromptList", "")
    
    If InStr(promptList, promptName) = 0 Then
        If promptList <> "" Then promptList = promptList & ","
        promptList = promptList & promptName
        Call SaveAPISetting("CustomPromptList", promptList)
    End If
    
    LogInfo "カスタムプロンプト保存完了: " & promptName
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムプロンプト保存中にエラーが発生しました", Err.Description
    LogError "SaveCustomPrompt", Err.Description
End Sub

' カスタムプロンプトの存在確認
Private Function IsCustomPromptExists(ByVal promptName As String) As Boolean
    Dim promptList As String
    promptList = GetAPISetting("CustomPromptList", "")
    IsCustomPromptExists = (InStr(promptList, promptName) > 0)
End Function

' カスタムプロンプト一覧表示
Private Sub ShowCustomPromptList()
    On Error GoTo ErrorHandler
    
    Dim promptList As String
    promptList = GetAPISetting("CustomPromptList", "")
    
    If promptList = "" Then
        ShowWarning "保存されているカスタムプロンプトはありません。"
        Exit Sub
    End If
    
    Dim promptNames() As String
    promptNames = Split(promptList, ",")
    
    Dim listText As String
    listText = "【カスタムプロンプト一覧】" & vbCrLf & vbCrLf
    
    Dim i As Integer
    For i = 0 To UBound(promptNames)
        If Trim(promptNames(i)) <> "" Then
            Dim promptName As String
            promptName = Trim(promptNames(i))
            
            Dim category As String
            Dim description As String
            Dim createdDate As String
            
            category = GetAPISetting("CustomPrompts\" & promptName & "\Category", "不明")
            description = GetAPISetting("CustomPrompts\" & promptName & "\Description", "")
            createdDate = GetAPISetting("CustomPrompts\" & promptName & "\CreatedDate", "")
            
            listText = listText & (i + 1) & ". " & promptName & vbCrLf & _
                      "   カテゴリ: " & category & vbCrLf & _
                      "   説明: " & IIf(description = "", "なし", description) & vbCrLf & _
                      "   作成日: " & createdDate & vbCrLf & vbCrLf
        End If
    Next i
    
    Call DisplayLongText("カスタムプロンプト一覧", listText)
    
    Exit Sub
    
ErrorHandler:
    ShowError "カスタムプロンプト一覧表示中にエラーが発生しました", Err.Description
    LogError "ShowCustomPromptList", Err.Description
End Sub

' =============================================================================
' プロンプト効果測定
' =============================================================================

' プロンプト効果測定
Private Sub MeasurePromptEffectiveness()
    ShowWarning "プロンプト効果測定機能は将来のバージョンで実装予定です。" & vbCrLf & vbCrLf & _
               "予定機能:" & vbCrLf & _
               "- プロンプト実行時間の計測" & vbCrLf & _
               "- 結果品質の評価" & vbCrLf & _
               "- プロンプト改善提案" & vbCrLf & _
               "- A/Bテスト機能"
    LogInfo "MeasurePromptEffectiveness: 未実装機能が呼び出されました"
End Sub

' =============================================================================
' プロンプトのインポート/エクスポート
' =============================================================================

' プロンプトのインポート/エクスポート
Private Sub ImportExportPrompts()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    choice = InputBox("プロンプトのインポート/エクスポート" & vbCrLf & vbCrLf & _
                     "1. プロンプトのエクスポート（テキストファイル）" & vbCrLf & _
                     "2. プロンプトのインポート（テキストファイル）" & vbCrLf & _
                     "3. Excelファイルへのエクスポート" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     "インポート/エクスポート")
    
    Select Case choice
        Case "1"
            Call ExportPromptsToText
        Case "2"
            Call ImportPromptsFromText
        Case "3"
            Call ExportPromptsToExcel
        Case ""
            ' キャンセル
        Case Else
            ShowWarning "無効な選択です。1-3の番号を入力してください。"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "インポート/エクスポート中にエラーが発生しました", Err.Description
    LogError "ImportExportPrompts", Err.Description
End Sub

' プロンプトのテキストファイルエクスポート
Private Sub ExportPromptsToText()
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    fileName = Environ("USERPROFILE") & "\Desktop\ExcelAI_Prompts_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open fileName For Output As #fileNum
    
    Print #fileNum, "Excel AI Helper - プロンプトエクスポート"
    Print #fileNum, "作成日時: " & Now
    Print #fileNum, "=========================================="
    Print #fileNum, ""
    
    ' プリセットプロンプトのエクスポート
    Print #fileNum, "[プリセットプロンプト]"
    Print #fileNum, ""
    Print #fileNum, "学習意欲判定:"
    Print #fileNum, GetPromptTemplate("LEARNING_MOTIVATION")
    Print #fileNum, ""
    Print #fileNum, "感情分析:"
    Print #fileNum, GetPromptTemplate("SENTIMENT_ANALYSIS")
    Print #fileNum, ""
    
    ' カスタムプロンプトのエクスポート
    Print #fileNum, "[カスタムプロンプト]"
    Print #fileNum, ""
    
    Dim promptList As String
    promptList = GetAPISetting("CustomPromptList", "")
    
    If promptList <> "" Then
        Dim promptNames() As String
        promptNames = Split(promptList, ",")
        
        Dim i As Integer
        For i = 0 To UBound(promptNames)
            If Trim(promptNames(i)) <> "" Then
                Dim promptName As String
                promptName = Trim(promptNames(i))
                
                Print #fileNum, "名前: " & promptName
                Print #fileNum, "カテゴリ: " & GetAPISetting("CustomPrompts\" & promptName & "\Category", "")
                Print #fileNum, "内容: " & GetAPISetting("CustomPrompts\" & promptName & "\Content", "")
                Print #fileNum, "説明: " & GetAPISetting("CustomPrompts\" & promptName & "\Description", "")
                Print #fileNum, ""
            End If
        Next i
    End If
    
    Close #fileNum
    
    ShowSuccess "プロンプトをエクスポートしました。" & vbCrLf & vbCrLf & "保存先: " & fileName
    LogInfo "プロンプトエクスポート完了: " & fileName
    
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    ShowError "プロンプトエクスポート中にエラーが発生しました", Err.Description
    LogError "ExportPromptsToText", Err.Description
End Sub

' =============================================================================
' ユーティリティ関数
' =============================================================================

' 長いテキストの分割表示
Private Sub DisplayLongText(ByVal title As String, ByVal text As String)
    Const maxLength As Integer = 1000
    
    If Len(text) <= maxLength Then
        MsgBox text, vbInformation, title
    Else
        Dim currentPos As Integer
        Dim pageNum As Integer
        Dim totalPages As Integer
        
        currentPos = 1
        pageNum = 1
        totalPages = Int((Len(text) - 1) / maxLength) + 1
        
        While currentPos <= Len(text)
            Dim pageText As String
            pageText = Mid(text, currentPos, maxLength)
            
            MsgBox pageText, vbInformation, title & " (" & pageNum & "/" & totalPages & ")"
            
            currentPos = currentPos + maxLength
            pageNum = pageNum + 1
        Wend
    End If
End Sub

' プロンプトテンプレート編集
Private Sub EditPromptTemplates()
    ShowWarning "プロンプトテンプレート編集機能は設定管理から利用してください。"
End Sub

' カスタムプロンプト編集
Private Sub EditCustomPrompt()
    ShowWarning "カスタムプロンプト編集機能は将来のバージョンで実装予定です。"
End Sub

' カスタムプロンプト削除
Private Sub DeleteCustomPrompt()
    ShowWarning "カスタムプロンプト削除機能は将来のバージョンで実装予定です。"
End Sub

' カスタムプロンプト使用
Private Sub UseCustomPrompt()
    ShowWarning "カスタムプロンプト使用機能は将来のバージョンで実装予定です。"
End Sub

' テキストファイルからのインポート
Private Sub ImportPromptsFromText()
    ShowWarning "プロンプトインポート機能は将来のバージョンで実装予定です。"
End Sub

' Excelファイルへのエクスポート
Private Sub ExportPromptsToExcel()
    ShowWarning "Excelファイルエクスポート機能は将来のバージョンで実装予定です。"
End Sub