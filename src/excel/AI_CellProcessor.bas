' AI_CellProcessor.bas
' Excel AI Helper - セル単位処理モジュール
' 作成日: 2024
' 説明: 選択したセルの内容を基にAI関数やデータ変換を生成する機能

Option Explicit

' =============================================================================
' セル処理メイン関数
' =============================================================================

' セルのAI関数生成メイン関数（リボンから呼び出される）
Public Sub GenerateAIFunction()
    On Error GoTo ErrorHandler
    
    LogInfo "セルAI関数生成処理を開始"
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then
        Exit Sub
    End If
    
    ' 設定確認
    If Not ValidateConfiguration() Then
        Exit Sub
    End If
    
    ' プロンプト編集ダイアログ表示
    Dim userPrompt As String
    Dim generationType As String
    Dim targetCell As Range
    
    If Not ShowPromptEditDialog(userPrompt, generationType, targetCell) Then
        LogInfo "ユーザーがキャンセルしました"
        Exit Sub
    End If
    
    ' AI関数生成実行
    Call ProcessCellGeneration(targetCell, userPrompt, generationType)
    
    LogInfo "セルAI関数生成処理を完了"
    
    Exit Sub
    
ErrorHandler:
    ClearProgress
    ShowError "セルAI関数生成中にエラーが発生しました", Err.Description
    LogError "GenerateAIFunction", Err.Description
End Sub

' =============================================================================
' プロンプト編集ダイアログ
' =============================================================================

' プロンプト編集ダイアログの表示
Private Function ShowPromptEditDialog(ByRef userPrompt As String, ByRef generationType As String, ByRef targetCell As Range) As Boolean
    On Error GoTo ErrorHandler
    
    ' 選択範囲の内容を取得
    Dim selectionContent As String
    selectionContent = GetSelectionContent(Selection)
    
    ' 生成タイプの選択
    Dim choice As String
    choice = InputBox("生成する内容を選択してください:" & vbCrLf & vbCrLf & _
                     "1. Excel関数生成" & vbCrLf & _
                     "2. データ変換" & vbCrLf & _
                     "3. 計算式生成" & vbCrLf & _
                     "4. データ補完" & vbCrLf & _
                     "5. カスタム生成" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     "生成タイプ選択")
    
    If choice = "" Then
        ShowPromptEditDialog = False
        Exit Function
    End If
    
    ' プリセットプロンプトの設定
    Dim basePrompt As String
    Select Case choice
        Case "1"
            generationType = "excel_function"
            basePrompt = "以下のデータまたは要求を基に、適切なExcel関数を生成してください。関数のみを回答してください。"
        Case "2"
            generationType = "data_transformation"
            basePrompt = "以下のデータを指定された形式に変換してください。変換後のデータのみを回答してください。"
        Case "3"
            generationType = "calculation_formula"
            basePrompt = "以下の要求に基づいて、適切な計算式を生成してください。式のみを回答してください。"
        Case "4"
            generationType = "data_completion"
            basePrompt = "以下のデータの欠損値や不足部分を推定して補完してください。補完値のみを回答してください。"
        Case "5"
            generationType = "custom"
            basePrompt = InputBox("カスタム生成のベースプロンプトを入力してください:", "カスタムプロンプト")
            If basePrompt = "" Then
                ShowPromptEditDialog = False
                Exit Function
            End If
        Case Else
            ShowWarning "無効な選択です。"
            ShowPromptEditDialog = False
            Exit Function
    End Select
    
    ' プロンプト編集ダイアログ
    Dim promptDialog As String
    promptDialog = "プロンプトを編集してください:" & vbCrLf & vbCrLf & _
                  "【選択範囲の内容】" & vbCrLf & _
                  Left(selectionContent, 200) & IIf(Len(selectionContent) > 200, "...", "") & vbCrLf & vbCrLf & _
                  "【ベースプロンプト】" & vbCrLf & basePrompt & vbCrLf & vbCrLf & _
                  "追加の指示があれば入力してください（空欄可）:"
    
    Dim additionalPrompt As String
    additionalPrompt = InputBox(promptDialog, "プロンプト編集")
    
    ' 最終プロンプトの構築
    userPrompt = basePrompt
    If additionalPrompt <> "" Then
        userPrompt = userPrompt & vbCrLf & vbCrLf & "追加指示: " & additionalPrompt
    End If
    
    ' 出力先セルの指定
    Dim outputChoice As String
    outputChoice = InputBox("結果を出力するセルを指定してください:" & vbCrLf & vbCrLf & _
                           "1. 選択範囲内の最初のセル" & vbCrLf & _
                           "2. 選択範囲の右隣のセル" & vbCrLf & _
                           "3. 指定のセル（例：D5）" & vbCrLf & vbCrLf & _
                           "番号またはセル名を入力してください:", _
                           "出力先指定", "1")
    
    If outputChoice = "" Then
        ShowPromptEditDialog = False
        Exit Function
    End If
    
    ' 出力先セルの設定
    Select Case outputChoice
        Case "1"
            Set targetCell = Selection.Cells(1, 1)
        Case "2"
            Set targetCell = Selection.Cells(1, Selection.Columns.Count + 1)
        Case "3"
            Dim customCell As String
            customCell = InputBox("出力先のセル名を入力してください（例：D5）:", "セル指定")
            If customCell = "" Then
                ShowPromptEditDialog = False
                Exit Function
            End If
            Set targetCell = Range(customCell)
        Case Else
            ' セル名として直接指定された場合
            On Error Resume Next
            Set targetCell = Range(outputChoice)
            On Error GoTo ErrorHandler
            If targetCell Is Nothing Then
                ShowWarning "無効なセル指定です。"
                ShowPromptEditDialog = False
                Exit Function
            End If
    End Select
    
    ' 最終確認
    Dim confirmMsg As String
    confirmMsg = "以下の設定で実行しますか？" & vbCrLf & vbCrLf & _
                "生成タイプ: " & generationType & vbCrLf & _
                "出力先: " & targetCell.Address & vbCrLf & _
                "選択範囲: " & Selection.Address & vbCrLf & vbCrLf & _
                "プロンプト: " & Left(userPrompt, 100) & "..."
    
    If ShowConfirm(confirmMsg) Then
        ShowPromptEditDialog = True
    Else
        ShowPromptEditDialog = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "プロンプト編集中にエラーが発生しました", Err.Description
    ShowPromptEditDialog = False
End Function

' =============================================================================
' セル生成処理メイン関数
' =============================================================================

' セル生成処理実行
Private Sub ProcessCellGeneration(ByVal targetCell As Range, ByVal userPrompt As String, ByVal generationType As String)
    On Error GoTo ErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    ' 処理開始メッセージ
    ShowProgress 0, 1, "AI関数生成を実行しています..."
    
    ' 選択範囲の内容を取得
    Dim selectionContent As String
    selectionContent = GetSelectionContent(Selection)
    
    ' システムプロンプトの取得
    Dim systemPrompt As String
    systemPrompt = GetAPISetting("CellProcessingTemplate", SYSTEM_PROMPT_CELL_PROCESSOR)
    
    ' 完全なユーザープロンプトの構築
    Dim fullUserPrompt As String
    fullUserPrompt = userPrompt & vbCrLf & vbCrLf & _
                    "【対象データ】" & vbCrLf & selectionContent
    
    ' 文字数制限チェック
    If Len(fullUserPrompt) > MAX_CONTENT_LENGTH Then
        ShowError "データが大きすぎます", "選択範囲のデータサイズを小さくしてください。" & vbCrLf & _
                  "現在のサイズ: " & Len(fullUserPrompt) & " 文字" & vbCrLf & _
                  "上限: " & MAX_CONTENT_LENGTH & " 文字"
        ClearProgress
        Exit Sub
    End If
    
    ' AI生成実行
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, fullUserPrompt)
    
    ' 結果の検証と後処理
    If result <> "" Then
        result = ProcessGenerationResult(result, generationType)
        
        ' プレビューと確認
        If ShowResultPreview(result, targetCell, generationType) Then
            ' 結果をセルに出力
            Call WriteResultToCell(targetCell, result, generationType)
            
            ' 成功メッセージ
            Dim processingTime As Long
            processingTime = DateDiff("s", startTime, Now)
            
            ShowSuccess "AI関数生成が完了しました。" & vbCrLf & vbCrLf & _
                       "処理時間: " & processingTime & "秒" & vbCrLf & _
                       "出力先: " & targetCell.Address
            
            LogInfo "セル生成成功: " & targetCell.Address & " = " & Left(result, 50)
        Else
            LogInfo "ユーザーが結果適用をキャンセル"
        End If
    Else
        ShowError "AI関数生成に失敗しました", "プロンプトや設定を確認してください。"
        LogError "ProcessCellGeneration", "AI生成結果が空"
    End If
    
    ClearProgress
    
    Exit Sub
    
ErrorHandler:
    ClearProgress
    ShowError "セル生成処理中にエラーが発生しました", Err.Description
    LogError "ProcessCellGeneration", Err.Description
End Sub

' =============================================================================
' ユーティリティ関数
' =============================================================================

' 選択範囲の内容を文字列として取得
Private Function GetSelectionContent(ByVal targetRange As Range) As String
    Dim content As String
    Dim cell As Range
    
    content = ""
    
    ' 範囲が単一セルの場合
    If targetRange.Cells.Count = 1 Then
        content = CStr(targetRange.Value)
    Else
        ' 複数セルの場合は構造化して出力
        content = "【セル範囲: " & targetRange.Address & "】" & vbCrLf
        
        Dim row As Integer
        Dim col As Integer
        For row = 1 To targetRange.Rows.Count
            For col = 1 To targetRange.Columns.Count
                Set cell = targetRange.Cells(row, col)
                If Trim(CStr(cell.Value)) <> "" Then
                    content = content & cell.Address(False, False) & ": " & CStr(cell.Value) & vbCrLf
                End If
            Next col
        Next row
    End If
    
    GetSelectionContent = content
End Function

' 生成結果の後処理
Private Function ProcessGenerationResult(ByVal result As String, ByVal generationType As String) As String
    On Error GoTo ErrorHandler
    
    result = Trim(result)
    
    Select Case generationType
        Case "excel_function"
            ' Excel関数の場合、=で始まることを確認
            If Left(result, 1) <> "=" And InStr(result, "=") = 0 Then
                result = "=" & result
            End If
            ' 複数行の場合は最初の行のみを使用
            If InStr(result, vbCrLf) > 0 Then
                result = Split(result, vbCrLf)(0)
            End If
            
        Case "calculation_formula"
            ' 計算式の場合も=で始まることを確認
            If Left(result, 1) <> "=" And IsNumeric(Left(result, 1)) Then
                result = "=" & result
            End If
            
        Case "data_transformation", "data_completion"
            ' データ変換・補完の場合はそのまま
            ' 改行が含まれている場合は最初の有効な値のみ使用
            Dim lines() As String
            lines = Split(result, vbCrLf)
            Dim i As Integer
            For i = 0 To UBound(lines)
                If Trim(lines(i)) <> "" Then
                    result = Trim(lines(i))
                    Exit For
                End If
            Next i
            
        Case "custom"
            ' カスタムの場合は特別な処理なし
    End Select
    
    ProcessGenerationResult = result
    
    Exit Function
    
ErrorHandler:
    ProcessGenerationResult = result ' エラーの場合は元の結果を返す
End Function

' 結果プレビューと確認
Private Function ShowResultPreview(ByVal result As String, ByVal targetCell As Range, ByVal generationType As String) As Boolean
    Dim previewMsg As String
    previewMsg = "生成された結果をプレビューしてください:" & vbCrLf & vbCrLf & _
                "【生成タイプ】" & generationType & vbCrLf & _
                "【出力先】" & targetCell.Address & vbCrLf & vbCrLf & _
                "【生成結果】" & vbCrLf & result & vbCrLf & vbCrLf & _
                "この結果をセルに適用しますか？"
    
    ShowResultPreview = ShowConfirm(previewMsg)
End Function

' 結果をセルに出力
Private Sub WriteResultToCell(ByVal targetCell As Range, ByVal result As String, ByVal generationType As String)
    On Error GoTo ErrorHandler
    
    ' 元の値をバックアップ（Undo機能用）
    Dim originalValue As Variant
    originalValue = targetCell.Value
    
    ' 結果の適用
    Select Case generationType
        Case "excel_function", "calculation_formula"
            ' 関数・計算式の場合はFormulaとして設定
            targetCell.Formula = result
        Case Else
            ' その他の場合は値として設定
            targetCell.Value = result
    End Select
    
    ' セルの書式設定
    If generationType = "excel_function" Or generationType = "calculation_formula" Then
        ' 関数セルは薄い緑の背景
        targetCell.Interior.Color = RGB(230, 255, 230)
    Else
        ' データセルは薄い青の背景
        targetCell.Interior.Color = RGB(230, 240, 255)
    End If
    
    ' フォントを少し太字に
    targetCell.Font.Bold = True
    
    LogInfo "セルに結果を出力: " & targetCell.Address & " = " & Left(result, 100)
    
    Exit Sub
    
ErrorHandler:
    ShowError "セルへの結果出力中にエラーが発生しました", Err.Description
    LogError "WriteResultToCell", Err.Description
End Sub

' =============================================================================
' 高度な機能（将来拡張）
' =============================================================================

' 複数セル一括生成
Public Sub BatchGenerateAIFunctions()
    ShowWarning "複数セル一括生成機能は将来のバージョンで実装予定です。"
    LogInfo "BatchGenerateAIFunctions: 未実装機能が呼び出されました"
End Sub

' 生成履歴管理
Public Sub ShowGenerationHistory()
    ShowWarning "生成履歴管理機能は将来のバージョンで実装予定です。"
    LogInfo "ShowGenerationHistory: 未実装機能が呼び出されました"
End Sub

' セル内容のAI分析
Public Sub AnalyzeCellContent()
    On Error GoTo ErrorHandler
    
    LogInfo "セル内容分析を開始"
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then
        Exit Sub
    End If
    
    ' 設定確認
    If Not ValidateConfiguration() Then
        Exit Sub
    End If
    
    ' 分析実行
    Dim content As String
    content = GetSelectionContent(Selection)
    
    If Trim(content) = "" Then
        ShowWarning "分析対象のデータが選択されていません。"
        Exit Sub
    End If
    
    ShowProgress 0, 1, "セル内容を分析中..."
    
    Dim systemPrompt As String
    systemPrompt = "あなたはExcelデータの分析専門家です。提供されたセルデータを分析し、内容の特徴、データ型、潜在的な問題点、改善提案を日本語で回答してください。"
    
    Dim userPrompt As String
    userPrompt = "以下のExcelセルデータを分析してください:" & vbCrLf & vbCrLf & content
    
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userPrompt)
    
    ClearProgress
    
    If result <> "" Then
        ' 結果表示用のフォーム（簡易版はMsgBox）
        MsgBox "【セル内容分析結果】" & vbCrLf & vbCrLf & result, vbInformation, "セル分析"
        LogInfo "セル内容分析完了"
    Else
        ShowError "セル内容分析に失敗しました", "APIの設定や接続を確認してください。"
    End If
    
    Exit Sub
    
ErrorHandler:
    ClearProgress
    ShowError "セル内容分析中にエラーが発生しました", Err.Description
    LogError "AnalyzeCellContent", Err.Description
End Sub