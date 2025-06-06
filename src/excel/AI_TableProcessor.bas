' AI_TableProcessor.bas
' Excel AI Helper - 表の行単位処理モジュール
' 作成日: 2024
' 説明: 選択した表の各行をAIで分析し、結果を隣接列に出力する機能

Option Explicit

' =============================================================================
' 表処理メイン関数
' =============================================================================

' 表の行分析メイン関数（リボンから呼び出される）
Public Sub AnalyzeTableRows()
    On Error GoTo ErrorHandler
    
    LogInfo "表の行分析処理を開始"
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then
        Exit Sub
    End If
    
    ' 設定確認
    If Not ValidateConfiguration() Then
        Exit Sub
    End If
    
    ' 分析設定ダイアログ表示
    Dim analysisType As String
    Dim outputColumn As Integer
    Dim customPrompt As String
    
    If Not ShowAnalysisSettingsDialog(analysisType, outputColumn, customPrompt) Then
        LogInfo "ユーザーがキャンセルしました"
        Exit Sub
    End If
    
    ' 表データの処理実行
    Call ProcessTableData(Selection, analysisType, outputColumn, customPrompt)
    
    LogInfo "表の行分析処理を完了"
    
    Exit Sub
    
ErrorHandler:
    ClearProgress
    ShowError "表の行分析中にエラーが発生しました", Err.Description
    LogError "AnalyzeTableRows", Err.Description
End Sub

' =============================================================================
' 分析設定ダイアログ
' =============================================================================

' 分析設定ダイアログの表示
Private Function ShowAnalysisSettingsDialog(ByRef analysisType As String, ByRef outputColumn As Integer, ByRef customPrompt As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 分析タイプの選択
    Dim choice As String
    choice = InputBox("実行する分析を選択してください:" & vbCrLf & vbCrLf & _
                     "1. 学習意欲判定（0-10スコア）" & vbCrLf & _
                     "2. 感情分析（ポジティブ/ネガティブ/中性）" & vbCrLf & _
                     "3. カテゴリ分類" & vbCrLf & _
                     "4. 要約生成" & vbCrLf & _
                     "5. カスタム分析" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     "分析タイプ選択")
    
    If choice = "" Then
        ShowAnalysisSettingsDialog = False
        Exit Function
    End If
    
    ' 分析タイプの設定
    Select Case choice
        Case "1"
            analysisType = "learning_motivation"
            customPrompt = "以下の行データを分析し、学習意欲の高さを0-10のスコアで評価してください。スコアのみを数値で回答してください。"
        Case "2"
            analysisType = "sentiment_analysis"
            customPrompt = "以下の行データを分析し、感情を「ポジティブ」「ネガティブ」「中性」のいずれかで判定してください。判定結果のみを回答してください。"
        Case "3"
            analysisType = "category_classification"
            customPrompt = "以下の行データを分析し、最も適切なカテゴリを1つ選んで回答してください。カテゴリ名のみを回答してください。"
        Case "4"
            analysisType = "summarization"
            customPrompt = "以下の行データを分析し、50文字以内で要約してください。要約のみを回答してください。"
        Case "5"
            analysisType = "custom"
            customPrompt = InputBox("カスタム分析のプロンプトを入力してください:" & vbCrLf & vbCrLf & _
                                   "※ 「以下の行データを分析し、」で始まるようにしてください。", _
                                   "カスタムプロンプト")
            If customPrompt = "" Then
                ShowAnalysisSettingsDialog = False
                Exit Function
            End If
        Case Else
            ShowWarning "無効な選択です。"
            ShowAnalysisSettingsDialog = False
            Exit Function
    End Select
    
    ' 出力列の指定
    Dim outputColumnInput As String
    outputColumnInput = InputBox("結果を出力する列を指定してください:" & vbCrLf & vbCrLf & _
                                "例: D（D列に出力）" & vbCrLf & _
                                "例: 4（4列目に出力）" & vbCrLf & vbCrLf & _
                                "空欄の場合は選択範囲の右隣の列に出力されます。", _
                                "出力列指定")
    
    ' 出力列の解析
    If outputColumnInput = "" Then
        ' 選択範囲の右隣の列
        outputColumn = Selection.Column + Selection.Columns.Count
    ElseIf IsNumeric(outputColumnInput) Then
        ' 数値で指定された場合
        outputColumn = CInt(outputColumnInput)
    Else
        ' 列文字で指定された場合（A, B, C...)
        outputColumn = Range(outputColumnInput & "1").Column
    End If
    
    ' 設定確認
    Dim confirmMsg As String
    confirmMsg = "以下の設定で実行しますか？" & vbCrLf & vbCrLf & _
                "分析タイプ: " & analysisType & vbCrLf & _
                "出力列: " & Split(Cells(1, outputColumn).Address, "$")(1) & "列" & vbCrLf & _
                "対象行数: " & Selection.Rows.Count & "行" & vbCrLf & vbCrLf & _
                "プロンプト: " & Left(customPrompt, 100) & "..."
    
    If ShowConfirm(confirmMsg) Then
        ShowAnalysisSettingsDialog = True
    Else
        ShowAnalysisSettingsDialog = False
    End If
    
    Exit Function
    
ErrorHandler:
    ShowError "分析設定中にエラーが発生しました", Err.Description
    ShowAnalysisSettingsDialog = False
End Function

' =============================================================================
' 表データ処理メイン関数
' =============================================================================

' 表データの処理実行
Private Sub ProcessTableData(ByVal targetRange As Range, ByVal analysisType As String, ByVal outputColumn As Integer, ByVal customPrompt As String)
    On Error GoTo ErrorHandler
    
    Dim startTime As Date
    startTime = Now
    
    ' 処理開始メッセージ
    ShowProgress 0, targetRange.Rows.Count, "表の行分析を開始しています..."
    
    ' ヘッダー行の取得
    Dim headerRow As Range
    Set headerRow = targetRange.Rows(1)
    
    ' ヘッダー列名の取得
    Dim headers As String
    headers = GetHeaderString(headerRow)
    
    ' 結果格納用配列
    Dim results() As String
    ReDim results(1 To targetRange.Rows.Count - 1) ' ヘッダー行を除く
    
    ' データ行の処理
    Dim i As Integer
    Dim currentRow As Range
    Dim rowData As String
    Dim analysisResult As String
    Dim successCount As Integer
    Dim errorCount As Integer
    
    successCount = 0
    errorCount = 0
    
    ' 各行を順次処理
    For i = 2 To targetRange.Rows.Count ' ヘッダー行をスキップ
        Set currentRow = targetRange.Rows(i)
        
        ' 進捗表示更新
        ShowProgress i - 1, targetRange.Rows.Count - 1, "行 " & (i - 1) & " を分析中..."
        
        ' 行データの取得
        rowData = GetRowDataString(currentRow, headers)
        
        ' 空行のスキップ
        If Trim(rowData) = "" Then
            results(i - 1) = "[空行]"
            GoTo NextIteration
        End If
        
        ' AIによる分析実行
        analysisResult = AnalyzeRowData(rowData, customPrompt, analysisType)
        
        If analysisResult <> "" Then
            results(i - 1) = analysisResult
            successCount = successCount + 1
            LogInfo "行 " & (i - 1) & " 分析成功: " & Left(analysisResult, 50)
        Else
            results(i - 1) = "[エラー]"
            errorCount = errorCount + 1
            LogError "ProcessTableData", "行 " & (i - 1) & " の分析に失敗"
        End If
        
        ' 処理の一時停止（API制限対応）
        If i Mod 10 = 0 Then
            Application.Wait Now + TimeValue("00:00:01") ' 1秒待機
        End If
        
NextIteration:
    Next i
    
    ' 結果をワークシートに出力
    Call WriteResultsToWorksheet(targetRange, outputColumn, results, analysisType)
    
    ' 処理完了メッセージ
    Dim processingTime As Long
    processingTime = DateDiff("s", startTime, Now)
    
    Dim completeMsg As String
    completeMsg = "表の行分析が完了しました。" & vbCrLf & vbCrLf & _
                 "処理時間: " & processingTime & "秒" & vbCrLf & _
                 "成功: " & successCount & "行" & vbCrLf & _
                 "エラー: " & errorCount & "行" & vbCrLf & _
                 "出力列: " & Split(Cells(1, outputColumn).Address, "$")(1) & "列"
    
    ShowSuccess completeMsg
    
    ClearProgress
    LogInfo "表の行分析完了 - 成功:" & successCount & " エラー:" & errorCount
    
    Exit Sub
    
ErrorHandler:
    ClearProgress
    ShowError "表データ処理中にエラーが発生しました", Err.Description
    LogError "ProcessTableData", Err.Description
End Sub

' =============================================================================
' データ処理ユーティリティ関数
' =============================================================================

' ヘッダー行の文字列化
Private Function GetHeaderString(ByVal headerRow As Range) As String
    Dim headers As String
    Dim cell As Range
    
    headers = ""
    For Each cell In headerRow.Cells
        If Trim(cell.Value) <> "" Then
            If headers <> "" Then headers = headers & ", "
            headers = headers & Trim(cell.Value)
        End If
    Next cell
    
    GetHeaderString = headers
End Function

' 行データの文字列化
Private Function GetRowDataString(ByVal dataRow As Range, ByVal headers As String) As String
    Dim rowData As String
    Dim cell As Range
    Dim cellIndex As Integer
    Dim headerArray() As String
    
    ' ヘッダー配列の作成
    headerArray = Split(headers, ", ")
    
    rowData = ""
    cellIndex = 0
    
    For Each cell In dataRow.Cells
        If cellIndex <= UBound(headerArray) Then
            If Trim(cell.Value) <> "" Then
                If rowData <> "" Then rowData = rowData & " | "
                rowData = rowData & headerArray(cellIndex) & ": " & Trim(cell.Value)
            End If
        End If
        cellIndex = cellIndex + 1
    Next cell
    
    GetRowDataString = rowData
End Function

' 単一行データの分析
Private Function AnalyzeRowData(ByVal rowData As String, ByVal customPrompt As String, ByVal analysisType As String) As String
    On Error GoTo ErrorHandler
    
    ' システムプロンプトの構築
    Dim systemPrompt As String
    systemPrompt = GetAPISetting("TableAnalysisTemplate", SYSTEM_PROMPT_TABLE_ANALYZER)
    
    ' ユーザープロンプトの構築
    Dim userPrompt As String
    userPrompt = customPrompt & vbCrLf & vbCrLf & _
                "分析対象データ:" & vbCrLf & rowData
    
    ' 文字数制限チェック
    If Len(userPrompt) > MAX_CONTENT_LENGTH Then
        LogError "AnalyzeRowData", "プロンプトが長すぎます: " & Len(userPrompt) & " 文字"
        AnalyzeRowData = "[文字数制限エラー]"
        Exit Function
    End If
    
    ' 他の関数からも呼び出せるように変更
    Dim result As String
    result = SendOpenAIRequest(systemPrompt, userPrompt)
    
    If result <> "" Then
        result = AI_TableProcessor.ValidateAndFormatResult(result, analysisType)
    End If
    
    AnalyzeRowData = result
    
    Exit Function
    
ErrorHandler:
    LogError "AnalyzeRowData", Err.Description
    AnalyzeRowData = "[処理エラー]"
End Function

' 分析結果の検証と整形
Private Function ValidateAndFormatResult(ByVal result As String, ByVal analysisType As String) As String
    On Error GoTo ErrorHandler
    
    Select Case analysisType
        Case "learning_motivation"
            ' 数値スコア（0-10）の検証
            If IsNumeric(result) Then
                Dim score As Double
                score = CDbl(result)
                If score >= 0 And score <= 10 Then
                    ValidateAndFormatResult = Format(score, "0.0")
                Else
                    ValidateAndFormatResult = "[範囲外: " & result & "]"
                End If
            Else
                ' 数値以外の場合、数値を抽出しようとする
                Dim numericPart As String
                numericPart = ExtractNumericValue(result)
                If numericPart <> "" Then
                    ValidateAndFormatResult = numericPart
                Else
                    ValidateAndFormatResult = "[非数値: " & Left(result, 20) & "]"
                End If
            End If
            
        Case "sentiment_analysis"
            ' 感情分析結果の正規化
            result = LCase(Trim(result))
            If InStr(result, "ポジティブ") > 0 Or InStr(result, "positive") > 0 Then
                ValidateAndFormatResult = "ポジティブ"
            ElseIf InStr(result, "ネガティブ") > 0 Or InStr(result, "negative") > 0 Then
                ValidateAndFormatResult = "ネガティブ"
            ElseIf InStr(result, "中性") > 0 Or InStr(result, "neutral") > 0 Then
                ValidateAndFormatResult = "中性"
            Else
                ValidateAndFormatResult = result ' そのまま返す
            End If
            
        Case Else
            ' その他の場合はそのまま返す（最大100文字）
            If Len(result) > 100 Then
                ValidateAndFormatResult = Left(result, 97) & "..."
            Else
                ValidateAndFormatResult = result
            End If
    End Select
    
    Exit Function
    
ErrorHandler:
    ValidateAndFormatResult = result ' エラーの場合は元の結果を返す
End Function

' 文字列から数値を抽出
Private Function ExtractNumericValue(ByVal text As String) As String
    Dim i As Integer
    Dim char As String
    Dim numericText As String
    Dim decimalFound As Boolean
    
    numericText = ""
    decimalFound = False
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If IsNumeric(char) Then
            numericText = numericText & char
        ElseIf char = "." And Not decimalFound Then
            numericText = numericText & char
            decimalFound = True
        ElseIf numericText <> "" Then
            ' 数値文字列が終了した
            Exit For
        End If
    Next i
    
    ' 数値として有効かチェック
    If IsNumeric(numericText) Then
        Dim value As Double
        value = CDbl(numericText)
        If value >= 0 And value <= 10 Then
            ExtractNumericValue = Format(value, "0.0")
        Else
            ExtractNumericValue = ""
        End If
    Else
        ExtractNumericValue = ""
    End If
End Function

' =============================================================================
' 結果出力関数
' =============================================================================

' 結果をワークシートに出力
Private Sub WriteResultsToWorksheet(ByVal targetRange As Range, ByVal outputColumn As Integer, ByRef results() As String, ByVal analysisType As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = targetRange.Worksheet
    
    ' ヘッダーの設定
    Dim headerText As String
    Select Case analysisType
        Case "learning_motivation"
            headerText = "学習意欲スコア"
        Case "sentiment_analysis"
            headerText = "感情分析"
        Case "category_classification"
            headerText = "カテゴリ"
        Case "summarization"
            headerText = "要約"
        Case "custom"
            headerText = "AI分析結果"
        Case Else
            headerText = "分析結果"
    End Select
    
    ' ヘッダー行の設定
    ws.Cells(targetRange.Row, outputColumn).Value = headerText
    ws.Cells(targetRange.Row, outputColumn).Font.Bold = True
    
    ' 結果データの出力
    Dim i As Integer
    For i = 1 To UBound(results)
        ws.Cells(targetRange.Row + i, outputColumn).Value = results(i)
        
        ' エラー行のハイライト
        If InStr(results(i), "[エラー]") > 0 Or InStr(results(i), "[範囲外") > 0 Then
            ws.Cells(targetRange.Row + i, outputColumn).Interior.Color = RGB(255, 200, 200) ' 薄い赤
        ElseIf results(i) = "[空行]" Then
            ws.Cells(targetRange.Row + i, outputColumn).Interior.Color = RGB(240, 240, 240) ' 薄いグレー
        End If
    Next i
    
    ' 列幅の自動調整
    ws.Columns(outputColumn).AutoFit
    
    LogInfo "結果をワークシートに出力完了: " & UBound(results) & "行"
    
    Exit Sub
    
ErrorHandler:
    ShowError "結果出力中にエラーが発生しました", Err.Description
    LogError "WriteResultsToWorksheet", Err.Description
End Sub