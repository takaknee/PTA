' AI_RibbonCallbacks.bas
' Excel AI Helper - リボンイベント処理モジュール
' 作成日: 2024
' 説明: Excelリボンからの各種コールバック関数を処理

Option Explicit

' =============================================================================
' リボンメイン関数
' =============================================================================

' メインメニュー表示（リボンから呼び出される）
Public Sub ShowMainMenu()
    On Error GoTo ErrorHandler
    
    LogInfo "メインメニューを表示"
    
    Dim choice As String
    Dim menuText As String
    
    menuText = APP_NAME & " - メインメニュー" & vbCrLf & vbCrLf & _
               "■ 主要機能" & vbCrLf & _
               "1. 表の行分析（学習意欲、感情分析など）" & vbCrLf & _
               "2. セルAI関数生成" & vbCrLf & _
               "3. セル内容の分析" & vbCrLf & vbCrLf & _
               "■ 管理機能" & vbCrLf & _
               "4. プロンプト管理" & vbCrLf & _
               "5. API設定" & vbCrLf & _
               "6. システム情報" & vbCrLf & vbCrLf & _
               "実行したい機能の番号を入力してください:"
    
    choice = InputBox(menuText, APP_NAME)
    
    Select Case choice
        Case "1"
            Call AI_TableProcessor.AnalyzeTableRows
        Case "2"
            Call AI_CellProcessor.GenerateAIFunction
        Case "3"
            Call AI_CellProcessor.AnalyzeCellContent
        Case "4"
            Call AI_PromptManager.ShowPromptManager
        Case "5"
            Call AI_ConfigManager.ShowConfigurationMenu
        Case "6"
            Call ShowSystemInformation
        Case ""
            ' キャンセル
        Case Else
            ShowWarning "無効な選択です。1-6の番号を入力してください。"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "メニュー処理中にエラーが発生しました", Err.Description
    LogError "ShowMainMenu", Err.Description
End Sub

' =============================================================================
' 表処理関連のリボンコールバック
' =============================================================================

' 表の行分析（直接呼び出し）
Public Sub OnTableAnalysisClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "リボンから表分析が実行されました"
    Call AI_TableProcessor.AnalyzeTableRows
    
    Exit Sub
    
ErrorHandler:
    ShowError "表分析実行中にエラーが発生しました", Err.Description
    LogError "OnTableAnalysisClick", Err.Description
End Sub

' 学習意欲判定（プリセット）
Public Sub OnLearningMotivationClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "学習意欲判定が実行されました"
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then Exit Sub
    If Not ValidateConfiguration() Then Exit Sub
    
    ' 学習意欲判定プリセットで実行
    Call ExecutePresetTableAnalysis("learning_motivation")
    
    Exit Sub
    
ErrorHandler:
    ShowError "学習意欲判定実行中にエラーが発生しました", Err.Description
    LogError "OnLearningMotivationClick", Err.Description
End Sub

' 感情分析（プリセット）
Public Sub OnSentimentAnalysisClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "感情分析が実行されました"
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then Exit Sub
    If Not ValidateConfiguration() Then Exit Sub
    
    ' 感情分析プリセットで実行
    Call ExecutePresetTableAnalysis("sentiment_analysis")
    
    Exit Sub
    
ErrorHandler:
    ShowError "感情分析実行中にエラーが発生しました", Err.Description
    LogError "OnSentimentAnalysisClick", Err.Description
End Sub

' プリセット表分析の実行
Private Sub ExecutePresetTableAnalysis(ByVal analysisType As String)
    On Error GoTo ErrorHandler
    
    ' 出力列の自動設定（選択範囲の右隣）
    Dim outputColumn As Integer
    outputColumn = Selection.Column + Selection.Columns.Count
    
    ' プリセットプロンプトの取得
    Dim customPrompt As String
    Select Case analysisType
        Case "learning_motivation"
            customPrompt = "以下の行データを分析し、学習意欲の高さを0-10のスコアで評価してください。スコアのみを数値で回答してください。"
        Case "sentiment_analysis"
            customPrompt = "以下の行データを分析し、感情を「ポジティブ」「ネガティブ」「中性」のいずれかで判定してください。判定結果のみを回答してください。"
        Case Else
            customPrompt = "以下の行データを分析してください。"
    End Select
    
    ' 確認ダイアログ
    Dim confirmMsg As String
    confirmMsg = "プリセット分析を実行しますか？" & vbCrLf & vbCrLf & _
                "分析タイプ: " & analysisType & vbCrLf & _
                "対象範囲: " & Selection.Address & vbCrLf & _
                "出力列: " & Split(Cells(1, outputColumn).Address, "$")(1) & "列"
    
    If ShowConfirm(confirmMsg) Then
        Call AI_TableProcessor.ProcessTableData(Selection, analysisType, outputColumn, customPrompt)
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "プリセット表分析実行中にエラーが発生しました", Err.Description
    LogError "ExecutePresetTableAnalysis", Err.Description
End Sub

' =============================================================================
' セル処理関連のリボンコールバック
' =============================================================================

' セルAI関数生成（直接呼び出し）
Public Sub OnCellFunctionClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "リボンからセル関数生成が実行されました"
    Call AI_CellProcessor.GenerateAIFunction
    
    Exit Sub
    
ErrorHandler:
    ShowError "セル関数生成実行中にエラーが発生しました", Err.Description
    LogError "OnCellFunctionClick", Err.Description
End Sub

' Excel関数生成（プリセット）
Public Sub OnExcelFunctionClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "Excel関数生成が実行されました"
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then Exit Sub
    If Not ValidateConfiguration() Then Exit Sub
    
    ' Excel関数生成プリセットで実行
    Call ExecutePresetCellGeneration("excel_function")
    
    Exit Sub
    
ErrorHandler:
    ShowError "Excel関数生成実行中にエラーが発生しました", Err.Description
    LogError "OnExcelFunctionClick", Err.Description
End Sub

' データ変換（プリセット）
Public Sub OnDataTransformClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "データ変換が実行されました"
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then Exit Sub
    If Not ValidateConfiguration() Then Exit Sub
    
    ' データ変換プリセットで実行
    Call ExecutePresetCellGeneration("data_transformation")
    
    Exit Sub
    
ErrorHandler:
    ShowError "データ変換実行中にエラーが発生しました", Err.Description
    LogError "OnDataTransformClick", Err.Description
End Sub

' プリセットセル生成の実行
Private Sub ExecutePresetCellGeneration(ByVal generationType As String)
    On Error GoTo ErrorHandler
    
    ' 出力先セル（選択範囲の右隣）
    Dim targetCell As Range
    Set targetCell = Selection.Cells(1, Selection.Columns.Count + 1)
    
    ' プリセットプロンプトの取得
    Dim userPrompt As String
    Select Case generationType
        Case "excel_function"
            userPrompt = "以下のデータまたは要求を基に、適切なExcel関数を生成してください。関数のみを回答してください。"
        Case "data_transformation"
            userPrompt = "以下のデータを適切な形式に変換してください。変換後のデータのみを回答してください。"
        Case Else
            userPrompt = "以下のデータを処理してください。"
    End Select
    
    ' 追加指示の入力
    Dim additionalInstruction As String
    additionalInstruction = InputBox("追加の指示があれば入力してください（空欄可）:" & vbCrLf & vbCrLf & _
                                   "例：「日付形式をyyyy/mm/ddに変換」" & vbCrLf & _
                                   "例：「SUMIFを使って条件付き合計」", _
                                   "追加指示")
    
    If additionalInstruction <> "" Then
        userPrompt = userPrompt & vbCrLf & vbCrLf & "追加指示: " & additionalInstruction
    End If
    
    ' 確認ダイアログ
    Dim confirmMsg As String
    confirmMsg = "プリセット生成を実行しますか？" & vbCrLf & vbCrLf & _
                "生成タイプ: " & generationType & vbCrLf & _
                "対象範囲: " & Selection.Address & vbCrLf & _
                "出力先: " & targetCell.Address
    
    If ShowConfirm(confirmMsg) Then
        Call AI_CellProcessor.ProcessCellGeneration(targetCell, userPrompt, generationType)
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "プリセットセル生成実行中にエラーが発生しました", Err.Description
    LogError "ExecutePresetCellGeneration", Err.Description
End Sub

' セル内容分析
Public Sub OnCellAnalysisClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "リボンからセル内容分析が実行されました"
    Call AI_CellProcessor.AnalyzeCellContent
    
    Exit Sub
    
ErrorHandler:
    ShowError "セル内容分析実行中にエラーが発生しました", Err.Description
    LogError "OnCellAnalysisClick", Err.Description
End Sub

' =============================================================================
' 設定・管理関連のリボンコールバック
' =============================================================================

' API設定
Public Sub OnAPIConfigClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "リボンからAPI設定が実行されました"
    Call AI_ConfigManager.ShowConfigurationMenu
    
    Exit Sub
    
ErrorHandler:
    ShowError "API設定実行中にエラーが発生しました", Err.Description
    LogError "OnAPIConfigClick", Err.Description
End Sub

' プロンプト管理
Public Sub OnPromptManagerClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "リボンからプロンプト管理が実行されました"
    Call AI_PromptManager.ShowPromptManager
    
    Exit Sub
    
ErrorHandler:
    ShowError "プロンプト管理実行中にエラーが発生しました", Err.Description
    LogError "OnPromptManagerClick", Err.Description
End Sub

' ヘルプ・情報
Public Sub OnHelpClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    LogInfo "ヘルプが表示されました"
    Call ShowHelpInformation
    
    Exit Sub
    
ErrorHandler:
    ShowError "ヘルプ表示中にエラーが発生しました", Err.Description
    LogError "OnHelpClick", Err.Description
End Sub

' =============================================================================
' 情報表示関数
' =============================================================================

' システム情報表示
Private Sub ShowSystemInformation()
    On Error GoTo ErrorHandler
    
    Dim sysInfo As String
    sysInfo = "【システム情報】" & vbCrLf & vbCrLf & _
             "アプリケーション名: " & APP_NAME & vbCrLf & _
             "バージョン: " & APP_VERSION & vbCrLf & _
             "Excel バージョン: " & Application.Version & vbCrLf & _
             "Windows バージョン: " & Environ("OS") & vbCrLf & vbCrLf & _
             "【API設定状況】" & vbCrLf & _
             "エンドポイント: " & IIf(GetAPIEndpoint() <> OPENAI_API_ENDPOINT, "設定済み", "未設定") & vbCrLf & _
             "APIキー: " & IIf(GetAPIKey() <> OPENAI_API_KEY, "設定済み", "未設定") & vbCrLf & _
             "設定状態: " & IIf(ValidateConfiguration(), "正常", "要設定") & vbCrLf & vbCrLf & _
             "【機能状況】" & vbCrLf & _
             "表の行分析: 利用可能" & vbCrLf & _
             "セルAI関数生成: 利用可能" & vbCrLf & _
             "プロンプト管理: 利用可能" & vbCrLf & _
             "設定管理: 利用可能"
    
    MsgBox sysInfo, vbInformation, "システム情報"
    
    Exit Sub
    
ErrorHandler:
    ShowError "システム情報表示中にエラーが発生しました", Err.Description
    LogError "ShowSystemInformation", Err.Description
End Sub

' ヘルプ情報表示
Private Sub ShowHelpInformation()
    On Error GoTo ErrorHandler
    
    Dim helpInfo As String
    helpInfo = "【Excel AI Helper ヘルプ】" & vbCrLf & vbCrLf & _
              "■ 主要機能" & vbCrLf & _
              "・表の行分析: 選択した表の各行をAIで分析" & vbCrLf & _
              "・セルAI関数生成: セル内容からExcel関数を生成" & vbCrLf & _
              "・セル内容分析: 選択セルの内容を詳細分析" & vbCrLf & vbCrLf & _
              "■ 使用方法" & vbCrLf & _
              "1. Azure OpenAI APIキーを設定" & vbCrLf & _
              "2. 分析対象のセル範囲を選択" & vbCrLf & _
              "3. リボンメニューから機能を実行" & vbCrLf & vbCrLf & _
              "■ 注意事項" & vbCrLf & _
              "・インターネット接続が必要" & vbCrLf & _
              "・Azure OpenAIサブスクリプションが必要" & vbCrLf & _
              "・大量データは処理に時間がかかります" & vbCrLf & vbCrLf & _
              "■ サポート" & vbCrLf & _
              "設定や使用方法で困った場合は" & vbCrLf & _
              "「API設定」から「ガイド付き初期設定」を実行してください。"
    
    Call DisplayLongText("ヘルプ", helpInfo)
    
    Exit Sub
    
ErrorHandler:
    ShowError "ヘルプ表示中にエラーが発生しました", Err.Description
    LogError "ShowHelpInformation", Err.Description
End Sub

' 長いテキストの分割表示（再利用）
Private Sub DisplayLongText(ByVal title As String, ByVal text As String)
    Const maxLength As Integer = 800
    
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

' =============================================================================
' クイックアクション関数
' =============================================================================

' クイック学習意欲判定
Public Sub QuickLearningMotivation()
    On Error GoTo ErrorHandler
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then Exit Sub
    If Not ValidateConfiguration() Then Exit Sub
    
    ' 確認なしで即座に実行
    Dim outputColumn As Integer
    outputColumn = Selection.Column + Selection.Columns.Count
    
    Dim customPrompt As String
    customPrompt = "以下の行データを分析し、学習意欲の高さを0-10のスコアで評価してください。スコアのみを数値で回答してください。"
    
    Call AI_TableProcessor.ProcessTableData(Selection, "learning_motivation", outputColumn, customPrompt)
    
    Exit Sub
    
ErrorHandler:
    ShowError "クイック学習意欲判定中にエラーが発生しました", Err.Description
    LogError "QuickLearningMotivation", Err.Description
End Sub

' クイック感情分析
Public Sub QuickSentimentAnalysis()
    On Error GoTo ErrorHandler
    
    ' 選択範囲の検証
    If Not ValidateSelection() Then Exit Sub
    If Not ValidateConfiguration() Then Exit Sub
    
    ' 確認なしで即座に実行
    Dim outputColumn As Integer
    outputColumn = Selection.Column + Selection.Columns.Count
    
    Dim customPrompt As String
    customPrompt = "以下の行データを分析し、感情を「ポジティブ」「ネガティブ」「中性」のいずれかで判定してください。判定結果のみを回答してください。"
    
    Call AI_TableProcessor.ProcessTableData(Selection, "sentiment_analysis", outputColumn, customPrompt)
    
    Exit Sub
    
ErrorHandler:
    ShowError "クイック感情分析中にエラーが発生しました", Err.Description
    LogError "QuickSentimentAnalysis", Err.Description
End Sub

' =============================================================================
' リボン状態管理
' =============================================================================

' リボンボタンの有効/無効判定
Public Function GetButtonEnabled(control As IRibbonControl) As Boolean
    ' 基本的には常に有効
    GetButtonEnabled = True
    
    ' 特定の条件下でボタンを無効化したい場合はここで制御
    ' 例: API設定が不完全な場合
    ' If Not ValidateConfiguration() Then
    '     GetButtonEnabled = False
    ' End If
End Function

' リボンボタンのラベル取得
Public Function GetButtonLabel(control As IRibbonControl) As String
    Select Case control.Id
        Case "btnTableAnalysis"
            GetButtonLabel = "表の行分析"
        Case "btnCellFunction"
            GetButtonLabel = "AI関数生成"
        Case "btnAPIConfig"
            GetButtonLabel = "API設定"
        Case Else
            GetButtonLabel = control.Id
    End Select
End Function