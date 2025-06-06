' AI_ConfigManager.bas
' Excel AI Helper - 設定管理モジュール
' 作成日: 2024
' 説明: API設定、プロンプトテンプレート、動作設定の管理

Option Explicit

' =============================================================================
' 設定管理メイン関数
' =============================================================================

' 設定画面メインメニュー
Public Sub ShowConfigurationMenu()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    Dim menuText As String
    
    menuText = "設定管理メニュー" & vbCrLf & vbCrLf & _
               "1. API設定（Azure OpenAI）" & vbCrLf & _
               "2. プロンプトテンプレート管理" & vbCrLf & _
               "3. 動作設定" & vbCrLf & _
               "4. 設定確認" & vbCrLf & _
               "5. 設定のリセット" & vbCrLf & vbCrLf & _
               "実行したい項目の番号を入力してください:"
    
    choice = InputBox(menuText, APP_NAME & " - 設定管理")
    
    Select Case choice
        Case "1"
            Call ConfigureAPI
        Case "2"
            Call ManagePromptTemplates
        Case "3"
            Call ConfigureGeneral
        Case "4"
            Call ShowCurrentConfiguration
        Case "5"
            Call ResetConfiguration
        Case ""
            ' キャンセル
        Case Else
            ShowWarning "無効な選択です。1-5の番号を入力してください。"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "設定管理中にエラーが発生しました", Err.Description
End Sub

' =============================================================================
' API設定管理
' =============================================================================

' API設定画面
Private Sub ConfigureAPI()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    choice = InputBox("API設定" & vbCrLf & vbCrLf & _
                     "1. ガイド付き初期設定" & vbCrLf & _
                     "2. エンドポイント変更" & vbCrLf & _
                     "3. APIキー変更" & vbCrLf & _
                     "4. モデル変更" & vbCrLf & _
                     "5. タイムアウト設定" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     "API設定")
    
    Select Case choice
        Case "1"
            Call GuidedAPISetup
        Case "2"
            Call ChangeAPIEndpoint
        Case "3"
            Call ChangeAPIKey
        Case "4"
            Call ChangeModelName
        Case "5"
            Call ChangeTimeout
        Case ""
            ' キャンセル
        Case Else
            ShowWarning "無効な選択です。1-5の番号を入力してください。"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "API設定中にエラーが発生しました", Err.Description
End Sub

' ガイド付きAPI設定
Private Sub GuidedAPISetup()
    On Error GoTo ErrorHandler
    
    ' 設定開始の説明
    MsgBox "Azure OpenAI APIの設定を行います。" & vbCrLf & vbCrLf & _
           "以下の情報をAzureポータルから取得してください：" & vbCrLf & _
           "1. エンドポイントURL" & vbCrLf & _
           "2. APIキー" & vbCrLf & _
           "3. デプロイメント名", _
           vbInformation, "Azure OpenAI 設定ガイド"
    
    ' ステップ1: エンドポイント
    Dim endpoint As String
    endpoint = InputBox("ステップ1: エンドポイントURL" & vbCrLf & vbCrLf & _
                       "Azure ポータルの「キーとエンドポイント」からエンドポイントURLをコピーしてください。" & vbCrLf & vbCrLf & _
                       "例: https://your-resource.openai.azure.com/", _
                       "エンドポイント設定")
    
    If endpoint = "" Then Exit Sub
    
    ' ステップ2: デプロイメント名
    Dim deploymentName As String
    deploymentName = InputBox("ステップ2: デプロイメント名" & vbCrLf & vbCrLf & _
                             "Azure OpenAI Studioでデプロイしたモデルのデプロイメント名を入力してください。" & vbCrLf & vbCrLf & _
                             "例: gpt-4, gpt-35-turbo", _
                             "デプロイメント名")
    
    If deploymentName = "" Then Exit Sub
    
    ' ステップ3: APIキー
    Dim apiKey As String
    apiKey = InputBox("ステップ3: APIキー" & vbCrLf & vbCrLf & _
                     "Azure ポータルの「キーとエンドポイント」から取得できます。" & vbCrLf & _
                     "Key1 または Key2 を使用してください。", _
                     "APIキー")
    
    If apiKey = "" Then Exit Sub
    
    ' ステップ4: API バージョン
    Dim apiVersion As String
    Dim versionChoice As String
    versionChoice = InputBox("ステップ4: API バージョン" & vbCrLf & vbCrLf & _
                            "1. 2024-02-15-preview（推奨）" & vbCrLf & _
                            "2. 2023-12-01-preview" & vbCrLf & _
                            "3. カスタム" & vbCrLf & vbCrLf & _
                            "番号を入力してください:", _
                            "APIバージョン")
    
    Select Case versionChoice
        Case "1"
            apiVersion = "2024-02-15-preview"
        Case "2"
            apiVersion = "2023-12-01-preview"
        Case "3"
            apiVersion = InputBox("カスタムAPIバージョンを入力してください:", "APIバージョン")
            If apiVersion = "" Then Exit Sub
        Case Else
            Exit Sub
    End Select
    
    ' 完全なエンドポイントURL構築
    Dim fullEndpoint As String
    fullEndpoint = endpoint & "/openai/deployments/" & deploymentName & "/chat/completions?api-version=" & apiVersion
    
    ' 設定の保存
    Call SaveAPISetting("Endpoint", fullEndpoint)
    Call SaveAPISetting("APIKey", apiKey)
    Call SaveAPISetting("Model", deploymentName)
    
    ' 設定テスト
    If ShowConfirm("設定を保存しました。接続テストを実行しますか？") Then
        Call TestAPIConnection
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "ガイド付き設定中にエラーが発生しました", Err.Description
End Sub

' APIエンドポイントの変更
Private Sub ChangeAPIEndpoint()
    Dim newEndpoint As String
    newEndpoint = InputBox("新しいAPIエンドポイントを入力してください:" & vbCrLf & vbCrLf & _
                          "現在の設定: " & GetAPISetting("Endpoint", OPENAI_API_ENDPOINT), _
                          "エンドポイント変更")
    
    If newEndpoint <> "" Then
        Call SaveAPISetting("Endpoint", newEndpoint)
        ShowSuccess "エンドポイントを更新しました。"
    End If
End Sub

' APIキーの変更
Private Sub ChangeAPIKey()
    Dim newKey As String
    newKey = InputBox("新しいAPIキーを入力してください:" & vbCrLf & vbCrLf & _
                     "セキュリティのため、現在のキーは表示されません。", _
                     "APIキー変更")
    
    If newKey <> "" Then
        Call SaveAPISetting("APIKey", newKey)
        ShowSuccess "APIキーを更新しました。"
    End If
End Sub

' モデル名の変更
Private Sub ChangeModelName()
    Dim newModel As String
    newModel = InputBox("新しいモデル名（デプロイメント名）を入力してください:" & vbCrLf & vbCrLf & _
                       "現在の設定: " & GetAPISetting("Model", OPENAI_MODEL), _
                       "モデル変更")
    
    If newModel <> "" Then
        Call SaveAPISetting("Model", newModel)
        ShowSuccess "モデル名を更新しました。"
    End If
End Sub

' タイムアウト設定の変更
Private Sub ChangeTimeout()
    Dim newTimeout As String
    newTimeout = InputBox("新しいタイムアウト値（秒）を入力してください:" & vbCrLf & vbCrLf & _
                         "現在の設定: " & GetAPISetting("Timeout", CStr(REQUEST_TIMEOUT)) & "秒" & vbCrLf & _
                         "推奨値: 30-120秒", _
                         "タイムアウト設定")
    
    If newTimeout <> "" And IsNumeric(newTimeout) Then
        Call SaveAPISetting("Timeout", newTimeout)
        ShowSuccess "タイムアウト値を更新しました。"
    ElseIf newTimeout <> "" Then
        ShowWarning "数値を入力してください。"
    End If
End Sub

' =============================================================================
' プロンプトテンプレート管理
' =============================================================================

' プロンプトテンプレート管理画面
Private Sub ManagePromptTemplates()
    On Error GoTo ErrorHandler
    
    Dim choice As String
    choice = InputBox("プロンプトテンプレート管理" & vbCrLf & vbCrLf & _
                     "1. 表分析用テンプレート編集" & vbCrLf & _
                     "2. セル処理用テンプレート編集" & vbCrLf & _
                     "3. カスタムテンプレート追加" & vbCrLf & _
                     "4. テンプレート一覧表示" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     "プロンプトテンプレート管理")
    
    Select Case choice
        Case "1"
            Call EditTableAnalysisTemplate
        Case "2"
            Call EditCellProcessingTemplate
        Case "3"
            Call AddCustomTemplate
        Case "4"
            Call ShowTemplateList
        Case ""
            ' キャンセル
        Case Else
            ShowWarning "無効な選択です。1-4の番号を入力してください。"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "プロンプトテンプレート管理中にエラーが発生しました", Err.Description
End Sub

' 表分析用テンプレート編集
Private Sub EditTableAnalysisTemplate()
    Dim currentTemplate As String
    currentTemplate = GetAPISetting("TableAnalysisTemplate", SYSTEM_PROMPT_TABLE_ANALYZER)
    
    Dim newTemplate As String
    ' 注：実際の実装では専用のフォームを使用
    newTemplate = InputBox("表分析用システムプロンプトを編集してください:" & vbCrLf & vbCrLf & _
                          "現在の設定:" & vbCrLf & currentTemplate, _
                          "表分析テンプレート編集", currentTemplate)
    
    If newTemplate <> "" Then
        Call SaveAPISetting("TableAnalysisTemplate", newTemplate)
        ShowSuccess "表分析用テンプレートを更新しました。"
    End If
End Sub

' セル処理用テンプレート編集
Private Sub EditCellProcessingTemplate()
    Dim currentTemplate As String
    currentTemplate = GetAPISetting("CellProcessingTemplate", SYSTEM_PROMPT_CELL_PROCESSOR)
    
    Dim newTemplate As String
    newTemplate = InputBox("セル処理用システムプロンプトを編集してください:" & vbCrLf & vbCrLf & _
                          "現在の設定:" & vbCrLf & currentTemplate, _
                          "セル処理テンプレート編集", currentTemplate)
    
    If newTemplate <> "" Then
        Call SaveAPISetting("CellProcessingTemplate", newTemplate)
        ShowSuccess "セル処理用テンプレートを更新しました。"
    End If
End Sub

' =============================================================================
' 設定の保存・読み込み
' =============================================================================

' 設定の保存（レジストリ使用）
Private Sub SaveAPISetting(ByVal key As String, ByVal value As String)
    On Error GoTo ErrorHandler
    
    ' レジストリパスの構築
    Dim registryPath As String
    registryPath = "HKEY_CURRENT_USER\Software\ExcelAIHelper\"
    
    ' CreateObjectを使用してレジストリアクセス
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.RegWrite registryPath & key, value, "REG_SZ"
    
    LogInfo "設定保存: " & key & " = " & Left(value, 50) & "..."
    
    Exit Sub
    
ErrorHandler:
    LogError "SaveAPISetting", "設定保存エラー - Key: " & key & ", Error: " & Err.Description
End Sub

' 設定の読み込み（レジストリ使用）
Private Function GetAPISetting(ByVal key As String, ByVal defaultValue As String) As String
    On Error GoTo ErrorHandler
    
    Dim registryPath As String
    registryPath = "HKEY_CURRENT_USER\Software\ExcelAIHelper\"
    
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    GetAPISetting = shell.RegRead(registryPath & key)
    
    Exit Function
    
ErrorHandler:
    ' エラーの場合はデフォルト値を返す
    GetAPISetting = defaultValue
End Function

' =============================================================================
' 設定表示・リセット機能
' =============================================================================

' 現在の設定表示
Private Sub ShowCurrentConfiguration()
    Dim configInfo As String
    configInfo = "現在の設定:" & vbCrLf & vbCrLf & _
                "アプリケーション名: " & APP_NAME & vbCrLf & _
                "バージョン: " & APP_VERSION & vbCrLf & vbCrLf & _
                "【API設定】" & vbCrLf & _
                "エンドポイント: " & Left(GetAPISetting("Endpoint", OPENAI_API_ENDPOINT), 50) & "..." & vbCrLf & _
                "APIキー: " & Left(GetAPISetting("APIKey", OPENAI_API_KEY), 10) & "..." & vbCrLf & _
                "モデル: " & GetAPISetting("Model", OPENAI_MODEL) & vbCrLf & _
                "タイムアウト: " & GetAPISetting("Timeout", CStr(REQUEST_TIMEOUT)) & "秒" & vbCrLf & vbCrLf & _
                "【動作設定】" & vbCrLf & _
                "最大処理文字数: " & MAX_CONTENT_LENGTH & vbCrLf & _
                "1回の最大処理行数: " & MAX_ROWS_PER_BATCH
    
    MsgBox configInfo, vbInformation, "設定確認"
End Sub

' 設定のリセット
Private Sub ResetConfiguration()
    If ShowConfirm("すべての設定をデフォルト値にリセットしますか？" & vbCrLf & vbCrLf & _
                  "この操作は元に戻せません。") Then
        
        ' 設定をデフォルト値で上書き
        Call SaveAPISetting("Endpoint", OPENAI_API_ENDPOINT)
        Call SaveAPISetting("APIKey", OPENAI_API_KEY)
        Call SaveAPISetting("Model", OPENAI_MODEL)
        Call SaveAPISetting("Timeout", CStr(REQUEST_TIMEOUT))
        Call SaveAPISetting("TableAnalysisTemplate", SYSTEM_PROMPT_TABLE_ANALYZER)
        Call SaveAPISetting("CellProcessingTemplate", SYSTEM_PROMPT_CELL_PROCESSOR)
        
        ShowSuccess "設定をリセットしました。"
        LogInfo "設定リセット実行"
    End If
End Sub

' =============================================================================
' API接続テスト
' =============================================================================

' API接続テスト
Private Sub TestAPIConnection()
    On Error GoTo ErrorHandler
    
    ShowProgress 0, 1, "API接続をテストしています..."
    
    Dim testResult As String
    testResult = SendOpenAIRequest("あなたは接続テスト用のアシスタントです。", "「接続テスト成功」と日本語で返答してください。")
    
    ClearProgress
    
    If testResult <> "" Then
        ShowSuccess "API接続テストが成功しました。" & vbCrLf & vbCrLf & "レスポンス: " & testResult
        LogInfo "API接続テスト成功"
    Else
        ShowError "API接続テストが失敗しました", "設定を確認してください。"
        LogError "TestAPIConnection", "API接続テスト失敗"
    End If
    
    Exit Sub
    
ErrorHandler:
    ClearProgress
    ShowError "API接続テスト中にエラーが発生しました", Err.Description
    LogError "TestAPIConnection", Err.Description
End Sub

' =============================================================================
' 拡張設定用関数（将来拡張）
' =============================================================================

' 一般設定
Private Sub ConfigureGeneral()
    ShowWarning "一般設定機能は将来のバージョンで実装予定です。"
End Sub

' カスタムテンプレート追加
Private Sub AddCustomTemplate()
    ShowWarning "カスタムテンプレート機能は将来のバージョンで実装予定です。"
End Sub

' テンプレート一覧表示
Private Sub ShowTemplateList()
    ShowWarning "テンプレート一覧機能は将来のバージョンで実装予定です。"
End Sub