' AI_ConfigManager.bas
' Outlook OpenAI マクロ - 設定管理モジュール
' 作成日: 2024
' 説明: API設定、プロンプト設定、その他の設定管理機能

Option Explicit

' =============================================================================
' 設定管理メイン関数
' =============================================================================

' 設定管理メニュー（メニューから呼び出される）
Public Sub ManageConfiguration()
    On Error GoTo ErrorHandler
    
    ' 設定メニューの表示
    Dim choice As String
    choice = InputBox("設定管理メニュー:" & vbCrLf & vbCrLf & _
                     "1. API設定を確認・変更" & vbCrLf & _
                     "2. プロンプト設定を管理" & vbCrLf & _
                     "3. 動作設定を変更" & vbCrLf & _
                     "4. 設定をエクスポート" & vbCrLf & _
                     "5. 設定をインポート" & vbCrLf & _
                     "6. 設定を初期化" & vbCrLf & _
                     "7. API接続テスト" & vbCrLf & vbCrLf & _
                     "番号を入力してください:", _
                     APP_NAME & " - 設定管理")
    
    Select Case choice
        Case "1"
            Call ManageAPISettings
        Case "2"
            Call ManagePromptSettings
        Case "3"
            Call ManageBehaviorSettings
        Case "4"
            Call ExportSettings
        Case "5"
            Call ImportSettings
        Case "6"
            Call ResetSettings
        Case "7"
            Call AI_ApiConnector.TestAPIConnection
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。1-7の番号を入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "設定管理中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' API設定管理
' =============================================================================

' API設定の管理
Private Sub ManageAPISettings()
    On Error GoTo ErrorHandler
    
    ' 現在の設定表示
    Dim currentSettings As String
    currentSettings = "現在のAPI設定:" & vbCrLf & vbCrLf & _
                     "エンドポイント: " & OPENAI_API_ENDPOINT & vbCrLf & _
                     "モデル: " & OPENAI_MODEL & vbCrLf & _
                     "APIキー: " & Left(OPENAI_API_KEY, 10) & "..." & vbCrLf & _
                     "タイムアウト: " & REQUEST_TIMEOUT & "秒" & vbCrLf & vbCrLf & _
                     "設定を変更しますか？"
    
    If MsgBox(currentSettings, vbYesNo + vbQuestion, "API設定") = vbNo Then
        Exit Sub
    End If
    
    ' API設定変更メニュー
    Dim settingChoice As String
    settingChoice = InputBox("変更する設定を選択してください:" & vbCrLf & vbCrLf & _
                           "1. APIエンドポイント" & vbCrLf & _
                           "2. APIキー" & vbCrLf & _
                           "3. モデル名" & vbCrLf & _
                           "4. タイムアウト設定" & vbCrLf & _
                           "5. 全ての設定をガイド付きで変更" & vbCrLf & vbCrLf & _
                           "番号を入力してください:", _
                           "API設定変更")
    
    Select Case settingChoice
        Case "1"
            Call ChangeAPIEndpoint
        Case "2"
            Call ChangeAPIKey
        Case "3"
            Call ChangeModelName
        Case "4"
            Call ChangeTimeout
        Case "5"
            Call GuidedAPISetup
        Case ""
            ' キャンセル
        Case Else
            ShowMessage "無効な選択です。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "API設定管理中にエラーが発生しました。", Err.Description
End Sub

' APIエンドポイントの変更
Private Sub ChangeAPIEndpoint()
    Dim newEndpoint As String
    newEndpoint = InputBox("新しいAPIエンドポイントを入力してください:" & vbCrLf & vbCrLf & _
                          "例: https://your-resource.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview" & vbCrLf & vbCrLf & _
                          "現在の設定: " & OPENAI_API_ENDPOINT, _
                          "APIエンドポイント変更", OPENAI_API_ENDPOINT)
    
    If newEndpoint <> "" And newEndpoint <> OPENAI_API_ENDPOINT Then
        If ValidateEndpoint(newEndpoint) Then
            ' 実際の変更はユーザーがコード内で手動で行う必要があることを説明
            ShowMessage "新しいエンドポイント: " & newEndpoint & vbCrLf & vbCrLf & _
                       "この設定を適用するには、AI_Common.bas ファイルの" & vbCrLf & _
                       "OPENAI_API_ENDPOINT 定数を手動で変更してください。", _
                       "設定変更が必要", vbInformation
        End If
    End If
End Sub

' APIキーの変更
Private Sub ChangeAPIKey()
    Dim newAPIKey As String
    newAPIKey = InputBox("新しいAPIキーを入力してください:" & vbCrLf & vbCrLf & _
                        "注意: APIキーは機密情報です。安全に管理してください。" & vbCrLf & vbCrLf & _
                        "現在の設定: " & Left(OPENAI_API_KEY, 10) & "...", _
                        "APIキー変更")
    
    If newAPIKey <> "" And newAPIKey <> OPENAI_API_KEY Then
        If ValidateAPIKey(newAPIKey) Then
            ShowMessage "新しいAPIキー: " & Left(newAPIKey, 10) & "..." & vbCrLf & vbCrLf & _
                       "この設定を適用するには、AI_Common.bas ファイルの" & vbCrLf & _
                       "OPENAI_API_KEY 定数を手動で変更してください。", _
                       "設定変更が必要", vbInformation
        End If
    End If
End Sub

' モデル名の変更
Private Sub ChangeModelName()
    Dim newModel As String
    Dim modelChoice As String
    
    modelChoice = InputBox("使用するモデルを選択してください:" & vbCrLf & vbCrLf & _
                          "1. gpt-4" & vbCrLf & _
                          "2. gpt-4-turbo" & vbCrLf & _
                          "3. gpt-35-turbo" & vbCrLf & _
                          "4. カスタム（手動入力）" & vbCrLf & vbCrLf & _
                          "現在の設定: " & OPENAI_MODEL, _
                          "モデル選択")
    
    Select Case modelChoice
        Case "1": newModel = "gpt-4"
        Case "2": newModel = "gpt-4-turbo"
        Case "3": newModel = "gpt-35-turbo"
        Case "4"
            newModel = InputBox("カスタムモデル名を入力してください:", "カスタムモデル", OPENAI_MODEL)
        Case "": Exit Sub
        Case Else
            ShowMessage "無効な選択です。", "入力エラー", vbExclamation
            Exit Sub
    End Select
    
    If newModel <> "" And newModel <> OPENAI_MODEL Then
        ShowMessage "新しいモデル: " & newModel & vbCrLf & vbCrLf & _
                   "この設定を適用するには、AI_Common.bas ファイルの" & vbCrLf & _
                   "OPENAI_MODEL 定数を手動で変更してください。", _
                   "設定変更が必要", vbInformation
    End If
End Sub

' タイムアウト設定の変更
Private Sub ChangeTimeout()
    Dim newTimeout As String
    Dim timeoutValue As Integer
    
    newTimeout = InputBox("新しいタイムアウト値（秒）を入力してください:" & vbCrLf & vbCrLf & _
                         "推奨値: 30-120秒" & vbCrLf & _
                         "現在の設定: " & REQUEST_TIMEOUT & "秒", _
                         "タイムアウト変更", CStr(REQUEST_TIMEOUT))
    
    If newTimeout <> "" Then
        If IsNumeric(newTimeout) Then
            timeoutValue = CInt(newTimeout)
            If timeoutValue >= 10 And timeoutValue <= 300 Then
                ShowMessage "新しいタイムアウト: " & timeoutValue & "秒" & vbCrLf & vbCrLf & _
                           "この設定を適用するには、AI_Common.bas ファイルの" & vbCrLf & _
                           "REQUEST_TIMEOUT 定数を手動で変更してください。", _
                           "設定変更が必要", vbInformation
            Else
                ShowMessage "タイムアウト値は10-300秒の範囲で入力してください。", "値が範囲外", vbExclamation
            End If
        Else
            ShowMessage "数値を入力してください。", "入力エラー", vbExclamation
        End If
    End If
End Sub

' ガイド付きAPI設定
Private Sub GuidedAPISetup()
    On Error GoTo ErrorHandler
    
    ShowMessage "ガイド付きAPI設定を開始します。" & vbCrLf & _
               "Azure OpenAI サービスの設定に必要な情報を順番に入力してください。", _
               "設定ガイド", vbInformation
    
    ' ステップ1: リソース名
    Dim resourceName As String
    resourceName = InputBox("Azure OpenAI リソース名を入力してください:" & vbCrLf & vbCrLf & _
                           "例: my-openai-resource" & vbCrLf & _
                           "（Azure ポータルで確認できます）", _
                           "リソース名")
    
    If resourceName = "" Then Exit Sub
    
    ' ステップ2: デプロイメント名
    Dim deploymentName As String
    deploymentName = InputBox("デプロイメント名を入力してください:" & vbCrLf & vbCrLf & _
                             "例: gpt-4-deployment" & vbCrLf & _
                             "（Azure OpenAI Studio で設定したモデルのデプロイメント名）", _
                             "デプロイメント名")
    
    If deploymentName = "" Then Exit Sub
    
    ' ステップ3: APIキー
    Dim apiKey As String
    apiKey = InputBox("APIキーを入力してください:" & vbCrLf & vbCrLf & _
                     "Azure ポータルの「キーとエンドポイント」から取得できます。" & vbCrLf & _
                     "Key1 または Key2 を使用してください。", _
                     "APIキー")
    
    If apiKey = "" Then Exit Sub
    
    ' ステップ4: API バージョン
    Dim apiVersion As String
    Dim versionChoice As String
    versionChoice = InputBox("API バージョンを選択してください:" & vbCrLf & vbCrLf & _
                            "1. 2024-02-15-preview（推奨）" & vbCrLf & _
                            "2. 2023-12-01-preview" & vbCrLf & _
                            "3. カスタム" & vbCrLf & vbCrLf & _
                            "番号を入力してください:", _
                            "APIバージョン")
    
    Select Case versionChoice
        Case "1": apiVersion = "2024-02-15-preview"
        Case "2": apiVersion = "2023-12-01-preview"
        Case "3": apiVersion = InputBox("カスタムAPIバージョンを入力してください:", "カスタムバージョン")
        Case "": Exit Sub
        Case Else
            ShowMessage "無効な選択です。", "入力エラー", vbExclamation
            Exit Sub
    End Select
    
    ' 設定内容の確認と表示
    Dim endpoint As String
    endpoint = "https://" & resourceName & ".openai.azure.com/openai/deployments/" & deploymentName & "/chat/completions?api-version=" & apiVersion
    
    Dim settingSummary As String
    settingSummary = "設定内容を確認してください:" & vbCrLf & vbCrLf & _
                    "リソース名: " & resourceName & vbCrLf & _
                    "デプロイメント名: " & deploymentName & vbCrLf & _
                    "APIバージョン: " & apiVersion & vbCrLf & _
                    "APIキー: " & Left(apiKey, 10) & "..." & vbCrLf & vbCrLf & _
                    "生成されたエンドポイント:" & vbCrLf & endpoint & vbCrLf & vbCrLf & _
                    "この設定でAPI接続テストを行いますか？"
    
    If MsgBox(settingSummary, vbYesNo + vbQuestion, "設定確認") = vbYes Then
        ' テスト実行（実際のテストは簡易版）
        ShowMessage "設定内容を AI_Common.bas に手動で適用してから、" & vbCrLf & _
                   "「API接続テスト」を実行してください。" & vbCrLf & vbCrLf & _
                   "設定値:" & vbCrLf & _
                   "OPENAI_API_ENDPOINT = """ & endpoint & """" & vbCrLf & _
                   "OPENAI_API_KEY = """ & apiKey & """", _
                   "設定完了", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowError "ガイド付き設定中にエラーが発生しました。", Err.Description
End Sub

' =============================================================================
' プロンプト設定管理
' =============================================================================

' プロンプト設定の管理
Private Sub ManagePromptSettings()
    On Error GoTo ErrorHandler
    
    ' プロンプト管理メニュー
    Dim promptChoice As String
    promptChoice = InputBox("プロンプト設定管理:" & vbCrLf & vbCrLf & _
                           "1. 現在のプロンプトを確認" & vbCrLf & _
                           "2. システムプロンプトをカスタマイズ" & vbCrLf & _
                           "3. プロンプトテンプレートを管理" & vbCrLf & _
                           "4. プロンプトを初期化" & vbCrLf & vbCrLf & _
                           "番号を入力してください:", _
                           "プロンプト管理")
    
    Select Case promptChoice
        Case "1"
            Call ShowCurrentPrompts
        Case "2"
            Call CustomizeSystemPrompts
        Case "3"
            Call ManagePromptTemplates
        Case "4"
            Call ResetPrompts
        Case ""
            ' キャンセル
        Case Else
            ShowMessage "無効な選択です。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "プロンプト設定管理中にエラーが発生しました。", Err.Description
End Sub

' 現在のプロンプト表示
Private Sub ShowCurrentPrompts()
    Dim promptInfo As String
    promptInfo = "現在のシステムプロンプト設定:" & vbCrLf & vbCrLf & _
                "【メール分析用】" & vbCrLf & SYSTEM_PROMPT_ANALYZER & vbCrLf & vbCrLf & _
                "【メール作成用】" & vbCrLf & SYSTEM_PROMPT_COMPOSER & vbCrLf & vbCrLf & _
                "【検索分析用】" & vbCrLf & SYSTEM_PROMPT_SEARCH
    
    ' 長いテキストなので分割表示
    ShowLongText promptInfo, "現在のプロンプト設定"
End Sub

' システムプロンプトのカスタマイズ
Private Sub CustomizeSystemPrompts()
    ShowMessage "システムプロンプトのカスタマイズ機能は準備中です。" & vbCrLf & _
               "現在は AI_Common.bas ファイル内の定数を直接編集してください。" & vbCrLf & vbCrLf & _
               "対象定数:" & vbCrLf & _
               "- SYSTEM_PROMPT_ANALYZER" & vbCrLf & _
               "- SYSTEM_PROMPT_COMPOSER" & vbCrLf & _
               "- SYSTEM_PROMPT_SEARCH", _
               "機能準備中", vbInformation
End Sub

' プロンプトテンプレート管理
Private Sub ManagePromptTemplates()
    ShowMessage "プロンプトテンプレート管理機能は準備中です。" & vbCrLf & _
               "将来のバージョンで実装予定です。", _
               "機能準備中", vbInformation
End Sub

' プロンプト初期化
Private Sub ResetPrompts()
    If MsgBox("プロンプト設定を初期状態に戻しますか？" & vbCrLf & _
              "（AI_Common.bas ファイルを手動で初期値に戻す必要があります）", _
              vbYesNo + vbQuestion, "プロンプト初期化") = vbYes Then
        
        ShowMessage "プロンプトを初期化するには、AI_Common.bas ファイル内の" & vbCrLf & _
                   "以下の定数を元の値に戻してください:" & vbCrLf & vbCrLf & _
                   "- SYSTEM_PROMPT_ANALYZER" & vbCrLf & _
                   "- SYSTEM_PROMPT_COMPOSER" & vbCrLf & _
                   "- SYSTEM_PROMPT_SEARCH", _
                   "初期化手順", vbInformation
    End If
End Sub

' =============================================================================
' 動作設定管理
' =============================================================================

' 動作設定の管理
Private Sub ManageBehaviorSettings()
    ShowMessage "動作設定管理機能は準備中です。" & vbCrLf & _
               "現在は AI_Common.bas ファイル内の定数を直接編集してください。" & vbCrLf & vbCrLf & _
               "主な設定項目:" & vbCrLf & _
               "- MAX_CONTENT_LENGTH（最大処理文字数）" & vbCrLf & _
               "- REQUEST_TIMEOUT（タイムアウト）", _
               "機能準備中", vbInformation
End Sub

' =============================================================================
' 設定インポート・エクスポート
' =============================================================================

' 設定のエクスポート
Private Sub ExportSettings()
    On Error GoTo ErrorHandler
    
    Dim settings As String
    settings = "# Outlook AI Helper 設定ファイル" & vbCrLf & _
              "# 作成日時: " & Now & vbCrLf & vbCrLf & _
              "OPENAI_API_ENDPOINT=" & OPENAI_API_ENDPOINT & vbCrLf & _
              "OPENAI_API_KEY=" & OPENAI_API_KEY & vbCrLf & _
              "OPENAI_MODEL=" & OPENAI_MODEL & vbCrLf & _
              "REQUEST_TIMEOUT=" & REQUEST_TIMEOUT & vbCrLf & _
              "MAX_CONTENT_LENGTH=" & MAX_CONTENT_LENGTH & vbCrLf & vbCrLf & _
              "# システムプロンプト" & vbCrLf & _
              "SYSTEM_PROMPT_ANALYZER=" & SYSTEM_PROMPT_ANALYZER & vbCrLf & _
              "SYSTEM_PROMPT_COMPOSER=" & SYSTEM_PROMPT_COMPOSER & vbCrLf & _
              "SYSTEM_PROMPT_SEARCH=" & SYSTEM_PROMPT_SEARCH
    
    ' デスクトップに保存
    Dim fileName As String
    fileName = Environ("USERPROFILE") & "\Desktop\OutlookAIHelper_Settings_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open fileName For Output As #fileNum
    Print #fileNum, settings
    Close #fileNum
    
    ShowSuccess "設定をエクスポートしました:" & vbCrLf & fileName
    
    Exit Sub
    
ErrorHandler:
    ShowError "設定エクスポート中にエラーが発生しました。", Err.Description
End Sub

' 設定のインポート
Private Sub ImportSettings()
    ShowMessage "設定インポート機能は準備中です。" & vbCrLf & _
               "現在は手動でAI_Common.basファイルを編集してください。", _
               "機能準備中", vbInformation
End Sub

' =============================================================================
' 設定初期化
' =============================================================================

' 設定の初期化
Private Sub ResetSettings()
    If MsgBox("全ての設定を初期状態に戻しますか？" & vbCrLf & _
              "この操作は元に戻せません。", _
              vbYesNo + vbExclamation, "設定初期化") = vbYes Then
        
        ShowMessage "設定を初期化するには、以下の手順を実行してください:" & vbCrLf & vbCrLf & _
                   "1. AI_Common.bas ファイルをバックアップ" & vbCrLf & _
                   "2. 元のファイルから定数値をコピーして復元" & vbCrLf & _
                   "3. VBA エディタで変更を保存" & vbCrLf & _
                   "4. Outlook を再起動", _
                   "初期化手順", vbInformation
    End If
End Sub

' =============================================================================
' バリデーション関数
' =============================================================================

' エンドポイントの検証
Private Function ValidateEndpoint(ByVal endpoint As String) As Boolean
    If Left(endpoint, 8) <> "https://" Then
        ShowMessage "エンドポイントは https:// で始まる必要があります。", "検証エラー", vbExclamation
        ValidateEndpoint = False
    ElseIf InStr(endpoint, "openai.azure.com") = 0 Then
        ShowMessage "Azure OpenAI のエンドポイントを指定してください。", "検証エラー", vbExclamation
        ValidateEndpoint = False
    Else
        ValidateEndpoint = True
    End If
End Function

' APIキーの検証
Private Function ValidateAPIKey(ByVal apiKey As String) As Boolean
    If Len(apiKey) < 10 Then
        ShowMessage "APIキーが短すぎます。正しいキーを入力してください。", "検証エラー", vbExclamation
        ValidateAPIKey = False
    ElseIf apiKey = "YOUR_API_KEY_HERE" Then
        ShowMessage "デフォルトのプレースホルダーです。実際のAPIキーを入力してください。", "検証エラー", vbExclamation
        ValidateAPIKey = False
    Else
        ValidateAPIKey = True
    End If
End Function

' 長いテキストの分割表示（他のモジュールからコピー）
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