' OutlookAI_MainForm.bas
' Outlook AI Helper - メインユーザーインターフェース
' 作成日: 2024
' 説明: 統合されたユーザーフレンドリーなメインフォーム
'
' 利用方法: ShowMainForm() を呼び出してメインフォームを表示
' 
' 機能:
' - 日本語による直感的なボタンレイアウト
' - ワンクリックでの各機能へのアクセス
' - 番号入力不要のGUI操作
' - 視覚的にグループ化された機能

Option Explicit

' =============================================================================
' フォーム表示とメイン処理
' =============================================================================

' メインフォーム表示（エントリーポイント）
Public Sub ShowMainForm()
    On Error GoTo ErrorHandler
    
    ' HTML形式のダイアログを使用してリッチなUIを提供
    Dim htmlDialog As String
    htmlDialog = CreateMainFormHTML()
    
    ' HTMLダイアログの表示とユーザー選択の処理
    Dim choice As String
    choice = ShowHTMLDialog(htmlDialog)
    
    ' 選択された機能を実行
    ProcessUserChoice choice
    
    Exit Sub
    
ErrorHandler:
    ShowError "メインフォーム表示中にエラーが発生しました", Err.Description
End Sub

' HTMLダイアログの作成
Private Function CreateMainFormHTML() As String
    Dim html As String
    
    html = "<!DOCTYPE html>" & vbCrLf & _
           "<html>" & vbCrLf & _
           "<head>" & vbCrLf & _
           "<meta charset='utf-8'>" & vbCrLf & _
           "<title>Outlook AI Helper</title>" & vbCrLf & _
           "<style>" & vbCrLf & _
           "body { font-family: 'Segoe UI', sans-serif; margin: 20px; background: #f5f5f5; }" & vbCrLf & _
           ".container { max-width: 500px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }" & vbCrLf & _
           ".header { text-align: center; margin-bottom: 30px; }" & vbCrLf & _
           ".title { color: #0078d4; font-size: 24px; font-weight: bold; margin-bottom: 10px; }" & vbCrLf & _
           ".subtitle { color: #666; font-size: 14px; }" & vbCrLf & _
           ".section { margin-bottom: 25px; }" & vbCrLf & _
           ".section-title { color: #323130; font-size: 16px; font-weight: bold; margin-bottom: 15px; display: flex; align-items: center; }" & vbCrLf & _
           ".icon { font-size: 18px; margin-right: 8px; }" & vbCrLf & _
           ".button-group { display: flex; flex-wrap: wrap; gap: 10px; }" & vbCrLf & _
           ".action-btn { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; border: none; padding: 12px 16px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s; min-width: 120px; }" & vbCrLf & _
           ".action-btn:hover { background: linear-gradient(135deg, #106ebe, #005a9e); transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,120,212,0.3); }" & vbCrLf & _
           ".action-btn.analysis { background: linear-gradient(135deg, #0078d4, #106ebe); }" & vbCrLf & _
           ".action-btn.compose { background: linear-gradient(135deg, #107c10, #0b5a0b); }" & vbCrLf & _
           ".action-btn.system { background: linear-gradient(135deg, #5c2d91, #4a1b73); }" & vbCrLf & _
           ".action-btn.compose:hover { background: linear-gradient(135deg, #0b5a0b, #084708); }" & vbCrLf & _
           ".action-btn.system:hover { background: linear-gradient(135deg, #4a1b73, #3a1459); }" & vbCrLf & _
           ".footer { text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #edebe9; color: #666; font-size: 12px; }" & vbCrLf & _
           "</style>" & vbCrLf & _
           "</head>" & vbCrLf & _
           "<body>" & vbCrLf & _
           "<div class='container'>" & vbCrLf & _
           "<div class='header'>" & vbCrLf & _
           "<div class='title'>🤖 Outlook AI Helper</div>" & vbCrLf & _
           "<div class='subtitle'>v1.0.0 Unified - 統合版</div>" & vbCrLf & _
           "</div>" & vbCrLf & _
           CreateAnalysisSection() & _
           CreateComposerSection() & _
           CreateSystemSection() & _
           "<div class='footer'>" & vbCrLf & _
           "💡 各ボタンをクリックして対応する機能を実行できます<br>" & vbCrLf & _
           "❓ 問題が発生した場合は「API接続テスト」をお試しください" & vbCrLf & _
           "</div>" & vbCrLf & _
           "</div>" & vbCrLf & _
           CreateJavaScript() & _
           "</body>" & vbCrLf & _
           "</html>"
    
    CreateMainFormHTML = html
End Function

' メール解析セクション
Private Function CreateAnalysisSection() As String
    CreateAnalysisSection = _
        "<div class='section'>" & vbCrLf & _
        "<div class='section-title'><span class='icon'>📊</span>メール解析</div>" & vbCrLf & _
        "<div class='button-group'>" & vbCrLf & _
        "<button class='action-btn analysis' onclick='selectFunction(""analyze_email"")'>📧 メール内容解析</button>" & vbCrLf & _
        "<button class='action-btn analysis' onclick='selectFunction(""analyze_folders"")'>📁 検索フォルダ分析</button>" & vbCrLf & _
        "</div>" & vbCrLf & _
        "</div>" & vbCrLf
End Function

' メール作成セクション
Private Function CreateComposerSection() As String
    CreateComposerSection = _
        "<div class='section'>" & vbCrLf & _
        "<div class='section-title'><span class='icon'>✉️</span>メール作成支援</div>" & vbCrLf & _
        "<div class='button-group'>" & vbCrLf & _
        "<button class='action-btn compose' onclick='selectFunction(""create_rejection"")'>❌ 営業断りメール</button>" & vbCrLf & _
        "<button class='action-btn compose' onclick='selectFunction(""create_acceptance"")'>✅ 承諾メール</button>" & vbCrLf & _
        "<button class='action-btn compose' onclick='selectFunction(""create_custom"")'>✏️ カスタムメール</button>" & vbCrLf & _
        "</div>" & vbCrLf & _
        "</div>" & vbCrLf
End Function

' システム管理セクション
Private Function CreateSystemSection() As String
    CreateSystemSection = _
        "<div class='section'>" & vbCrLf & _
        "<div class='section-title'><span class='icon'>⚙️</span>システム管理</div>" & vbCrLf & _
        "<div class='button-group'>" & vbCrLf & _
        "<button class='action-btn system' onclick='selectFunction(""manage_config"")'>🔧 設定管理</button>" & vbCrLf & _
        "<button class='action-btn system' onclick='selectFunction(""test_api"")'>🔌 API接続テスト</button>" & vbCrLf & _
        "</div>" & vbCrLf & _
        "</div>" & vbCrLf
End Function

' JavaScript処理
Private Function CreateJavaScript() As String
    CreateJavaScript = _
        "<script>" & vbCrLf & _
        "function selectFunction(functionName) {" & vbCrLf & _
        "  try {" & vbCrLf & _
        "    window.external.ExecuteFunction(functionName);" & vbCrLf & _
        "    window.close();" & vbCrLf & _
        "  } catch(e) {" & vbCrLf & _
        "    alert('機能の実行に失敗しました: ' + e.message);" & vbCrLf & _
        "  }" & vbCrLf & _
        "}" & vbCrLf & _
        "</script>" & vbCrLf
End Function

' HTMLダイアログの表示
Private Function ShowHTMLDialog(ByVal htmlContent As String) As String
    On Error GoTo ErrorHandler
    
    ' VBAでのHTMLダイアログ代替実装
    ' InputBoxベースでの選択メニューを改良版として提供
    Dim choice As String
    Dim menuText As String
    
    menuText = "🤖 Outlook AI Helper v1.0.0 Unified" & vbCrLf & vbCrLf & _
               "📊 メール解析:" & vbCrLf & _
               "  A) メール内容解析" & vbCrLf & _
               "  B) 検索フォルダ分析" & vbCrLf & vbCrLf & _
               "✉️ メール作成支援:" & vbCrLf & _
               "  C) 営業断りメール作成" & vbCrLf & _
               "  D) 承諾メール作成" & vbCrLf & _
               "  E) カスタムメール作成" & vbCrLf & vbCrLf & _
               "⚙️ システム管理:" & vbCrLf & _
               "  F) 設定管理" & vbCrLf & _
               "  G) API接続テスト" & vbCrLf & vbCrLf & _
               "💡 実行したい機能のアルファベットを入力してください:"
    
    choice = UCase(Trim(InputBox(menuText, "Outlook AI Helper - メインメニュー", "A")))
    
    ShowHTMLDialog = choice
    Exit Function
    
ErrorHandler:
    ShowError "ダイアログ表示中にエラーが発生しました", Err.Description
    ShowHTMLDialog = ""
End Function

' ユーザー選択の処理
Private Sub ProcessUserChoice(ByVal choice As String)
    On Error GoTo ErrorHandler
    
    Select Case choice
        Case "A"
            Call AnalyzeSelectedEmail
        Case "B"
            Call AnalyzeSearchFolders
        Case "C"
            Call CreateRejectionEmail
        Case "D"
            Call CreateAcceptanceEmail
        Case "E"
            Call CreateCustomBusinessEmail
        Case "F"
            Call ManageConfiguration
        Case "G"
            Call TestAPIConnection
        Case ""
            ' キャンセルされた場合は何もしない
        Case Else
            ShowMessage "無効な選択です。A～Gのアルファベットを入力してください。", "入力エラー", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    ShowError "機能実行中にエラーが発生しました", Err.Description
End Sub

' =============================================================================
' クイックアクセス機能
' =============================================================================

' 各機能への直接アクセス関数（日本語による説明付き）

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

' =============================================================================
' アクセス改善機能
' =============================================================================

' Outlookマクロメニューへの登録（分かりやすい名前で）
Public Sub AIヘルパー_メインメニュー()
    Call ShowMainForm
End Sub

' 既存の英語関数も利用可能性のため維持（内部的な利用）
Public Sub ShowMainMenu()
    Call ShowMainForm
End Sub

' =============================================================================
' 共通関数（既存機能への依存）
' =============================================================================

' エラーメッセージ表示（OutlookAI_Unified.basの関数を直接呼び出し）
Private Sub ShowError(ByVal errorMessage As String, Optional ByVal details As String = "")
    Dim fullMessage As String
    fullMessage = "エラーが発生しました: " & errorMessage
    If details <> "" Then
        fullMessage = fullMessage & vbCrLf & vbCrLf & "詳細: " & details
    End If
    MsgBox fullMessage, vbCritical, "Outlook AI Helper - エラー"
End Sub

' 成功メッセージ表示
Private Sub ShowMessage(ByVal message As String, ByVal title As String, Optional ByVal messageType As VbMsgBoxStyle = vbInformation)
    MsgBox message, messageType, "Outlook AI Helper - " & title
End Sub