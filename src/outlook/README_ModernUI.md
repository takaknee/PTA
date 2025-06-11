# Outlook AI Helper - モダンUI版

## 概要

Outlook AI Helper にモダンなHTMLベースのUIを追加した改良版です。従来のInputBoxベースの操作から、視覚的で直感的なボタンベースの操作に進化しました。

## 新機能

### 🚀 モダンHTMLダイアログUI
- **視覚的操作**: ボタンクリックで機能を実行
- **カテゴリ分け**: メール解析、メール作成、システム管理でグループ化
- **レスポンシブデザイン**: 美しいグラデーションと現代的なボタンデザイン
- **日本語完全対応**: すべての操作が日本語で表示

### 🔄 フォールバック機能
- **自動フォールバック**: HTMLダイアログが失敗時に改良版メニューを表示
- **後方互換性**: 既存の`ShowMainMenu()`も引き続き利用可能
- **段階的エラー処理**: 複数レベルのフォールバック機能

## インストール方法

### 新規インストール

1. **OutlookAI_Unified.bas** をVBAプロジェクトにインポート
2. **OutlookAI_MainForm.bas** をVBAプロジェクトにインポート
3. API設定を完了（OPENAI_API_KEY, OPENAI_API_ENDPOINT）
4. `AIヘルパー_統合メニュー` を実行

### 既存環境のアップグレード

1. **OutlookAI_Unified.bas** を最新版に更新
2. **OutlookAI_MainForm.bas** を新規追加
3. 既存の使い方はそのまま利用可能

## 使用方法

### 🎯 推奨: モダンUI
VBAエディタで以下の関数を実行：

```vba
AIヘルパー_統合メニュー
```

または日本語名：

```vba
統合メニュー
```

### 📋 フォールバック: 改良版メニュー
HTMLダイアログが利用できない場合、自動的に改良版メニューが表示されます：

```vba
ShowEnhancedMainMenu
```

### 🔧 従来の方法（後方互換性）
既存の方法もそのまま利用できます：

```vba
ShowMainMenu
```

## 機能一覧

### 📊 メール解析
- **📧 メール解析**: 選択されたメールをAIで分析
- **📁 フォルダ分析**: 受信トレイの整理状況を分析

### ✉️ メール作成
- **❌ 営業断り**: 丁寧な営業断りメールを生成
- **✅ 承諾メール**: ビジネス承諾メールを生成
- **📝 カスタムメール**: カスタムプロンプトでメール生成

### ⚙️ システム管理
- **⚙️ 設定**: API設定の確認と管理
- **🔌 API テスト**: OpenAI API接続テスト
- **ℹ️ バージョン**: バージョン情報表示

## 技術仕様

### HTMLダイアログUI
- **レンダリング**: Internet Explorer オブジェクト使用
- **スタイル**: CSS3 グラデーション、アニメーション対応
- **スクリプト**: JavaScript による関数呼び出し
- **レスポンシブ**: 固定サイズ（600x700px）、中央配置

### フォールバック階層
```
1. HTMLダイアログUI (ShowMainForm)
   ↓ 失敗時
2. 改良版メニュー (ShowEnhancedMainMenu)
   ↓ 失敗時
3. 従来メニュー (ShowMainMenu)
```

## トラブルシューティング

### Q: 「OutlookAI_MainForm.bas が見つからない」エラー

**A:** フォールバック機能により改良版メニューが表示されます。  
新しいUIを使用したい場合は OutlookAI_MainForm.bas をインポートしてください。

### Q: HTMLダイアログが表示されない

**A:** 以下を確認してください：
- Internet Explorer が正常に動作するか
- VBAマクロのセキュリティ設定が適切か
- OutlookAI_MainForm.bas が正しくインポートされているか

### Q: 従来の方法が使えなくなった

**A:** 後方互換性は完全に保持されています：
- `ShowMainMenu()` は引き続き利用可能
- すべての既存関数名（英語）も動作します
- 既存のワークフローに影響はありません

### Q: 日本語関数名が認識されない

**A:** 以下を確認してください：
- OutlookAI_Unified.bas が最新版に更新されているか
- VBAプロジェクトに正しくインポートされているか
- マクロのセキュリティ設定が適切か

## 開発者向け情報

### アーキテクチャ

```
OutlookAI_Unified.bas          # 既存機能 + 統合UI呼び出し
├── 従来の英語関数名            # AnalyzeSelectedEmail等
├── 日本語エイリアス関数        # メール内容解析等
├── 改良版メニュー             # ShowEnhancedMainMenu
├── 統合UI呼び出し             # AIヘルパー_統合メニュー
└── 検索フォルダ分析           # AnalyzeSearchFolders (新規)

OutlookAI_MainForm.bas         # 新規追加
├── 統合フォームUI             # ShowMainForm
├── HTMLダイアログ            # CreateMainFormHTML
├── CSS/JavaScriptコード       # CreateMainFormCSS等
└── フォールバック処理          # ShowEnhancedMainMenu呼び出し
```

### 実装方針
- **最小変更**: 既存コードへの影響を最小限に抑制
- **段階的フォールバック**: 複数レベルのエラー処理
- **後方互換性**: 既存機能の完全保持
- **日本語対応**: UI、エラーメッセージ、ドキュメントの完全日本語化

## テスト方法

以下のテスト関数を実行してください：

```vba
' 基本機能テスト
TestBasicFunctions

' モダンUI機能テスト
TestModernUIFunctions

' MainForm可用性テスト
TestMainFormAvailability
```

## 変更履歴

### v1.0.0 Modern UI (2024)
- HTMLベースのモダンUI追加
- 統合メニュー機能実装
- フォールバック機能追加
- 検索フォルダ分析機能追加
- 改良版メニュー追加
- 日本語エイリアス関数追加
- 完全な後方互換性維持

---

## サポート

- 技術的な問題: GitHub Issues
- 使用方法: このREADMEファイル
- API設定: OutlookAI_Unified.bas内のコメント参照