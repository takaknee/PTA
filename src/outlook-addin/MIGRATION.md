# VSTOからOffice Add-insへの移行ガイド

## 移行の背景

Microsoft VSTOは終了予定の技術であり、今後のサポートや新機能の提供は期待できません。一方、Office Add-ins（旧称Office Web Add-ins）は現在のMicrosoftが推進する標準的なOffice拡張技術です。

### 移行する理由

1. **技術の将来性**: VSTOはレガシー技術、Office Add-insは現役・将来技術
2. **クロスプラットフォーム対応**: Windows以外のMac、Web版Outlookでも動作
3. **配信の簡素化**: ClickOnceやMSIが不要、マニフェストファイルのみ
4. **開発・保守の容易さ**: Web技術（HTML/CSS/JavaScript）による開発
5. **セキュリティ**: サンドボックス環境での実行によるセキュリティ向上

## 技術比較

| 項目 | VSTO | Office Add-ins |
|------|------|----------------|
| **開発言語** | C#/VB.NET | JavaScript/TypeScript/HTML/CSS |
| **実行環境** | .NET Framework | ブラウザエンジン |
| **配信方法** | ClickOnce/MSI | マニフェストファイル |
| **対応プラットフォーム** | Windowsのみ | Windows/Mac/Web |
| **API** | Outlook Object Model | Office JavaScript API |
| **UI統合** | WinForms/WPF | HTML/CSS |
| **デバッグ** | Visual Studio | ブラウザ開発者ツール |
| **更新方法** | 再インストール | 自動更新 |
| **セキュリティモデル** | フルトラスト | サンドボックス |

## 機能移行マッピング

### 1. 基本機能

| VSTO機能 | Office Add-ins移行先 |
|----------|---------------------|
| `ThisAddIn.Application` | `Office.context.mailbox` |
| `Application_ItemSend` | Office.EventType.ItemSend |
| `Ribbon UI` | VersionOverrides/ExtensionPoint |
| `Task Pane` | TaskPane Add-in |
| `Form Regions` | Content Add-in |

### 2. メール操作

```csharp
// VSTO
var mailItem = Application.ActiveInspector().CurrentItem as MailItem;
string subject = mailItem.Subject;
string body = mailItem.Body;

// Office Add-ins
Office.context.mailbox.item.subject.getAsync((result) => {
    const subject = result.value;
});

Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
    const body = result.value;
});
```

### 3. 設定管理

```csharp
// VSTO
Properties.Settings.Default.ApiKey = "key";
Properties.Settings.Default.Save();

// Office Add-ins
Office.context.roamingSettings.set('apiKey', 'key');
Office.context.roamingSettings.saveAsync();
```

### 4. HTTP通信

```csharp
// VSTO
using (var client = new HttpClient())
{
    var response = await client.PostAsync(url, content);
    return await response.Content.ReadAsStringAsync();
}

// Office Add-ins
const response = await fetch(url, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(data)
});
const result = await response.json();
```

## 段階的移行計画

### Phase 1: 準備フェーズ（完了）

- [x] Office Add-ins プロジェクト構造作成
- [x] マニフェストファイル作成
- [x] 基本HTMLページ作成
- [x] CSS スタイル定義

### Phase 2: 基本機能移植（完了）

- [x] メール読み取り機能
- [x] メール作成支援機能
- [x] タスクペイン実装
- [x] リボンボタン統合

### Phase 3: 高度機能移植

- [x] AI API統合（OpenAI/Azure OpenAI）
- [x] 設定管理機能
- [x] 履歴管理機能
- [x] エラーハンドリング

### Phase 4: 検証・最適化

- [ ] 機能テスト
- [ ] パフォーマンス最適化
- [ ] UI/UX改善
- [ ] ドキュメント整備

### Phase 5: 配信準備

- [ ] 配信戦略策定
- [ ] ユーザーガイド作成
- [ ] 移行手順書作成
- [ ] サポート体制構築

## 実装上の注意点

### 1. 非同期処理

Office Add-insはすべて非同期APIです：

```javascript
// 正しい実装
Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
        const body = result.value;
        // 処理続行
    }
});

// またはPromiseベース
function getEmailBody() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}
```

### 2. セキュリティ制限

Office Add-insはサンドボックス環境で実行されるため：

- ローカルファイルシステムへの直接アクセス不可
- レジストリアクセス不可
- 他のアプリケーションとの直接通信不可
- HTTPS必須（開発時を除く）

### 3. APIの制限

一部のVSTO機能は制限されます：

- Outlookオブジェクトモデルのフルアクセス不可
- イベントハンドリングの制限
- UI統合の制限（リボン/タスクペインのみ）

## テスト戦略

### 1. 機能テスト

```javascript
// テスト用の関数例
async function testEmailAnalysis() {
    try {
        const emailData = await getTestEmailData();
        const analysis = await performAIAnalysis(emailData);
        
        console.log('テスト成功:', analysis);
        return true;
    } catch (error) {
        console.error('テスト失敗:', error);
        return false;
    }
}
```

### 2. プラットフォーム横断テスト

- Windows Outlook Desktop
- Mac Outlook
- Outlook on the Web
- 各ブラウザでの動作確認

### 3. API統合テスト

- OpenAI API接続テスト
- Azure OpenAI接続テスト
- エラーケース処理テスト

## 配信戦略

### 1. 段階的ロールアウト

1. **パイロット展開**: 小規模グループでの検証
2. **ベータ展開**: 限定ユーザーでの実用テスト
3. **本格展開**: 全ユーザーへの配信

### 2. 配信方法

#### サイドロード（推奨）

```xml
<!-- manifest.xml をOutlookにインポート -->
<OfficeApp>
  <!-- マニフェスト設定 -->
</OfficeApp>
```

#### SharePoint App Catalog

1. SharePoint管理センターでApp Catalogを有効化
2. manifest.xmlをアップロード
3. 組織内で配信

#### Microsoft 365 管理センター

1. 統合アプリメニューから追加
2. カスタムアプリとしてマニフェストをアップロード
3. ユーザー/グループに配信

## ユーザー向け移行ガイド

### VSTO版からの変更点

1. **インストール方法**: EXEファイルからマニフェストファイルへ
2. **UI**: WinFormsからHTML/CSSベースへ
3. **設定場所**: レジストリからクラウド設定へ
4. **更新方法**: 手動更新から自動更新へ

### 操作方法の変更

基本的な操作は同じですが、以下の点が変更されます：

- 設定画面のデザイン
- エラーメッセージの表示方法
- タスクペインの表示位置

### データ移行

VSTO版で保存していた設定は手動で再設定が必要です：

1. APIキーの再入力
2. カスタムプロンプトの再設定
3. プリファレンスの再選択

## トラブルシューティング

### よくある問題

1. **マニフェストエラー**: XMLの形式確認、URLアクセス確認
2. **HTTPS証明書エラー**: 開発証明書のインストール
3. **API接続エラー**: CORS設定、APIキーの確認
4. **パフォーマンス**: 大量データ処理の最適化

### サポート

- 技術的な問題: GitHub Issues
- 使用方法: README.md
- 移行支援: 段階的サポート計画

## 今後のロードマップ

### 短期（1-3ヶ月）

- [ ] 全機能の動作確認
- [ ] パフォーマンス最適化
- [ ] ユーザーフィードバック収集

### 中期（3-6ヶ月）

- [ ] 追加AI プロバイダー対応
- [ ] 高度な解析機能
- [ ] 大規模組織向け機能

### 長期（6ヶ月以降）

- [ ] Microsoft Teams連携
- [ ] SharePoint統合
- [ ] 多言語対応

## 結論

Office Add-ins への移行により、以下の利益が期待できます：

1. **将来への対応**: 最新のMicrosoft技術への対応
2. **利用範囲の拡大**: クロスプラットフォーム対応
3. **保守性の向上**: Web技術による開発・保守の容易さ
4. **配信の簡略化**: マニフェストベースの簡単配信
5. **セキュリティの向上**: サンドボックス実行環境

移行作業は段階的に実施し、ユーザーへの影響を最小限に抑えながら進めていきます。