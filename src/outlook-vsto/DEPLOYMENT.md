# VSTO配信・展開ガイド

## 概要

PTA Outlook アドイン（VSTO版）のM365環境における配信方法と自動更新機能について説明します。

## 配信方法の比較

### 1. ClickOnce配信（推奨）

#### 利点
- ✅ **簡単な展開**: Webサーバーまたはファイル共有からの配信
- ✅ **自動更新**: アプリケーション起動時の自動更新チェック
- ✅ **ユーザー権限**: 管理者権限不要でインストール可能
- ✅ **ロールバック**: 以前のバージョンへの復元機能
- ✅ **部分更新**: 変更された部分のみダウンロード

#### 実装手順
```xml
<!-- プロジェクトファイルでの設定 -->
<PropertyGroup>
  <PublishUrl>\\shared-server\outlook-addin\</PublishUrl>
  <InstallUrl>\\shared-server\outlook-addin\</InstallUrl>
  <UpdateEnabled>true</UpdateEnabled>
  <UpdateMode>Foreground</UpdateMode>
  <UpdateInterval>7</UpdateInterval>
  <UpdateIntervalUnits>Days</UpdateIntervalUnits>
  <ApplicationRevision>1</ApplicationRevision>
  <ApplicationVersion>1.0.0.%2a</ApplicationVersion>
</PropertyGroup>
```

#### 配信手順
1. Visual Studioで「発行」を実行
2. 共有フォルダまたはWebサーバーに配置
3. ユーザーにsetup.exeまたは.applicationファイルを配布
4. ユーザーがインストールを実行

### 2. Microsoft Store for Business配信

#### 利点
- ✅ **企業向け配信**: IT部門による一括管理
- ✅ **セキュリティ**: Microsoft認証による安全性
- ✅ **自動更新**: Store経由での自動更新
- ✅ **ライセンス管理**: 使用状況の追跡・管理

#### 制限事項
- ❌ **審査プロセス**: Microsoft審査に時間が必要
- ❌ **パッケージング**: MSIXパッケージ形式への変換が必要
- ❌ **企業ストア設定**: Microsoft Store for Businessの設定が必要

### 3. Group Policy配信

#### 利点
- ✅ **強制展開**: 管理者による一括強制インストール
- ✅ **ポリシー統合**: 既存のActive Directoryポリシーと統合
- ✅ **一元管理**: 大規模組織での効率的な展開

#### 実装手順
1. MSIパッケージの作成
2. GPOでのソフトウェアインストールポリシー設定
3. 対象ユーザー・PCグループへの適用

#### GPO設定例
```
コンピューターの構成 > ポリシー > ソフトウェアの設定 > ソフトウェアのインストール
- 新規 > パッケージ
- MSIファイルを指定
- 「割り当て済み」または「発行済み」を選択
```

### 4. System Center Configuration Manager (SCCM)

#### 利点
- ✅ **エンタープライズ配信**: 大規模組織向け
- ✅ **詳細制御**: 配信スケジュール・対象の細かな制御
- ✅ **レポート機能**: インストール状況の詳細な追跡

## 自動更新機能の実装

### ClickOnceベースの自動更新

#### C#での実装例
```csharp
using System.Deployment.Application;

public class AutoUpdateService
{
    public async Task CheckForUpdatesAsync()
    {
        if (ApplicationDeployment.IsNetworkDeployed)
        {
            ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;
            
            try
            {
                UpdateCheckInfo info = await Task.Run(() => ad.CheckForDetailedUpdate());
                
                if (info.UpdateAvailable)
                {
                    if (info.IsUpdateRequired)
                    {
                        // 必須更新の場合は強制更新
                        await PerformMandatoryUpdateAsync(ad);
                    }
                    else
                    {
                        // オプション更新の場合はユーザーに確認
                        await PromptOptionalUpdateAsync(ad, info);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"更新チェック中にエラーが発生しました: {ex.Message}");
            }
        }
    }

    private async Task PerformMandatoryUpdateAsync(ApplicationDeployment ad)
    {
        try
        {
            await Task.Run(() => ad.Update());
            MessageBox.Show("アプリケーションが更新されました。再起動します。");
            Application.Restart();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"更新に失敗しました: {ex.Message}");
        }
    }

    private async Task PromptOptionalUpdateAsync(ApplicationDeployment ad, UpdateCheckInfo info)
    {
        var result = MessageBox.Show(
            $"新しいバージョン（{info.AvailableVersion}）が利用可能です。\\n更新しますか？",
            "アップデート",
            MessageBoxButtons.YesNo
        );

        if (result == DialogResult.Yes)
        {
            try
            {
                await Task.Run(() => ad.Update());
                MessageBox.Show("更新が完了しました。");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新に失敗しました: {ex.Message}");
            }
        }
    }
}
```

#### アドイン統合
```csharp
// ThisAddIn.cs での実装
private AutoUpdateService _updateService;

private async void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    // 他の初期化処理...
    
    _updateService = new AutoUpdateService();
    
    // 起動時に更新チェック（バックグラウンドで実行）
    _ = Task.Run(async () =>
    {
        await Task.Delay(5000); // 5秒待機してから更新チェック
        await _updateService.CheckForUpdatesAsync();
    });
}
```

### 設定ベースの更新制御

#### App.configでの設定
```xml
<appSettings>
  <add key="EnableAutoUpdate" value="true" />
  <add key="UpdateCheckInterval" value="7" />
  <add key="UpdateUrl" value="\\\\server\\share\\outlook-addin\\" />
  <add key="ForceUpdateVersion" value="1.0.1" />
</appSettings>
```

#### 設定読み込み実装
```csharp
public class UpdateConfiguration
{
    public bool EnableAutoUpdate => 
        bool.Parse(ConfigurationManager.AppSettings["EnableAutoUpdate"] ?? "true");
    
    public int CheckIntervalDays => 
        int.Parse(ConfigurationManager.AppSettings["UpdateCheckInterval"] ?? "7");
    
    public string UpdateUrl => 
        ConfigurationManager.AppSettings["UpdateUrl"] ?? "";
    
    public Version ForceUpdateVersion
    {
        get
        {
            var versionString = ConfigurationManager.AppSettings["ForceUpdateVersion"];
            return string.IsNullOrEmpty(versionString) ? null : new Version(versionString);
        }
    }
}
```

## 配信用ファイル構成

### ClickOnce配信用ファイル
```
publish/
├── Application Files/
│   └── OutlookPTAAddin_1_0_0_1/
│       ├── OutlookPTAAddin.exe.deploy
│       ├── OutlookPTAAddin.exe.manifest
│       └── 依存関係ファイル群...
├── OutlookPTAAddin.application      # アプリケーションマニフェスト
└── setup.exe                        # セットアップブートストラッパー
```

### MSI配信用ファイル
```
installer/
├── OutlookPTAAddin.msi              # MSIインストーラー
├── setup.exe                        # セットアップブートストラッパー
└── prerequisites/                   # 前提条件ファイル
    ├── .NET Framework 4.8/
    └── Visual Studio Tools for Office Runtime/
```

## セキュリティ考慮事項

### コード署名
- **重要**: ClickOnceアプリケーションには必ずコード署名を適用
- **証明書**: 企業証明書またはパブリック証明書を使用
- **信頼レベル**: FullTrustでの実行が必要

### ネットワークセキュリティ
- **HTTPS**: Web配信時はHTTPS必須
- **ファイアウォール**: 必要なポート（80/443）の開放
- **プロキシ**: 企業プロキシ環境での動作確認

### 権限設定
- **ユーザー権限**: 一般ユーザー権限でのインストール・実行
- **レジストリ**: 必要最小限のレジストリアクセス
- **ファイルシステム**: ユーザープロファイル内でのデータ保存

## トラブルシューティング

### よくある問題と解決方法

#### 1. ClickOnce配信エラー
```
エラー: "アプリケーションの要件を満たしていません"
解決策: 
- .NET Framework 4.8の確認・インストール
- VSTOランタイムの確認・インストール
- セキュリティ設定の確認
```

#### 2. 自動更新失敗
```
エラー: "更新サーバーに接続できません"
解決策:
- ネットワーク接続の確認
- プロキシ設定の確認
- ファイアウォール設定の確認
```

#### 3. インストール権限エラー
```
エラー: "管理者権限が必要です"
解決策:
- ClickOnceでの配信に変更
- インストール場所をユーザープロファイルに変更
- MSI署名の確認
```

## 推奨配信戦略

### フェーズ1: パイロット配信
1. **小規模テスト**: IT部門内での試験運用
2. **フィードバック収集**: 使用感・問題点の把握
3. **設定調整**: 環境に応じた設定の最適化

### フェーズ2: 段階的展開
1. **部門別展開**: 部門ごとの段階的ロールアウト
2. **サポート体制**: ヘルプデスク対応の準備
3. **監視・追跡**: インストール状況の監視

### フェーズ3: 全体展開
1. **一括配信**: 全対象ユーザーへの配信
2. **自動更新有効化**: 継続的な更新システムの稼働
3. **運用保守**: 定期的なメンテナンス・更新

## まとめ

M365環境でのVSTO配信において、**ClickOnce配信**が最も簡単で効果的な方法です。自動更新機能と組み合わせることで、継続的な運用・保守が可能になります。

大規模組織の場合は、Group PolicyやSCCMとの併用も検討し、組織の規模と要件に応じて最適な配信戦略を選択してください。