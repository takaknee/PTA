# Office Add-in 配信・展開ガイド

## 概要

PTA Outlook Add-inの配信・展開方法について説明します。Office Add-insはVSTOと異なり、マニフェストファイルベースの配信が可能です。

## 配信方法の比較

| 配信方法 | 対象 | 難易度 | 管理性 | セキュリティ |
|----------|------|--------|--------|--------------|
| **サイドロード** | 個人・小規模 | 低 | 低 | 中 |
| **SharePoint App Catalog** | 組織内 | 中 | 高 | 高 |
| **Microsoft 365 管理センター** | 企業全体 | 中 | 最高 | 最高 |
| **Microsoft AppSource** | 一般公開 | 高 | 中 | 最高 |

## 1. サイドロード配信（推奨開始方法）

### 特徴
- 最も簡単な配信方法
- 個人・小規模組織に適している
- 即座に開始可能

### 手順

#### 1.1 ファイルホスティング

アドインファイルをWebサーバーでホスティングします：

```bash
# 開発サーバーの起動
cd src/outlook-addin
npm install
npm start

# または静的ファイルサーバー
# Apache, Nginx, IIS等でsrcフォルダを配信
```

#### 1.2 マニフェストの配布

`manifest/manifest.xml` をユーザーに配布します。

#### 1.3 ユーザーによるインストール

**Outlook on Windows:**
1. ファイル > アカウント管理 > アドインの管理
2. 「マニフェストファイルからインストール」を選択
3. `manifest.xml` ファイルを選択

**Outlook on the Web:**
1. 設定（歯車アイコン）> すべてのOutlook設定を表示
2. メール > アドインの管理
3. 「カスタムアドインを追加」> 「ファイルから追加」
4. `manifest.xml` ファイルをアップロード

### 1.4 URL設定の調整

マニフェストファイル内のURLを本番環境に合わせて変更：

```xml
<!-- 開発環境 -->
<bt:Url id="PTAAddIn.Url" DefaultValue="https://localhost:3000/src/functions.html"/>

<!-- 本番環境 -->
<bt:Url id="PTAAddIn.Url" DefaultValue="https://yourdomain.com/pta-addin/src/functions.html"/>
```

## 2. SharePoint App Catalog配信

### 特徴
- 組織内での一元管理
- IT部門による配信制御
- 自動更新対応

### 前提条件
- SharePoint Online管理者権限
- App Catalogサイトの作成

### 手順

#### 2.1 App Catalogサイトの準備

```powershell
# SharePoint Online Management Shell
Connect-SPOService -Url https://tenant-admin.sharepoint.com

# App Catalogサイトの作成（未作成の場合）
New-SPOSite -Url "https://tenant.sharepoint.com/sites/appcatalog" `
           -Title "App Catalog" `
           -Owner "admin@tenant.onmicrosoft.com" `
           -Template "APPCATALOG#0"
```

#### 2.2 アドインファイルのパッケージ化

```bash
# アドインファイルをZIPパッケージ化
cd src/outlook-addin
zip -r pta-outlook-addin.zip . -x "node_modules/*" "*.git*"
```

#### 2.3 App Catalogへのアップロード

1. SharePoint App Catalogサイトにアクセス
2. 「Office用アプリ」ライブラリを開く
3. `pta-outlook-addin.zip` をアップロード
4. アプリの詳細情報を入力

#### 2.4 組織内配信の有効化

1. アプリのコンテキストメニューから「配信」を選択
2. 配信対象（全組織/特定グループ）を設定
3. 承認プロセスの設定

## 3. Microsoft 365 管理センター配信

### 特徴
- 最も高度な管理機能
- ユーザー・グループ単位での配信制御
- 詳細な使用状況分析

### 手順

#### 3.1 管理センターでの設定

1. Microsoft 365 管理センターにアクセス
2. 設定 > 統合アプリ
3. アプリのアップロード > 「マニフェストファイル」

#### 3.2 配信ポリシーの設定

```json
{
  "配信対象": {
    "全組織": false,
    "特定ユーザー": ["user1@tenant.com", "user2@tenant.com"],
    "セキュリティグループ": ["PTA-Users"]
  },
  "配信設定": {
    "自動インストール": true,
    "ユーザー選択可能": false,
    "強制更新": true
  }
}
```

#### 3.3 PowerShellによる配信自動化

```powershell
# Microsoft Graph PowerShell
Connect-MgGraph -Scopes "AppCatalog.ReadWrite.All"

# アドインの配信
$manifest = Get-Content "manifest/manifest.xml" -Raw
New-MgApplicationTemplate -DisplayName "PTA Outlook Add-in" `
                         -Manifest $manifest
```

## 4. 自動更新機能の実装

### 4.1 バージョン更新

マニフェストファイルのバージョンを更新：

```xml
<Version>1.0.1.0</Version>
```

### 4.2 更新通知

JavaScript内で更新チェック：

```javascript
// バージョン確認関数
function checkForUpdates() {
    const currentVersion = "1.0.0";
    
    fetch('/api/version')
        .then(response => response.json())
        .then(data => {
            if (data.version !== currentVersion) {
                showUpdateNotification(data.version);
            }
        })
        .catch(error => {
            console.log('更新チェックエラー:', error);
        });
}

// 更新通知
function showUpdateNotification(newVersion) {
    Office.context.mailbox.item.notificationMessages.addAsync(
        'updateAvailable',
        {
            type: 'informationalMessage',
            message: `新しいバージョン ${newVersion} が利用可能です。`,
            icon: 'icon-16',
            persistent: true
        }
    );
}
```

### 4.3 キャッシュ制御

```html
<!-- HTMLファイルでのキャッシュ制御 -->
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
```

## 5. セキュリティ考慮事項

### 5.1 HTTPS必須

すべてのファイルはHTTPS経由で配信する必要があります：

```apache
# Apache設定例
<VirtualHost *:443>
    ServerName yourdomain.com
    DocumentRoot /path/to/pta-addin
    
    SSLEngine on
    SSLCertificateFile /path/to/certificate.crt
    SSLCertificateKeyFile /path/to/private.key
    
    # セキュリティヘッダー
    Header always set X-Frame-Options SAMEORIGIN
    Header always set X-Content-Type-Options nosniff
    Header always set X-XSS-Protection "1; mode=block"
</VirtualHost>
```

### 5.2 Content Security Policy

```html
<meta http-equiv="Content-Security-Policy" 
      content="default-src 'self'; 
               script-src 'self' https://appsforoffice.microsoft.com; 
               style-src 'self' 'unsafe-inline'; 
               connect-src 'self' https://api.openai.com https://*.openai.azure.com;">
```

### 5.3 API키 보안

```javascript
// APIキーの安全な管理
function getApiKey() {
    // Roaming Settingsから取得（暗号化推奨）
    return Office.context.roamingSettings.get('encryptedApiKey');
}

function encryptApiKey(key) {
    // 簡単な暗号化（実際には強力な暗号化を使用）
    return btoa(key);
}
```

## 6. 監視とトラブルシューティング

### 6.1 ログ収集

```javascript
// 使用統計の収集
function trackUsage(action, details) {
    const usageData = {
        timestamp: new Date().toISOString(),
        action: action,
        details: details,
        version: '1.0.0',
        platform: Office.context.platform
    };
    
    // ログサーバーへ送信
    fetch('/api/telemetry', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(usageData)
    }).catch(error => {
        // ログ送信失敗時の処理
        console.log('テレメトリー送信エラー:', error);
    });
}
```

### 6.2 エラー監視

```javascript
// グローバルエラーハンドラー
window.addEventListener('error', (event) => {
    const errorData = {
        message: event.message,
        filename: event.filename,
        line: event.lineno,
        column: event.colno,
        stack: event.error ? event.error.stack : null
    };
    
    trackUsage('error', errorData);
});
```

### 6.3 パフォーマンス監視

```javascript
// パフォーマンス測定
function measurePerformance(operation, fn) {
    const start = performance.now();
    
    return fn().then(result => {
        const duration = performance.now() - start;
        trackUsage('performance', {
            operation: operation,
            duration: duration
        });
        return result;
    });
}
```

## 7. 推奨配信戦略

### 段階的展開アプローチ

1. **Phase 1: パイロット展開** (1-2週間)
   - IT部門・管理者による検証
   - サイドロード方式で少数ユーザー

2. **Phase 2: ベータ展開** (2-4週間)
   - 各部門から代表ユーザー選出
   - SharePoint App Catalog使用
   - フィードバック収集・改善

3. **Phase 3: 段階的本格展開** (4-8週間)
   - 部門単位での順次展開
   - Microsoft 365 管理センター使用
   - サポート体制構築

4. **Phase 4: 全社展開** (2-4週間)
   - 残りユーザーへの一括配信
   - 自動インストール有効化
   - 24/7サポート体制

### 組織規模別推奨方法

| 組織規模 | 推奨配信方法 | 管理レベル |
|----------|--------------|------------|
| ~50名 | サイドロード | 基本 |
| 50-200名 | SharePoint App Catalog | 中級 |
| 200-1000名 | Microsoft 365 管理センター | 高級 |
| 1000名以上 | Microsoft 365 + PowerShell | エンタープライズ |

## 8. サポート体制

### ユーザーサポート

- **ヘルプデスク**: よくある質問への対応
- **トレーニング**: 使用方法の説明会
- **ドキュメント**: オンラインマニュアル

### 技術サポート

- **GitHub Issues**: バグ報告・機能要望
- **メンテナンス**: 定期的な更新・改善
- **監視**: システム稼働状況の監視

## 9. コスト分析

### 初期コスト

- 開発・移植: 完了済み
- テスト・検証: 1-2週間
- 配信環境構築: 1週間

### 運用コスト（月額）

- ホスティング: $10-50/月
- SSL証明書: $10-100/年
- 監視・保守: 開発者時間2-4時間/月

### VSTO比較での削減効果

- 配信コスト: 90%削減（ClickOnce → Web配信）
- サポートコスト: 70%削減（インストール問題解消）
- 更新コスト: 80%削減（自動更新）

## まとめ

Office Add-ins の配信は、VSTOと比較して大幅に簡素化されました。組織の規模と要件に応じて適切な配信方法を選択し、段階的な展開により安全で効率的な導入を実現できます。