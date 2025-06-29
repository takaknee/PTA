# Gmail整理スクリプト - 使用例

## 特定送信者のメール削除機能

### 基本的な使用例

#### 1. 単一送信者の削除
```
メニュー: Gmail整理 > 特定送信者のメールを削除
入力: spam@example.com
結果: 「spam@example.com」から受信したすべてのメールが削除されます
```

#### 2. 複数送信者の削除
```
メニュー: Gmail整理 > 特定送信者のメールを削除
入力: newsletter@company.com, ads@promotion.jp, noreply@marketing.com
結果: 3つの送信者からのメールがすべて削除されます
```

#### 3. 実際の使用シナリオ

**迷惑メール対策**
```
入力: spam@fake-bank.com, phishing@scam.com, fake@lottery.com
効果: 既知の迷惑メール送信者からのメールを一括クリーンアップ
```

**ニュースレッター整理**
```
入力: newsletter@techcrunch.com, updates@medium.com, digest@dev.to
効果: 配信停止したニュースレターの過去メールを削除
```

**プロモーション整理**
```
入力: sale@retailer.com, offers@shopping.com, deals@outlet.com
効果: 不要なプロモーションメールを一括削除
```

### 注意事項

1. **削除は元に戻せません**: 削除前に重要なメールがないか確認してください
2. **メールアドレスの正確性**: 正しいメールアドレスを入力してください
3. **大量削除時の制限**: 一度に大量のメールを削除する場合、Gmail APIの制限により複数回実行が必要な場合があります

### エラー例と対処法

#### 無効なメールアドレス
```
入力: invalid-email
エラー: "無効なメールアドレスが含まれています: invalid-email"
対処: 正しい形式のメールアドレス（@を含む）を入力してください
```

#### 存在しない送信者
```
入力: nonexistent@nowhere.com
結果: "削除対象のメールがこれ以上見つかりません"
説明: 指定した送信者からのメールが存在しない場合の正常な動作です
```

### 処理フロー

1. **入力**: 送信者メールアドレス（カンマ区切り）
2. **検証**: メールアドレス形式のチェック
3. **確認**: 削除対象の確認ダイアログ
4. **実行**: 各送信者からのメールを順次削除
5. **結果**: 削除件数と結果の詳細表示

### パフォーマンス

- **単一送信者**: 通常数秒で完了
- **複数送信者**: 送信者数に応じて時間が増加（1送信者あたり数秒）
- **大量メール**: 500件/回の制限により、複数回実行が必要な場合があります

### 送信元分析機能との組み合わせ

1. **送信元別メール分析**を実行してスプレッドシートに送信者一覧を作成
2. 削除したい送信者のメールアドレスをコピー
3. **特定送信者のメールを削除**でペースト
4. カンマ区切りで複数送信者を指定して一括削除

この組み合わせにより、効率的なメール整理が可能になります。