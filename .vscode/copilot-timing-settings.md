# VS Code GitHub Copilot タイミング制御設定 解説

## 追加された設定の効果

### 1. 提案表示のタイミング制御

```json
"github.copilot.advanced": {
    "inlineSuggestCount": 3,
    "debounceMs": 75,        // 75ms後に提案を表示（デフォルト: 100ms）
    "listCount": 10
}
```

- **`debounceMs`**: キー入力停止から提案表示までの待機時間
- より短い時間（75ms）で素早く提案が表示される

### 2. 自動提案とインライン表示

```json
"editor.quickSuggestionsDelay": 100,  // 100ms後に候補を表示
"editor.inlineSuggest.enabled": true, // インライン提案を有効化
"editor.inlineSuggest.showToolbar": "onHover" // ホバー時にツールバー表示
```

### 3. 自動受け入れ設定

```json
"editor.acceptSuggestionOnCommitCharacter": true, // セミコロンなどで自動受け入れ
"editor.acceptSuggestionOnEnter": "on",          // Enterキーで受け入れ
"editor.tabCompletion": "on"                     // Tabキーで補完
```

### 4. 自動保存連携

```json
"files.autoSave": "afterDelay",  // 遅延後に自動保存
"files.autoSaveDelay": 1000     // 1秒（1000ms）後に保存
```

## 実際の動作

### タイミングフロー

1. **キー入力停止** → **75ms待機** → **Copilot提案表示**
2. **提案表示** → **ユーザー操作待機**
3. **Tab/Enter** → **提案受け入れ** → **1秒後に自動保存**

### キーボードショートカット

- **Tab**: Copilot提案を受け入れ
- **Enter**: 提案を受け入れて改行
- **Ctrl+→**: 提案の一部のみ受け入れ
- **Esc**: 提案をキャンセル

### カスタマイズ例

さらに細かい制御が必要な場合は以下の設定も調整可能：

```json
{
    // より積極的な提案（短時間）
    "github.copilot.advanced": {
        "debounceMs": 50,
        "inlineSuggestCount": 5
    },
    "editor.quickSuggestionsDelay": 50,
    
    // より保守的な提案（長時間）
    "github.copilot.advanced": {
        "debounceMs": 200,
        "inlineSuggestCount": 1
    },
    "editor.quickSuggestionsDelay": 300
}
```

## 効果的な使用方法

1. **コードを書き始める** → 少し待つ（75ms）→ **提案が表示**
2. **良い提案の場合** → **Tab** で即座に受け入れ
3. **部分的に良い場合** → **Ctrl+→** で単語単位で受け入れ
4. **不要な場合** → **Esc** でキャンセルして続行

これらの設定により、Copilotの提案が適切なタイミングで表示され、効率的なコーディングが可能になります。
