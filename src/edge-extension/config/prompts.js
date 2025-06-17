/**
 * AI機能のプロンプト設定ファイル
 * Copyright (c) 2024 AI Development Team
 * ES5互換形式（Service Worker対応）
 */

// プロンプト設定管理オブジェクト
var PromptManager = {
  // VSCode設定解析用プロンプト
  VSCODE_ANALYSIS: {
    template: "あなたはVSCode設定の専門家です。以下のVSCodeドキュメントページから設定項目を抽出し、HTML構造で日本語解説を提供してください。\n\n" +
      "ページタイトル: {{pageTitle}}\n" +
      "ページURL: {{pageUrl}}\n" +
      "ページ内容: {{pageContent}}\n\n" +
      "以下のHTML構造で回答してください（HTMLタグのみで、マークダウンは使用しないでください）：\n\n" +
      "<div class=\"vscode-analysis-content\">\n" +
      "    <div class=\"analysis-section\">\n" +
      "        <h3>📋 設定項目一覧</h3>\n" +
      "        <div class=\"settings-group\">\n" +
      "            <h4>主要設定</h4>\n" +
      "            <div class=\"setting-item\">\n" +
      "                <strong class=\"setting-name\">設定名</strong>\n" +
      "                <code class=\"setting-value\">設定値の例</code>\n" +
      "                <p class=\"setting-description\">設定の説明</p>\n" +
      "            </div>\n" +
      "        </div>\n" +
      "        <div class=\"settings-group\">\n" +
      "            <h4>追加設定（オプション）</h4>\n" +
      "            <div class=\"setting-item\">\n" +
      "                <strong class=\"setting-name\">設定名</strong>\n" +
      "                <code class=\"setting-value\">設定値の例</code>\n" +
      "                <p class=\"setting-description\">設定の説明</p>\n" +
      "            </div>\n" +
      "        </div>\n" +
      "    </div>\n" +
      "    \n" +
      "    <div class=\"analysis-section\">\n" +
      "        <h3>🛠️ サンプル設定ファイル (settings.json)</h3>\n" +
      "        <pre class=\"settings-json\"><code>{\n" +
      "    // 抽出された設定項目のJSON例\n" +
      "}</code></pre>\n" +
      "    </div>\n" +
      "    \n" +
      "    <div class=\"analysis-section\">\n" +
      "        <h3>💡 使用方法</h3>\n" +
      "        <ol class=\"usage-steps\">\n" +
      "            <li>手順1の詳細説明</li>\n" +
      "            <li>手順2の詳細説明</li>\n" +
      "            <li>手順3の詳細説明</li>\n" +
      "        </ol>\n" +
      "    </div>\n" +
      "    \n" +
      "    <div class=\"analysis-section\">\n" +
      "        <h3>⚠️ 注意点</h3>\n" +
      "        <ul class=\"warnings-list\">\n" +
      "            <li>注意点1の詳細</li>\n" +
      "            <li>注意点2の詳細</li>\n" +
      "        </ul>\n" +
      "    </div>\n" +
      "</div>\n\n" +
      "重要: 必ずHTML構造で回答し、マークダウン記法は使用しないでください。VSCodeドキュメントの内容に基づいて、実用的で分かりやすい設定解説をHTML形式で提供してください。",

    build: function (data) {
      var prompt = this.template;
      prompt = prompt.replace('{{pageTitle}}', data.pageTitle || '');
      prompt = prompt.replace('{{pageUrl}}', data.pageUrl || '');
      prompt = prompt.replace('{{pageContent}}', data.pageContent || '');
      return prompt;
    }
  },

  // メール解析用プロンプト
  EMAIL_ANALYSIS: {
    template: "以下のメール内容を分析し、PTA活動に関連する重要な情報を日本語で整理してください。\n\n" +
      "送信者: {{sender}}\n" +
      "件名: {{subject}}\n" +
      "本文: {{content}}\n\n" +
      "以下の形式で回答してください：\n\n" +
      "## 📧 メール概要\n" +
      "- **種類**: [連絡事項/イベント告知/緊急連絡/その他]\n" +
      "- **重要度**: [高/中/低]\n" +
      "- **対象**: [全保護者/特定学年/役員/その他]\n\n" +
      "## 📋 主要な内容\n" +
      "1. [要点1]\n" +
      "2. [要点2]\n" +
      "3. [要点3]\n\n" +
      "## 📅 日程・期限\n" +
      "- [関連する日程や期限があれば記載]\n\n" +
      "## 🎯 必要なアクション\n" +
      "- [保護者が取るべき行動があれば記載]\n\n" +
      "## 💡 補足情報\n" +
      "- [その他の重要な情報や注意点]",

    build: function (data) {
      var prompt = this.template;
      prompt = prompt.replace('{{sender}}', data.sender || '');
      prompt = prompt.replace('{{subject}}', data.subject || '');
      prompt = prompt.replace('{{content}}', data.content || '');
      return prompt;
    }
  },

  // メール返信作成用プロンプト
  EMAIL_COMPOSE: {
    template: "以下のメール内容に対する適切な返信を日本語で作成してください。\n\n" +
      "元メール:\n" +
      "送信者: {{sender}}\n" +
      "件名: {{subject}}\n" +
      "本文: {{content}}\n\n" +
      "返信要件: {{requirements}}\n\n" +
      "以下の形式で返信メールを作成してください：\n\n" +
      "件名: {{replySubject}}\n\n" +
      "{{senderName}}様\n\n" +
      "いつもお世話になっております。\n\n" +
      "[返信内容をここに記載]\n\n" +
      "何かご不明な点がございましたら、お気軽にお声がけください。\n\n" +
      "よろしくお願いいたします。\n\n" +
      "[あなたの名前]",

    build: function (data) {
      var prompt = this.template;
      prompt = prompt.replace('{{sender}}', data.sender || '');
      prompt = prompt.replace('{{subject}}', data.subject || '');
      prompt = prompt.replace('{{content}}', data.content || '');
      prompt = prompt.replace('{{requirements}}', data.requirements || '丁寧で適切な返信');
      prompt = prompt.replace('{{replySubject}}', data.replySubject || 'Re: ' + (data.subject || ''));
      prompt = prompt.replace('{{senderName}}', data.senderName || data.sender || '');
      return prompt;
    }
  },

  // ページ解析用プロンプト
  PAGE_ANALYSIS: {
    template: "以下のWebページの内容を分析し、重要な情報を日本語で整理してください。\n\n" +
      "ページタイトル: {{pageTitle}}\n" +
      "ページURL: {{pageUrl}}\n" +
      "ページ内容: {{pageContent}}\n\n" +
      "以下の形式で回答してください：\n\n" +
      "## 📄 ページ概要\n" +
      "- **種類**: [ニュース/資料/手順書/その他]\n" +
      "- **主なトピック**: [メインテーマ]\n\n" +
      "## 📋 重要なポイント\n" +
      "1. [要点1]\n" +
      "2. [要点2]\n" +
      "3. [要点3]\n\n" +
      "## 🔗 関連リンク・参考情報\n" +
      "- [重要なリンクや参考情報があれば記載]\n\n" +
      "## 💡 活用方法\n" +
      "- [この情報をどのように活用できるか]",

    build: function (data) {
      var prompt = this.template;
      prompt = prompt.replace('{{pageTitle}}', data.pageTitle || '');
      prompt = prompt.replace('{{pageUrl}}', data.pageUrl || '');
      prompt = prompt.replace('{{pageContent}}', data.pageContent || '');
      return prompt;
    }
  },

  // 選択テキスト解析用プロンプト
  SELECTION_ANALYSIS: {
    template: "以下の選択されたテキストを分析し、重要な情報を日本語で整理してください。\n\n" +
      "選択テキスト: {{selectedText}}\n" +
      "コンテキスト: {{context}}\n\n" +
      "以下の形式で回答してください：\n\n" +
      "## 📝 テキスト概要\n" +
      "- **内容の種類**: [情報/指示/質問/その他]\n" +
      "- **重要度**: [高/中/低]\n\n" +
      "## 📋 主要な内容\n" +
      "1. [要点1]\n" +
      "2. [要点2]\n" +
      "3. [要点3]\n\n" +
      "## 💡 解釈・補足\n" +
      "- [テキストの意味や背景の説明]\n\n" +
      "## 🎯 推奨アクション\n" +
      "- [このテキストに基づいて取るべき行動があれば記載]",

    build: function (data) {
      var prompt = this.template;
      prompt = prompt.replace('{{selectedText}}', data.selectedText || '');
      prompt = prompt.replace('{{context}}', data.context || '');
      return prompt;
    }
  },

  // Teams転送用コンテンツ生成
  TEAMS_CONTENT: {
    template: "以下の情報をTeams転送用に適切な形式で整理してください：\n\n" +
      "タイトル: {{title}}\n" +
      "URL: {{url}}\n" +
      "内容: {{content}}\n\n" +
      "簡潔で分かりやすい形式で整理してください。",

    build: function (data) {
      var prompt = this.template;
      prompt = prompt.replace('{{title}}', data.title || '');
      prompt = prompt.replace('{{url}}', data.url || '');
      prompt = prompt.replace('{{content}}', data.content || '');
      return prompt;
    }
  },

  // カレンダー追加用コンテンツ生成
  CALENDAR_CONTENT: {
    template: "以下の情報からカレンダーイベント用の詳細を抽出してください：\n\n" +
      "タイトル: {{title}}\n" +
      "URL: {{url}}\n" +
      "内容: {{content}}\n\n" +
      "イベントのタイトル、日時、場所、説明を含めてください。",

    build: function (data) {
      var prompt = this.template;
      prompt = prompt.replace('{{title}}', data.title || '');
      prompt = prompt.replace('{{url}}', data.url || '');
      prompt = prompt.replace('{{content}}', data.content || '');
      return prompt;
    }
  },

  // プロンプト取得のメイン関数
  getPrompt: function (type, data) {
    switch (type) {
      case 'vscode-analysis':
        return this.VSCODE_ANALYSIS.build(data);
      case 'email-analysis':
        return this.EMAIL_ANALYSIS.build(data);
      case 'email-compose':
        return this.EMAIL_COMPOSE.build(data);
      case 'page-analysis':
        return this.PAGE_ANALYSIS.build(data);
      case 'selection-analysis':
        return this.SELECTION_ANALYSIS.build(data);
      case 'teams-content':
        return this.TEAMS_CONTENT.build(data);
      case 'calendar-content':
        return this.CALENDAR_CONTENT.build(data);
      default:
        throw new Error('未知のプロンプトタイプ: ' + type);
    }
  },

  // 利用可能なプロンプトタイプ一覧
  getAvailableTypes: function () {
    return [
      'vscode-analysis',
      'email-analysis',
      'email-compose',
      'page-analysis',
      'selection-analysis',
      'teams-content',
      'calendar-content'
    ];
  }
};

// デフォルト設定
var DEFAULT_PROMPT_SETTINGS = {
  maxContentLength: 20000,
  enableLogging: true,
  retryAttempts: 3,
  timeoutMs: 30000
};

// Service Worker環境でのグローバル公開
if (typeof self !== 'undefined') {
  self.PromptManager = PromptManager;
  self.DEFAULT_PROMPT_SETTINGS = DEFAULT_PROMPT_SETTINGS;
}
