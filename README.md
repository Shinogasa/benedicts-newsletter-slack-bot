# benedicts-newsletter-slack-bot

## About

Benedict's Newsletterを翻訳+要約してSlackへ投稿するGoogle App Script

## Features

### Slack投稿フォーマット

このボットは以下の形式でSlackに投稿します:

#### 親投稿
- メールの件名をそのまま表示
- 例: `Benedict's Newsletter: No. 585`

#### スレッド投稿
1. **要約投稿**
   - ヘッダー: `【要約】`
   - 内容: メールマガジンの主要トピックを箇条書きで要約(約300文字)
   - Slackマークアップで見出しを強調表示

2. **全文翻訳投稿**
   - ヘッダー: `【日本語全文訳】`
   - 内容: メールマガジン全文の日本語翻訳
   - Slackマークアップで読みやすく整形

### 翻訳・要約機能

- **翻訳**: Gemini 2.5 Flashを使用して英語メールを自然な日本語に翻訳
- **要約**: 複数のトピックを特定し、各トピックの重要ポイントを抽出
- **マークアップ**: Slack形式で見やすく整形

## Usage

### 1. Google AI Studioに登録しAPIキーを入手

Geminiによる翻訳と要約を行うので[Google AI Studio](https://aistudio.google.com/apikey)でAPIキーを作成し控えておく  

### 2. Slack Appの作成

Slack Appを利用してチャンネルへ投稿するためあらかじめ[ここ](https://api.slack.com/apps)からAppを作る  

OAuth & Permissions > Scopes より Bot Token Scopes で chat:write をつけ、上部OAuth TokensよりOAuthトークンを作成する

### 3. GASの作成
 Google Driveより 新規 > その他 > Google App Scriptで新規App Scriptプロジェクトを作成

左バー プロジェクトの設定 > スクリプトプロパティを追加 より下記内容をいれて保存


|プロパティ|値|
|---|---|
|GEMINI_API_KEY|1で取得したAPIキー|
|SLACK_API_TOKEN|2で作成したOAuthトークン|
|SLACK_CHANNEL_ID|投稿したいスラックチャンネルのID|

保存したらエディタへ[スクリプト](./script.gs)を貼り付け

### 4. テスト

プルダウンからtestSpecificEmailを選択して実行するとテストが実行される
