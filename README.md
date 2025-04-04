# benedicts-newsletter-slack-bot

## About

Benedict's Newsletterを翻訳+要約してSlackへ投稿するGoogle App Script

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
|YOUR_SLACK_CHANNEL_ID|投稿したいスラックチャンネルのID|

保存したらエディタへ[スクリプト](./script.gs)を貼り付け

### 4. テスト

プルダウンからtestSpecificEmailを選択して実行するとテストが実行される
