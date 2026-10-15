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

保存したら、エディタで `src/` 配下の各ファイルと同じ名前のスクリプトファイルを作り、中身を貼り付ける

|ファイル|役割|
|---|---|
|[main.gs](./src/main.gs)|GASから実行するエントリポイント（`processEmailToSlack` / `testSpecificEmail`）|
|[processor.gs](./src/processor.gs)|メール1通の処理と、既読にする条件|
|[gemini.gs](./src/gemini.gs)|Gemini APIでの翻訳・要約|
|[slack.gs](./src/slack.gs)|Slackへの投稿と削除|
|[config.gs](./src/config.gs)|スクリプトプロパティの読み込みと検証|

GASは全ファイルを1つのグローバルスコープで読み込むので、ファイルの順番は問わない。
ただし、別のファイルと同じ名前の関数や定数を作るとエラーになる

### 4. テスト

プルダウンから `testSpecificEmail` を選択して実行すると、`src/main.gs` の `TEST_EMAIL_SUBJECT` の件名のメールを1通処理する。
Slackへは実際に投稿されるが、メールは既読にしない

### 5. トリガーの設定

`processEmailToSlack` を時間主導型トリガーで定期実行する。

- 3件の投稿（親・要約・全文訳）がすべて成功したメールだけを既読にする
- 要約か全文訳の投稿に失敗したら、親投稿を削除し、メールは未読のまま残す。次回の実行で最初からやり直す
- 1通でも失敗すると実行の最後に例外を投げるので、GASのトリガー失敗通知メールが届く

## Development

GASのAPIに触れる部分は `src/main.gs` だけにまとめ、それ以外は Node.js の標準テストランナーでテストできる。
依存パッケージは無い。

```sh
npm test
```

`test/main.test.js` は `src/*.gs` をGASと同じく1つのグローバルスコープに読み込み、偽のGAS APIで配線を確かめる
