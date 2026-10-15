/**
 * Benedict's Newsletter 自動翻訳・要約 Slack 投稿ボット
 * GASから呼ぶエントリポイント。本物のGAS APIを組み立てて processor.gs に渡す
 * （このファイルだけは GAS 上でしか動かないので、Node のテスト対象外）
 */

// testSpecificEmail で処理するメールの件名
const TEST_EMAIL_SUBJECT = "Benedict's Newsletter: No. 585";

/**
 * 本物の外部依存を組み立てる
 */
function createDeps() {
  const config = loadConfig((key) => PropertiesService.getScriptProperties().getProperty(key));
  const fetch = (url, options) => UrlFetchApp.fetch(url, options);
  const sleep = (ms) => Utilities.sleep(ms);
  return {
    config,
    gemini: createGeminiClient(fetch, config.geminiApiKey, { sleep }),
    slack: createSlackClient(fetch, config.slackApiToken, config.slackChannelId),
    log: (message) => Logger.log(message),
    sleep,
  };
}

/**
 * メイン処理（トリガーから実行する）
 * 1通でも失敗したら最後に例外を投げ、GASのトリガー失敗通知を届ける
 */
function processEmailToSlack() {
  const deps = createDeps();
  const query = buildSearchQuery(deps.config.senderEmail, new Date());
  deps.log(`メール検索クエリ: ${query}`);

  const threads = GmailApp.search(query);
  deps.log(`検索結果: ${threads.length}件のスレッド`);

  const { processed, failures } = processThreads(threads, deps);
  deps.log(`処理完了: 成功 ${processed}件 / 失敗 ${failures.length}件`);

  if (failures.length > 0) {
    const detail = failures.map((f) => `${f.subject}: ${f.error.message}`).join('\n');
    throw new Error(`${failures.length}件のメールを処理できませんでした（未読のまま次回再試行します）\n${detail}`);
  }
}

/**
 * 件名を指定して1通だけ処理する（動作確認用）
 * 実際にSlackへ投稿するが、メールは既読にしない
 */
function testSpecificEmail() {
  const deps = createDeps();
  const threads = GmailApp.search(`subject:"${TEST_EMAIL_SUBJECT}"`);
  const message = threads.length > 0 ? threads[0].getMessages()[0] : null;
  if (!message) {
    throw new Error(`件名「${TEST_EMAIL_SUBJECT}」のメールが見つかりませんでした`);
  }

  processMessage(message, deps, { markAsRead: false });
}
