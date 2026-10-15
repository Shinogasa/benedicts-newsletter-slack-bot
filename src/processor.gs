/**
 * メール1通を翻訳・要約してSlackへ投稿する処理の本体
 * 外部とのやり取り（Gemini・Slack・ログ・sleep）はすべて deps で受け取る
 *
 * 完了の扱い: 3件の投稿がすべて成功したときだけ既読にする。
 * 途中で失敗したメールは未読のまま残り、次回の実行で最初からやり直す（重複より欠落を避ける）。
 */

// APIのレート制限を避けるため、メッセージごとに空ける間隔
const PROCESS_INTERVAL_MS = 1000;

/**
 * Gmail検索用の日付 'YYYY/MM/DD' を返す
 * @param {Date} date
 * @return {string}
 */
function formatGmailDate(date) {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${date.getFullYear()}/${month}/${day}`;
}

/**
 * 指定した送信元からの、今日より前に届いた未読メールを探すクエリを返す
 * @param {string} senderEmail
 * @param {Date} today
 * @return {string}
 */
function buildSearchQuery(senderEmail, today) {
  return `is:unread from:${senderEmail} before:${formatGmailDate(today)}`;
}

/**
 * スレッド投稿に失敗したとき、親投稿を消して重複を防ぐ
 * 削除の失敗はログに残すだけにし、呼び出し元では元の投稿エラーを優先する
 */
function removeParentPost(slack, parentTs, log) {
  try {
    slack.remove(parentTs);
  } catch (removeError) {
    log(`親投稿の削除に失敗しました。Slackに親投稿だけが残っています (ts=${parentTs}): ${removeError.message}`);
  }
}

/**
 * メール1通を処理する
 * @param {GmailMessage} message
 * @param {{gemini: Object, slack: Object, log: function(string)}} deps
 * @param {{markAsRead?: boolean}} [options]
 * @throws {Error} どこかの段階で失敗したとき（そのときメールは既読にならない）
 */
function processMessage(message, deps, options = {}) {
  const { gemini, slack, log } = deps;
  const { markAsRead = true } = options;
  const subject = message.getSubject();
  log(`メッセージ処理開始: ${subject}`);

  const body = message.getPlainBody();
  if (!body) {
    throw new Error('メール本文が取得できませんでした');
  }

  const { translation, summary } = gemini.translateAndSummarize(body);

  const parent = slack.post(subject);
  try {
    slack.post(`【要約】\n${summary}`, parent.ts);
    slack.post(`【日本語全文訳】\n${translation}`, parent.ts);
  } catch (error) {
    removeParentPost(slack, parent.ts, log);
    throw error;
  }

  if (markAsRead) {
    message.markRead();
  }
  log(`処理完了: ${subject}`);
}

/**
 * スレッド内の未読メッセージを順に処理する。1通の失敗で止めずに続行する
 * @param {GmailThread[]} threads
 * @param {{gemini: Object, slack: Object, log: function(string), sleep: function(number)}} deps
 * @return {{processed: number, failures: {subject: string, error: Error}[]}}
 */
function processThreads(threads, deps) {
  const unreadMessages = threads
    .flatMap((thread) => thread.getMessages())
    .filter((message) => message.isUnread());

  const failures = [];
  unreadMessages.forEach((message) => {
    try {
      processMessage(message, deps);
    } catch (error) {
      deps.log(`メッセージ処理に失敗しました: ${message.getSubject()} - ${error.message}`);
      failures.push({ subject: message.getSubject(), error });
    }
    deps.sleep(PROCESS_INTERVAL_MS);
  });

  return { processed: unreadMessages.length - failures.length, failures };
}

// GASには module が無い。Node のテストから読み込むときだけ公開する
if (typeof module !== 'undefined') {
  module.exports = { buildSearchQuery, processMessage, processThreads };
}
