/**
 * Slack への投稿と削除
 * HTTP送信は引数の fetch に任せ、このファイルはリクエストの組み立てと結果の判定だけを行う
 */

const SLACK_API_BASE = 'https://slack.com/api/';

/**
 * Slackクライアントを作る
 * @param {function(string, Object): {getContentText: function(): string}} fetch
 * @param {string} token - Bot User OAuth Token
 * @param {string} channel - 投稿先チャンネルID
 */
function createSlackClient(fetch, token, channel) {
  const call = (method, payload) => {
    const response = fetch(SLACK_API_BASE + method, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: `Bearer ${token}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const result = JSON.parse(response.getContentText());
    if (!result.ok) {
      const hint = result.error === 'not_in_channel'
        ? `（ボットをチャンネル ${channel} に招待してください）`
        : '';
      throw new Error(`Slack ${method} が失敗しました: ${result.error}${hint}`);
    }
    return result;
  };

  return {
    /**
     * メッセージを投稿する
     * @param {string} text
     * @param {string} [threadTs] - 指定するとそのスレッドへ返信する
     * @return {{ts: string}}
     */
    post(text, threadTs) {
      const payload = { channel, text, link_names: true };
      if (threadTs) {
        payload.thread_ts = threadTs;
      }
      return { ts: call('chat.postMessage', payload).ts };
    },

    /**
     * ボット自身の投稿を削除する
     * @param {string} ts
     */
    remove(ts) {
      call('chat.delete', { channel, ts });
    },
  };
}

// GASには module が無い。Node のテストから読み込むときだけ公開する
if (typeof module !== 'undefined') {
  module.exports = { createSlackClient };
}
