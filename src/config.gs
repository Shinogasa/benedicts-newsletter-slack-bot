/**
 * 設定の読み込み
 * スクリプトプロパティから必要な値を取り出し、欠けていれば実行を止める
 */

const SENDER_EMAIL = 'list@ben-evans.com';

const REQUIRED_PROPERTIES = ['GEMINI_API_KEY', 'SLACK_API_TOKEN', 'SLACK_CHANNEL_ID'];

/**
 * 設定を読み込む
 * @param {function(string): (string|null)} getProperty - プロパティ名から値を返す関数
 * @return {{geminiApiKey: string, slackApiToken: string, slackChannelId: string, senderEmail: string}}
 * @throws {Error} 必須プロパティが未設定のとき
 */
function loadConfig(getProperty) {
  const missing = REQUIRED_PROPERTIES.filter((key) => !getProperty(key));
  if (missing.length > 0) {
    throw new Error(`スクリプトプロパティが未設定です: ${missing.join(', ')}`);
  }

  return {
    geminiApiKey: getProperty('GEMINI_API_KEY'),
    slackApiToken: getProperty('SLACK_API_TOKEN'),
    slackChannelId: getProperty('SLACK_CHANNEL_ID'),
    senderEmail: SENDER_EMAIL,
  };
}

// GASには module が無い。Node のテストから読み込むときだけ公開する
if (typeof module !== 'undefined') {
  module.exports = { loadConfig };
}
