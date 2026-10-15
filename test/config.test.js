const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadConfig } = require('../src/config.gs');

const fakeProperties = (values) => (key) => (key in values ? values[key] : null);

test('必須プロパティがすべてあれば設定を返す', () => {
  const config = loadConfig(fakeProperties({
    GEMINI_API_KEY: 'gemini-key',
    SLACK_API_TOKEN: 'xoxb-token',
    SLACK_CHANNEL_ID: 'C123',
  }));

  assert.deepEqual(config, {
    geminiApiKey: 'gemini-key',
    slackApiToken: 'xoxb-token',
    slackChannelId: 'C123',
    senderEmail: 'list@ben-evans.com',
  });
});

test('欠けているプロパティ名をすべて含めて例外を投げる', () => {
  assert.throws(
    () => loadConfig(fakeProperties({ SLACK_API_TOKEN: 'xoxb-token', SLACK_CHANNEL_ID: '' })),
    /GEMINI_API_KEY, SLACK_CHANNEL_ID/,
  );
});
