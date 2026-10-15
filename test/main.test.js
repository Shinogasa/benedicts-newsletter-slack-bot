// GASと同じく src/*.gs を1つのグローバルスコープに読み込み、main.gs の配線を確かめる
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC_DIR = path.join(__dirname, '..', 'src');

const geminiResponse = JSON.stringify({
  candidates: [{ content: { parts: [{ text: JSON.stringify({ translation: '全文訳', summary: '要約' }) }] } }],
});

// GASのグローバルAPIを偽物にした実行環境を作る
const createGasContext = ({ slackResults }) => {
  const message = {
    markedRead: false,
    getSubject: () => "Benedict's Newsletter: No. 600",
    getPlainBody: () => 'English body',
    isUnread: () => true,
    markRead: () => { message.markedRead = true; },
  };
  const fetched = [];
  const context = {
    message,
    fetched,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => ({ GEMINI_API_KEY: 'g', SLACK_API_TOKEN: 's', SLACK_CHANNEL_ID: 'C1' })[key],
      }),
    },
    GmailApp: { search: () => [{ getMessages: () => [message] }] },
    UrlFetchApp: {
      fetch: (url, options) => {
        fetched.push(url);
        if (url.includes('generativelanguage')) {
          return { getResponseCode: () => 200, getContentText: () => geminiResponse };
        }
        const result = slackResults.shift() || { ok: true, ts: '1.0' };
        return { getContentText: () => JSON.stringify(result) };
      },
    },
    Logger: { log: () => {} },
    Utilities: { sleep: () => {} },
  };
  vm.createContext(context);
  fs.readdirSync(SRC_DIR)
    .filter((file) => file.endsWith('.gs'))
    .forEach((file) => vm.runInContext(fs.readFileSync(path.join(SRC_DIR, file), 'utf8'), context));
  return context;
};

test('processEmailToSlack は投稿して既読にする', () => {
  const context = createGasContext({ slackResults: [] });

  vm.runInContext('processEmailToSlack()', context);

  assert.equal(context.fetched.filter((url) => url.endsWith('chat.postMessage')).length, 3);
  assert.equal(context.message.markedRead, true);
});

test('processEmailToSlack は失敗があれば最後に例外を投げる', () => {
  const context = createGasContext({ slackResults: [{ ok: false, error: 'invalid_auth' }] });

  assert.throws(() => vm.runInContext('processEmailToSlack()', context), /invalid_auth/);
  assert.equal(context.message.markedRead, false);
});
