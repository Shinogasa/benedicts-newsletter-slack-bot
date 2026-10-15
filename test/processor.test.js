const { test } = require('node:test');
const assert = require('node:assert/strict');
const {
  buildSearchQuery,
  processMessage,
  processThreads,
} = require('../src/processor.gs');

const fakeMessage = ({ subject = 'Newsletter No. 1', body = 'English body', unread = true } = {}) => {
  const message = {
    markedRead: false,
    getSubject: () => subject,
    getPlainBody: () => body,
    isUnread: () => unread,
    markRead: () => { message.markedRead = true; },
  };
  return message;
};

const fakeThread = (...messages) => ({ getMessages: () => messages });

// failOn に指定した番目（1始まり）の投稿だけ失敗する Slack
const fakeSlack = ({ failOn = [], failRemove = false } = {}) => {
  const slack = {
    posts: [],
    removed: [],
    post(text, threadTs) {
      const count = slack.posts.length + 1;
      if (failOn.includes(count)) {
        throw new Error(`post ${count} failed`);
      }
      slack.posts.push({ text, threadTs });
      return { ts: `ts-${count}` };
    },
    remove(ts) {
      if (failRemove) {
        throw new Error('remove failed');
      }
      slack.removed.push(ts);
    },
  };
  return slack;
};

const fakeGemini = ({ fail = false } = {}) => ({
  translateAndSummarize: () => {
    if (fail) {
      throw new Error('gemini failed');
    }
    return { translation: '全文訳', summary: '要約' };
  },
});

const deps = (overrides = {}) => ({
  gemini: fakeGemini(),
  slack: fakeSlack(),
  log: () => {},
  sleep: () => {},
  ...overrides,
});

test('検索クエリは送信元と、ゼロ埋めした今日の日付を含む', () => {
  const query = buildSearchQuery('list@example.com', new Date(2026, 0, 5));

  assert.equal(query, 'is:unread from:list@example.com before:2026/01/05');
});

test('3件とも投稿できたら既読にする', () => {
  const message = fakeMessage();
  const d = deps();

  processMessage(message, d);

  assert.deepEqual(d.slack.posts, [
    { text: 'Newsletter No. 1', threadTs: undefined },
    { text: '【要約】\n要約', threadTs: 'ts-1' },
    { text: '【日本語全文訳】\n全文訳', threadTs: 'ts-1' },
  ]);
  assert.equal(message.markedRead, true);
});

test('markAsRead: false なら投稿しても既読にしない', () => {
  const message = fakeMessage();

  processMessage(message, deps(), { markAsRead: false });

  assert.equal(message.markedRead, false);
});

test('本文が空なら投稿せず例外にする', () => {
  const message = fakeMessage({ body: '' });
  const d = deps();

  assert.throws(() => processMessage(message, d), /本文/);
  assert.equal(d.slack.posts.length, 0);
  assert.equal(message.markedRead, false);
});

test('Geminiが失敗したら投稿せず既読にもしない', () => {
  const message = fakeMessage();
  const d = deps({ gemini: fakeGemini({ fail: true }) });

  assert.throws(() => processMessage(message, d), /gemini failed/);
  assert.equal(d.slack.posts.length, 0);
  assert.equal(message.markedRead, false);
});

test('親投稿が失敗したら既読にしない', () => {
  const message = fakeMessage();
  const d = deps({ slack: fakeSlack({ failOn: [1] }) });

  assert.throws(() => processMessage(message, d), /post 1 failed/);
  assert.equal(message.markedRead, false);
  assert.deepEqual(d.slack.removed, []);
});

test('スレッド投稿が失敗したら親投稿を削除し、既読にしない', () => {
  const message = fakeMessage();
  const d = deps({ slack: fakeSlack({ failOn: [3] }) });

  assert.throws(() => processMessage(message, d), /post 3 failed/);
  assert.deepEqual(d.slack.removed, ['ts-1']);
  assert.equal(message.markedRead, false);
});

test('親投稿の削除も失敗したら、元の投稿エラーを投げる', () => {
  const message = fakeMessage();
  const logs = [];
  const d = deps({ slack: fakeSlack({ failOn: [2], failRemove: true }), log: (m) => logs.push(m) });

  assert.throws(() => processMessage(message, d), /post 2 failed/);
  assert.ok(logs.some((m) => m.includes('remove failed')));
  assert.equal(message.markedRead, false);
});

test('未読のメッセージだけを処理し、1件の失敗で止めずに失敗を集める', () => {
  const ok = fakeMessage({ subject: 'ok' });
  const broken = fakeMessage({ subject: 'broken', body: '' });
  const alreadyRead = fakeMessage({ subject: 'read', unread: false });
  const d = deps();

  const result = processThreads([fakeThread(broken, alreadyRead), fakeThread(ok)], d);

  assert.equal(result.processed, 1);
  assert.equal(result.failures.length, 1);
  assert.equal(result.failures[0].subject, 'broken');
  assert.equal(ok.markedRead, true);
  assert.equal(d.slack.posts.some((p) => p.text === 'read'), false);
});
