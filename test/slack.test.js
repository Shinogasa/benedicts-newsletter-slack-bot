const { test } = require('node:test');
const assert = require('node:assert/strict');
const { createSlackClient } = require('../src/slack.gs');

// 呼び出しを記録し、決まったSlackレスポンスを返す fetch
const recordingFetch = (responseJson) => {
  const calls = [];
  const fetch = (url, options) => {
    calls.push({ url, options, payload: JSON.parse(options.payload) });
    return { getContentText: () => JSON.stringify(responseJson) };
  };
  return { fetch, calls };
};

test('親投稿は thread_ts を付けずに送り、ts を返す', () => {
  const { fetch, calls } = recordingFetch({ ok: true, ts: '111.222' });

  const result = createSlackClient(fetch, 'xoxb-token', 'C123').post('件名');

  assert.deepEqual(result, { ts: '111.222' });
  assert.equal(calls[0].url, 'https://slack.com/api/chat.postMessage');
  assert.equal(calls[0].options.headers.Authorization, 'Bearer xoxb-token');
  assert.deepEqual(calls[0].payload, { channel: 'C123', text: '件名', link_names: true });
});

test('スレッド投稿は thread_ts を付ける', () => {
  const { fetch, calls } = recordingFetch({ ok: true, ts: '333.444' });

  createSlackClient(fetch, 'token', 'C123').post('要約', '111.222');

  assert.equal(calls[0].payload.thread_ts, '111.222');
});

test('ok:false はエラーコード付きで例外にする', () => {
  const { fetch } = recordingFetch({ ok: false, error: 'invalid_auth' });

  assert.throws(() => createSlackClient(fetch, 'token', 'C123').post('件名'), /invalid_auth/);
});

test('not_in_channel にはボットの招待を案内する', () => {
  const { fetch } = recordingFetch({ ok: false, error: 'not_in_channel' });

  assert.throws(() => createSlackClient(fetch, 'token', 'C123').post('件名'), /招待/);
});

test('remove は chat.delete に channel と ts を送る', () => {
  const { fetch, calls } = recordingFetch({ ok: true });

  createSlackClient(fetch, 'token', 'C123').remove('111.222');

  assert.equal(calls[0].url, 'https://slack.com/api/chat.delete');
  assert.deepEqual(calls[0].payload, { channel: 'C123', ts: '111.222' });
});
