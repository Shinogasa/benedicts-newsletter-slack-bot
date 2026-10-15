const { test } = require('node:test');
const assert = require('node:assert/strict');
const {
  buildGeminiRequest,
  parseGeminiResponse,
  createGeminiClient,
} = require('../src/gemini.gs');

// Gemini API のレスポンス本文を組み立てる
const geminiBody = (text) => JSON.stringify({ candidates: [{ content: { parts: [{ text }] } }] });

const validResult = JSON.stringify({ translation: '全文訳', summary: '*要点:* 要約' });

test('APIキーはURLに載せずヘッダで送る', () => {
  const { url, options } = buildGeminiRequest('secret-key', 'body');

  assert.ok(!url.includes('secret-key'));
  assert.equal(options.headers['x-goog-api-key'], 'secret-key');
  assert.equal(options.method, 'post');
  assert.equal(options.muteHttpExceptions, true);
});

test('プロンプトにメール本文を含める', () => {
  const { options } = buildGeminiRequest('key', 'NEWSLETTER BODY');
  const payload = JSON.parse(options.payload);

  assert.match(payload.contents[0].parts[0].text, /NEWSLETTER BODY/);
});

test('JSONの翻訳と要約を取り出す', () => {
  assert.deepEqual(parseGeminiResponse(200, geminiBody(validResult)), {
    translation: '全文訳',
    summary: '*要点:* 要約',
  });
});

test('コードフェンスで囲まれたJSONも読める', () => {
  const fenced = '```json\n' + validResult + '\n```';

  assert.equal(parseGeminiResponse(200, geminiBody(fenced)).translation, '全文訳');
});

test('200以外のステータスはコード付きで例外にする', () => {
  assert.throws(() => parseGeminiResponse(429, '{"error":"rate"}'), /429/);
});

test('候補のテキストが無ければ例外にする', () => {
  assert.throws(() => parseGeminiResponse(200, JSON.stringify({ candidates: [] })), /テキスト/);
});

test('モデルの出力がJSONでなければ例外にする', () => {
  assert.throws(() => parseGeminiResponse(200, geminiBody('ごめんなさい')), /JSON/);
});

test('翻訳か要約が欠けていれば例外にする', () => {
  const partial = JSON.stringify({ translation: '全文訳' });

  assert.throws(() => parseGeminiResponse(200, geminiBody(partial)), /summary/);
});

// 指定したステータスを順に返す fetch。使い切ったら最後のステータスを返し続ける
const sequenceFetch = (codes) => {
  const calls = [];
  const fetch = () => {
    const code = codes[Math.min(calls.length, codes.length - 1)];
    calls.push(code);
    return {
      getResponseCode: () => code,
      getContentText: () => (code === 200 ? geminiBody(validResult) : '{"error":{}}'),
    };
  };
  return { fetch, calls };
};

const retryOptions = (sleeps) => ({ sleep: (ms) => sleeps.push(ms), random: () => 0 });

test('503や429はバックオフしながら再試行し、成功すれば結果を返す', () => {
  const { fetch, calls } = sequenceFetch([503, 429, 200]);
  const sleeps = [];

  const result = createGeminiClient(fetch, 'key', retryOptions(sleeps)).translateAndSummarize('body');

  assert.equal(result.translation, '全文訳');
  assert.deepEqual(calls, [503, 429, 200]);
  assert.equal(sleeps.length, 2);
  assert.ok(sleeps[1] > sleeps[0], `待ち時間が伸びていない: ${sleeps}`);
});

test('再試行は最大3回で、それでも503なら例外にする', () => {
  const { fetch, calls } = sequenceFetch([503]);
  const sleeps = [];

  assert.throws(
    () => createGeminiClient(fetch, 'key', retryOptions(sleeps)).translateAndSummarize('body'),
    /503/,
  );
  assert.equal(calls.length, 4);
  assert.equal(sleeps.length, 3);
});

test('400などの再試行しても直らないエラーは、すぐ例外にする', () => {
  const { fetch, calls } = sequenceFetch([400]);
  const sleeps = [];

  assert.throws(
    () => createGeminiClient(fetch, 'key', retryOptions(sleeps)).translateAndSummarize('body'),
    /400/,
  );
  assert.equal(calls.length, 1);
  assert.equal(sleeps.length, 0);
});

test('クライアントは fetch の結果を解析して返す', () => {
  const calls = [];
  const fetch = (url, options) => {
    calls.push({ url, options });
    return { getResponseCode: () => 200, getContentText: () => geminiBody(validResult) };
  };

  const result = createGeminiClient(fetch, 'key').translateAndSummarize('body');

  assert.equal(calls.length, 1);
  assert.equal(result.summary, '*要点:* 要約');
});
