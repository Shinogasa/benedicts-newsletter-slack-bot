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
