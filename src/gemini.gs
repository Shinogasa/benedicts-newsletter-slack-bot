/**
 * Gemini API による翻訳・要約
 * HTTP送信は引数の fetch に任せ、このファイルはリクエストの組み立てと解析だけを行う
 */

const GEMINI_ENDPOINT =
  'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent';

// エラーメッセージに載せるレスポンス本文の上限
const GEMINI_ERROR_BODY_LIMIT = 500;

/**
 * Geminiに渡すプロンプトを組み立てる
 * @param {string} newsletterText - メールマガジン本文（英語）
 * @return {string}
 */
function buildGeminiPrompt(newsletterText) {
  return `
以下の英語のIT動向に関するメールマガジンの内容について、指示に従って処理してください。

# 指示
1. まず、メールマガジン全文を自然な日本語に翻訳してください。
2. そしてメールマガジン内にURLがあった場合は取り除いてください。
3. 次に、翻訳した内容に基づいて、メールマガジンに含まれる複数の主要なトピックを特定し、それぞれのトピックについて最も重要なポイントを日本語で数行ずつ（全体で300文字程度を目安に）要約してください。
4. 要約は箇条書き形式で、各要点の見出しは Slack 向けのマークアップで *見出し:* という形式にしてください（アスタリスク1つで囲む）。例：「*AI開発の進展:* AIモデルが急速に進歩しています」
5. 翻訳結果に関しても同様にSlack向けのマークアップで出力してSlack上で各話題が見やすいようにしてください。
6. 各トピックの間には空白行を入れて見やすくしてください。
7. 最後に、以下のJSON形式（マークダウン記法なしの純粋なJSON）で、翻訳結果と要約結果のみを返してください。

\`\`\`json
{
  "translation": "ここに日本語の全文翻訳を入れてください",
  "summary": "ここに日本語の要約を入れてください（各見出しは *見出し:* の形式）"
}
\`\`\`

# メールマガジン本文
${newsletterText}
`;
}

/**
 * UrlFetchApp.fetch に渡すURLとオプションを組み立てる
 * APIキーはURLに載せない（例外メッセージやログにURLが出ても漏れないように）
 * @param {string} apiKey
 * @param {string} newsletterText
 * @return {{url: string, options: Object}}
 */
function buildGeminiRequest(apiKey, newsletterText) {
  const payload = {
    contents: [{ parts: [{ text: buildGeminiPrompt(newsletterText) }] }],
  };

  return {
    url: GEMINI_ENDPOINT,
    options: {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    },
  };
}

/**
 * Gemini APIのレスポンスから翻訳と要約を取り出す
 * @param {number} responseCode
 * @param {string} responseBody
 * @return {{translation: string, summary: string}}
 * @throws {Error} ステータス異常・形式不正・項目欠落のとき
 */
function parseGeminiResponse(responseCode, responseBody) {
  if (responseCode !== 200) {
    throw new Error(
      `Gemini APIがエラーを返しました (${responseCode}): ${responseBody.slice(0, GEMINI_ERROR_BODY_LIMIT)}`,
    );
  }

  const candidate = (JSON.parse(responseBody).candidates || [])[0];
  const parts = candidate && candidate.content && candidate.content.parts;
  const textResult = parts && parts[0] && parts[0].text;
  if (!textResult) {
    throw new Error('Gemini APIのレスポンスにテキストがありません');
  }

  // 指示に反してコードフェンスで囲んで返すことがあるので外す
  const fenced = textResult.match(/```json\s*([\s\S]*?)\s*```/);
  const jsonText = fenced ? fenced[1].trim() : textResult.trim();

  let result;
  try {
    result = JSON.parse(jsonText);
  } catch (e) {
    throw new Error(`Geminiの出力をJSONとして解析できません: ${e.message}`);
  }

  const missing = ['translation', 'summary'].filter((key) => !result[key]);
  if (missing.length > 0) {
    throw new Error(`Geminiの出力に項目がありません: ${missing.join(', ')}`);
  }

  return { translation: result.translation, summary: result.summary };
}

/**
 * Geminiクライアントを作る
 * @param {function(string, Object): {getResponseCode: function(): number, getContentText: function(): string}} fetch
 * @param {string} apiKey
 */
function createGeminiClient(fetch, apiKey) {
  return {
    translateAndSummarize(newsletterText) {
      const { url, options } = buildGeminiRequest(apiKey, newsletterText);
      const response = fetch(url, options);
      return parseGeminiResponse(response.getResponseCode(), response.getContentText());
    },
  };
}

// GASには module が無い。Node のテストから読み込むときだけ公開する
if (typeof module !== 'undefined') {
  module.exports = { buildGeminiRequest, parseGeminiResponse, createGeminiClient };
}
