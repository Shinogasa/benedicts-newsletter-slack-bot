/**
 * Benedict's Newsletter 自動翻訳・要約 Slack 投稿ボット
 * メールマガジンの内容を自動で翻訳・要約し、Slackに投稿します
 */

// 設定関連
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const CONFIG = {
  GEMINI_API_KEY: SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY'),
  SLACK_API_TOKEN: SCRIPT_PROPERTIES.getProperty('SLACK_API_TOKEN'),
  SENDER_EMAIL: 'list@ben-evans.com',
  SLACK_CHANNEL: SCRIPT_PROPERTIES.getProperty('YOUR_SLACK_CHANNEL_ID')
};

// API関連の定数
const API = {
  GEMINI_ENDPOINT: `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`,
  SLACK_ENDPOINT: 'https://slack.com/api/chat.postMessage'
};

/**
 * メイン処理：メールを検索し、翻訳・要約してSlackに投稿する
 */
function processEmailToSlack() {
  // 検索条件: 指定送信元から、昨日以降の未読メール
  const query = `is:unread from:${CONFIG.SENDER_EMAIL} after:${getFormattedDateYesterday()}`;
  Logger.log(`メール検索クエリ: ${query}`);

  const threads = GmailApp.search(query);
  Logger.log(`検索結果: ${threads.length}件のスレッドが見つかりました`);

  if (threads.length === 0) {
    Logger.log('新着メールはありません');
    return;
  }

  // 各スレッドの処理
  threads.forEach(thread => {
    const messages = thread.getMessages();
    
    // 各メッセージの処理
    messages.forEach(message => {
      if (message.isUnread()) {
        processMessage(message);
        Utilities.sleep(1000); // APIレート制限回避
      }
    });
  });
}

/**
 * 個別のメッセージを処理する
 * @param {GmailMessage} message - 処理するメールメッセージ
 */
function processMessage(message) {
  Logger.log(`メッセージ処理開始: ${message.getSubject()}`);
  const originalBody = message.getPlainBody();

  if (!originalBody) {
    Logger.log('メール本文が取得できませんでした');
    return;
  }

  try {
    // 翻訳と要約を取得
    const processedContent = getTranslationAndSummary(originalBody);
    
    if (!processedContent) {
      Logger.log(`翻訳・要約の取得に失敗しました: ${message.getSubject()}`);
      return;
    }
    
    // Slackに投稿
    postContentToSlack(message, processedContent, originalBody);
    
    // 処理完了したメールを既読にする
    message.markRead();
    Logger.log(`処理完了: ${message.getSubject()}`);
    
  } catch (error) {
    Logger.log(`メッセージ処理中にエラーが発生しました: ${error} - 件名: ${message.getSubject()}`);
  }
}

/**
 * 翻訳と要約をSlackに投稿
 * @param {GmailMessage} message - 元のメールメッセージ
 * @param {Object} processedContent - 翻訳と要約を含むオブジェクト {translation, summary}
 * @param {string} originalBody - 元のメール本文
 */
function postContentToSlack(message, processedContent, originalBody) {
  const date = message.getDate().toLocaleDateString('ja-JP');
  const mainPost = postToSlack(
    `【新着IT動向メールマガジン要約 (${date})】\n${processedContent.summary}`,
    CONFIG.SLACK_CHANNEL
  );

  if (!mainPost || !mainPost.ok) {
    Logger.log(`Slackへの投稿に失敗しました: ${JSON.stringify(mainPost)}`);
    return;
  }

  // スレッドに翻訳と原文を投稿（スレッドIDが取得できた場合のみ）
  if (mainPost.ts) {
    const threadMessage = `【日本語全文訳】\n${processedContent.translation}\n\n---\n\n【Original English Text】\n${originalBody}`;
    postToSlack(threadMessage, CONFIG.SLACK_CHANNEL, mainPost.ts);
  } else {
    Logger.log('スレッドIDが取得できなかったため、スレッド投稿をスキップします');
  }
}

/**
 * Gemini APIを使用してテキストを翻訳・要約する
 * @param {string} textToProcess - 処理する英語テキスト
 * @return {Object|null} - 翻訳と要約を含むオブジェクト、エラー時はnull
 */
function getTranslationAndSummary(textToProcess) {
  const results = callGeminiApi(textToProcess);
  
  if (!results || !results.translation || !results.summary) {
    Logger.log('翻訳または要約の取得に失敗しました');
    return null;
  }
  
  return {
    translation: results.translation,
    summary: results.summary
  };
}

/**
 * Gemini APIを呼び出して翻訳と要約を取得
 * @param {string} textToProcess - 処理する英語テキスト
 * @return {Object|null} - 翻訳と要約を含むオブジェクト、エラー時はnull
 */
function callGeminiApi(textToProcess) {
  // Geminiに渡すプロンプト
  const prompt = `
以下の英語のIT動向に関するメールマガジンの内容について、指示に従って処理してください。

# 指示
1. まず、メールマガジン全文を自然な日本語に翻訳してください。
2. そしてメールマガジン内にURLがあった場合は取り除いてください。
3. 次に、翻訳した内容に基づいて、メールマガジンに含まれる複数の主要なトピックを特定し、それぞれのトピックについて最も重要なポイントを日本語で数行ずつ（全体で300文字程度を目安に）要約してください。
4. 要約は箇条書き形式で、各要点の見出しは Slack 向けのマークアップで *見出し:* という形式にしてください（アスタリスク1つで囲む）。例：「*AI開発の進展:* AIモデルが急速に進歩しています」
5. 各トピックの間には空白行を入れて見やすくしてください。
6. 最後に、以下のJSON形式（マークダウン記法なしの純粋なJSON）で、翻訳結果と要約結果のみを返してください。

\`\`\`json
{
  "translation": "ここに日本語の全文翻訳を入れてください",
  "summary": "ここに日本語の要約を入れてください（各見出しは *見出し:* の形式）"
}
\`\`\`

# メールマガジン本文
${textToProcess}
`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    Logger.log("Gemini API呼び出し開始...");
    const response = UrlFetchApp.fetch(API.GEMINI_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode !== 200) {
      Logger.log(`Gemini API呼び出しエラー (${responseCode}): ${responseBody}`);
      return null;
    }

    return parseGeminiResponse(responseBody);
  } catch (error) {
    Logger.log(`Gemini API呼び出し中にエラーが発生しました: ${error}`);
    return null;
  }
}

/**
 * Gemini APIのレスポンスをパース
 * @param {string} responseBody - APIからのレスポンス本文
 * @return {Object|null} - パース結果、エラー時はnull
 */
function parseGeminiResponse(responseBody) {
  try {
    const jsonResponse = JSON.parse(responseBody);
    const candidate = jsonResponse.candidates && jsonResponse.candidates[0];
    const content = candidate && candidate.content;
    const parts = content && content.parts;
    const textResult = parts && parts[0] && parts[0].text;

    if (!textResult) {
      Logger.log("Gemini APIのレスポンスからテキスト結果が見つかりませんでした");
      return null;
    }

    // JSON文字列を抽出
    const jsonTextMatch = textResult.match(/```json\s*([\s\S]*?)\s*```/);
    const jsonText = jsonTextMatch ? jsonTextMatch[1].trim() : textResult.trim();

    const result = JSON.parse(jsonText);
    if (!result.translation || !result.summary) {
      Logger.log("パースされたJSONに翻訳または要約が含まれていません");
      return null;
    }

    return result;
  } catch (e) {
    Logger.log(`Gemini APIレスポンスのパースに失敗しました: ${e}`);
    return null;
  }
}

/**
 * Slackにメッセージを投稿
 * @param {string} text - 投稿するテキスト
 * @param {string} channel - 投稿先チャンネル
 * @param {string} [threadTs=null] - スレッドの親メッセージタイムスタンプ
 * @return {Object} - Slack APIからのレスポンス
 */
function postToSlack(text, channel, threadTs = null) {
  try {
    Logger.log(`Slackへの投稿: ${channel}${threadTs ? ' (スレッド)' : ''}`);
    
    const payload = {
      channel: channel,
      text: text,
      link_names: true
    };
    
    if (threadTs) {
      payload.thread_ts = threadTs;
    }
    
    const options = {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: {
        'Authorization': `Bearer ${CONFIG.SLACK_API_TOKEN}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(API.SLACK_ENDPOINT, options);
    const responseJson = JSON.parse(response.getContentText());
    
    if (!responseJson.ok) {
      Logger.log(`Slackエラー: ${responseJson.error}`);
      if (responseJson.error === 'not_in_channel') {
        Logger.log(`ボットをチャンネル ${channel} に招待する必要があります`);
      }
    }
    
    return responseJson;
  } catch (error) {
    Logger.log(`Slack投稿中にエラーが発生しました: ${error}`);
    return { ok: false, error: error.toString() };
  }
}

/**
 * 昨日の日付を 'YYYY/MM/DD' 形式で取得
 * @return {string} フォーマットされた日付
 */
function getFormattedDateYesterday() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const year = yesterday.getFullYear();
  const month = ('0' + (yesterday.getMonth() + 1)).slice(-2);
  const day = ('0' + yesterday.getDate()).slice(-2);
  return `${year}/${month}/${day}`;
}

// --- デバッグ・テスト用関数 ---

/**
 * 特定のメールを手動でテスト処理する
 * 件名で検索して処理します
 */
function testSpecificEmail() {
  // 具体的な件名で検索
  const threads = GmailApp.search('subject:"Benedict\'s Newsletter: No. 585"');
  
  if (threads.length === 0) {
    Logger.log('指定した件名のスレッドが見つかりませんでした');
    return;
  }
  
  const messages = threads[0].getMessages();
  
  if (messages.length === 0) {
    Logger.log('スレッド内にメッセージが見つかりませんでした');
    return;
  }
  
  const message = messages[0];
  Logger.log(`テスト処理開始: ${message.getSubject()}`);
  
  const originalBody = message.getPlainBody();
  if (!originalBody) {
    Logger.log('メール本文が取得できませんでした');
    return;
  }

  try {
    const geminiResult = callGeminiApi(originalBody);
    
    if (!geminiResult) {
      Logger.log('Gemini APIからの結果取得に失敗しました');
      return;
    }
    
    Logger.log('翻訳と要約の取得に成功しました');
    
    // Slack投稿テスト
    const postResult = postToSlack(
      `【テスト要約】\n${geminiResult.summary}`,
      CONFIG.SLACK_CHANNEL
    );
    
    if (postResult && postResult.ok && postResult.ts) {
      const threadMessage = `【テスト日本語訳】\n${geminiResult.translation}\n\n---\n\n【テスト原文】\n${originalBody}`;
      postToSlack(threadMessage, CONFIG.SLACK_CHANNEL, postResult.ts);
      Logger.log('Slackへのテスト投稿に成功しました');
    } else {
      Logger.log(`Slackへのテスト投稿に問題がありました: ${JSON.stringify(postResult)}`);
    }
  } catch (error) {
    Logger.log(`テスト処理中にエラーが発生しました: ${error}`);
  }
}
