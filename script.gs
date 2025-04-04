// スクリプトプロパティから設定値を取得
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const GEMINI_API_KEY = SCRIPT_PROPERTIES.getProperty('GEMINI_API_KEY');
const SENDER_EMAIL = 'list@ben-evans.com'; // ★★★ 送信元メールアドレスに書き換えてください ★★★
const SLACK_CHANNEL = SCRIPT_PROPERTIES.getProperty('YOUR_SLACK_CHANNEL_ID');

// Gemini APIのエンドポイント (text-onlyモデルの場合)
const GEMINI_API_ENDPOINT = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

/**
 * メイン処理：メールを検索し、翻訳・要約してSlackに投稿する
 */
function processEmailToSlack() {
  // 検索条件: 指定送信元から、昨日以降の未読メール
  const query = `is:unread from:${SENDER_EMAIL} after:${getFormattedDateYesterday()}`;
  Logger.log(`Searching Gmail with query: ${query}`);

  const threads = GmailApp.search(query);
  Logger.log(`Found ${threads.length} threads.`);

  if (threads.length === 0) {
    Logger.log('No new emails found.');
    return;
  }

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      // 未読メッセージのみ処理 (念のため)
      if (message.isUnread()) {
        Logger.log(`Processing message: ${message.getSubject()}`);
        const originalBody = message.getPlainBody(); // HTMLからプレーンテキストを取得

        if (!originalBody) {
            Logger.log('Could not get email body.');
            return; // 本文がなければスキップ
        }

        try {
          // Gemini APIで翻訳と要約を取得
          const geminiResult = callGeminiApi(originalBody);

          if (geminiResult && geminiResult.translation && geminiResult.summary) {
            // Slackに投稿
            const postResult = postToSlack(
              `【新着IT動向メールマガジン要約 (${message.getDate().toLocaleDateString('ja-JP')})】\n${geminiResult.summary}`,
              SLACK_CHANNEL
            );

            if (postResult && postResult.ok && postResult.ts) {
              // スレッドに翻訳と原文を投稿
              const threadMessage = `【日本語全文訳】\n${geminiResult.translation}\n\n---\n\n【Original English Text】\n${originalBody}`;
              postToSlack(threadMessage, SLACK_CHANNEL, postResult.ts); // tsを指定してスレッド投稿

              // 処理済みメールを既読にする
              message.markRead();
              Logger.log(`Message processed and marked as read: ${message.getSubject()}`);
            } else {
              Logger.log(`Failed to post summary to Slack or get timestamp. Message subject: ${message.getSubject()}`);
            }
          } else {
             Logger.log(`Failed to get translation/summary from Gemini API. Message subject: ${message.getSubject()}`);
          }
        } catch (error) {
          Logger.log(`Error processing message: ${error} - Subject: ${message.getSubject()}`);
          // 必要に応じてエラー通知などを追加
        }
         Utilities.sleep(1000); // APIレート制限回避のため少し待機
      }
    });
  });
}

/**
 * Gemini APIを呼び出して翻訳と要約を取得する関数
 * @param {string} textToProcess 処理対象の英語テキスト
 * @return {object|null} { translation: string, summary: string } 形式のオブジェクト、またはエラー時 null
 */
function callGeminiApi(textToProcess) {
  // Geminiに渡すプロンプト
  const prompt = `
以下の英語のIT動向に関するメールマガジンの内容について、指示に従って処理してください。

# 指示
1. まず、メールマガジン全文を自然な日本語に翻訳してください。
2. そしてメールマガジン内にURLがあった場合は取り除いてください。
3. 次に、翻訳した内容に基づいて、メールマガジンに含まれる複数の主要なトピックを特定し、それぞれのトピックについて最も重要なポイントを日本語で数行ずつ（全体で300文字程度を目安に）要約してください。
4. 要約は箇条書き形式で、各要点の見出しは Slack 向けのマークアップで *見出し:* という形式にしてください（アスタリスク1つで囲む）。例：「* AI開発の進展:* AIモデルが急速に進歩しています」
5. 見出しの記載をしたら改行をしてから要約文を続け(例: * AI開発の進展:*\n)、要約文の次は空白行を入れて見やすくしてください。
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
    muteHttpExceptions: true // エラー時にもレスポンスを取得するため
  };

  try {
    Logger.log("Calling Gemini API...");
    const response = UrlFetchApp.fetch(GEMINI_API_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      Logger.log("Gemini API call successful.");
      const jsonResponse = JSON.parse(responseBody);
      // Geminiの応答構造に合わせて調整が必要な場合があります
      const candidate = jsonResponse.candidates && jsonResponse.candidates[0];
      const content = candidate && candidate.content;
      const parts = content && content.parts;
      const textResult = parts && parts[0] && parts[0].text;

      if (textResult) {
        Logger.log("Gemini API response text found.");
        // JSON文字列を抽出してパース (```json ``` を除去する必要がある場合)
        const jsonTextMatch = textResult.match(/```json\s*([\s\S]*?)\s*```/);
        const jsonText = jsonTextMatch ? jsonTextMatch[1].trim() : textResult.trim();

        try {
            const result = JSON.parse(jsonText);
             if (result.translation && result.summary) {
                Logger.log("Successfully parsed translation and summary.");
                return result;
            } else {
                 Logger.log("Parsed JSON does not contain translation or summary.");
                 Logger.log(`Parsed JSON: ${JSON.stringify(result)}`);
                 return null;
            }
        } catch (e) {
            Logger.log(`Failed to parse JSON from Gemini response: ${e}`);
            Logger.log(`Raw text result from Gemini: ${textResult}`);
            return null;
        }
      } else {
        Logger.log("Could not find text result in Gemini API response.");
        Logger.log(`Full Gemini Response: ${responseBody}`);
        return null;
      }
    } else {
      Logger.log(`Gemini API call failed with status ${responseCode}: ${responseBody}`);
      return null;
    }
  } catch (error) {
    Logger.log(`Error calling Gemini API: ${error}`);
    return null;
  }
}


/**
 * Slackにメッセージを投稿する関数
 * @param {string} text 投稿するテキスト
 * @param {string} channel 投稿先のチャンネル名 or ID
 * @param {string} [threadTs=null] スレッド投稿する場合の親メッセージのタイムスタンプ
 * @return {object|null} Slack APIからのレスポンス(JSONパース後)、またはエラー時 null
 */
function postToSlack(text, channel, threadTs = null) {
  try {
    const SLACK_TOKEN = SCRIPT_PROPERTIES.getProperty('SLACK_API_TOKEN');
    const API_URL = 'https://slack.com/api/chat.postMessage';
    
    // ログを追加して確認
    Logger.log(`Posting to channel: ${channel}`);
    
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
        'Authorization': `Bearer ${SLACK_TOKEN}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(API_URL, options);
    const responseJson = JSON.parse(response.getContentText());
    
    Logger.log(`Full Slack response: ${JSON.stringify(responseJson)}`);
    
    if (!responseJson.ok) {
      Logger.log(`Slack error: ${responseJson.error}`);
      if (responseJson.error === 'not_in_channel') {
        Logger.log(`Bot needs to be invited to the channel ${channel} first!`);
      }
    }
    
    return responseJson;
  } catch (error) {
    Logger.log(`Error in postToSlack: ${error}`);
    return { ok: false, error: error.toString() };
  }
}

/**
 * 昨日の日付を 'YYYY/MM/DD' 形式で取得するヘルパー関数
 * @return {string}
 */
function getFormattedDateYesterday() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const year = yesterday.getFullYear();
  const month = ('0' + (yesterday.getMonth() + 1)).slice(-2);
  const day = ('0' + yesterday.getDate()).slice(-2);
  return `${year}/${month}/${day}`;
}

/**
 * 要約テキストをSlack用にフォーマットする
 * Slackでは *テキスト* で太字になります
 * @param {string} summary 要約テキスト
 * @return {string} Slack用にフォーマットされた要約テキスト
 */
function formatSummaryForSlack(summary) {
  // 行ごとに分割
  const lines = summary.split('\n');
  const formattedLines = [];
  
  for (const line of lines) {
    // 各行を確認
    if (line.startsWith('・**') && line.includes(':**')) {
      // 「・**見出し:** 内容」形式を「*見出し:* 内容」形式に変換
      const formattedLine = line.replace(/・\*\*(.+?):\*\*/, '*$1:*');
      formattedLines.push(formattedLine);
    } else {
      // それ以外の行はそのまま
      formattedLines.push(line);
    }
  }
  
  return formattedLines.join('\n');
}

// --- デバッグ用関数 ---
/**
 * 特定のメール（手動テスト用）を処理する関数
 * Gmailでテストしたいメールを開き、URLの末尾の長い英数字（Message ID）を使う
 */
function testSpecificEmail() {
  // GmailApp.search はGmailThreadを返す
  const threads = GmailApp.search('subject:"Benedict\'s Newsletter: No. 585"');
  
  if (threads.length > 0) {
    // スレッドからメッセージを取得
    const messages = threads[0].getMessages();
    
    if (messages.length > 0) {
      const message = messages[0]; // 最初のメッセージを取得
      
      Logger.log(`Processing specific message: ${message.getSubject()}`);
      const originalBody = message.getPlainBody();
      
      if (!originalBody) {
        Logger.log('Could not get email body.');
        return;
      }

   try {
       const geminiResult = callGeminiApi(originalBody);
      //  Logger.log(`Gemini Result: ${JSON.stringify(geminiResult)}`); // 結果をログに出力
       Logger.log(`Gemini Result translation: ${JSON.stringify(geminiResult.translation)}`);
       Logger.log(`Gemini Result summary: ${JSON.stringify(geminiResult.summary)}`);
       if (geminiResult && geminiResult.translation && geminiResult.summary) {
           // Slack投稿（テストではコメントアウトしても良い）
           
           const postResult = postToSlack(
               `【テスト要約】\n${geminiResult.summary}`,
               SLACK_CHANNEL
           );
           Logger.log(`Slack Post Result: ${JSON.stringify(postResult)}`);

           if (postResult && postResult.ok && postResult.ts) {
               const threadMessage = `【テスト日本語訳】\n${geminiResult.translation}\n\n---\n\n【テスト原文】\n${originalBody}`;
               postToSlack(threadMessage, SLACK_CHANNEL, postResult.ts);
           } else if (postResult && postResult.ok && postResult.ts === null) {
               Logger.log("メイン投稿は成功しましたが、tsが取得できなかったためスレッド投稿はスキップします。(Webhookの制限の可能性)");
               // tsが取れない場合は、別メッセージとして投稿するなどの代替案
               // const threadMessage = `【テスト日本語訳】...\n【テスト原文】... (元メッセージ: ${message.getSubject()})`;
               // postToSlack(threadMessage, SLACK_CHANNEL);
           }
           
       }
   } catch (error) {
       Logger.log(`Error during test processing: ${error}`);
   }
      } else {
      Logger.log('No messages found in thread.');
    }
  } else {
    Logger.log('No threads found with that subject.');
  }
}
