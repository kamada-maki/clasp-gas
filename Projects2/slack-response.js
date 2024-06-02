const properties = PropertiesService.getScriptProperties();
const SLACK_TOKEN = properties.getProperty("SLACK_TOKEN");
const SPREADSHEET_ID = properties.getProperty("SPREADSHEET_ID");
const SHEET_NAME = properties.getProperty("SHEET_NAME");
const OPENAI_API_KEY = properties.getProperty("OPENAI_API_KEY");
const VERIFICATION_TOKEN = properties.getProperty("VERIFICATION_TOKEN");

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // ロックを30秒間待つ

    logToSheet("Received request: " + JSON.stringify(e.postData.contents));
    const json = JSON.parse(e.postData.contents);

    // Verification token check
    if (json.token !== VERIFICATION_TOKEN) {
      throw new Error("Invalid token.");
    }

    // URL verification request
    if (json.type === "url_verification") {
      logToSheet("URL verification request received.");
      lock.releaseLock();
      return ContentService.createTextOutput(json.challenge);
    }

    // Event handling
    if (
      json.event &&
      (json.event.type === "message" || json.event.type === "app_mention") &&
      !json.event.bot_id
    ) {
      const eventId = json.event.client_msg_id || json.event.ts; // 確実に一意のIDを取得
      logToSheet("Message event received: " + JSON.stringify(json.event));

      // Check if the event has already been processed
      if (isEventProcessed(eventId)) {
        logToSheet("Event already processed: " + eventId);
        lock.releaseLock();
        return ContentService.createTextOutput("OK");
      }

      // マークは最初に実行される
      markEventAsProcessed(eventId);

      const question = json.event.text;
      const channel = json.event.channel;
      const threadTs = json.event.thread_ts || json.event.ts; // スレッドのタイムスタンプを取得
      const data = getDataFromSheet();
      const answer = getAnswerFromGPT(question, data);

      logToSheet("Answer generated: " + answer);
      sendAnswerToSlack(channel, answer, threadTs);
    }

    lock.releaseLock();
    return ContentService.createTextOutput("OK");
  } catch (error) {
    // エラーログを記録
    logToSheet("Error: " + error.message);
    lock.releaseLock();
    return ContentService.createTextOutput("Error: " + error.message);
  }
}

function isEventProcessed(eventId) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("processed_events");

  // シートが存在するか確認
  if (!sheet) {
    logToSheet("Error: processed_events sheet not found");
    throw new Error("processed_events sheet not found");
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === eventId) {
      return true;
    }
  }
  return false;
}

function markEventAsProcessed(eventId) {
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("processed_events");

  // シートが存在するか確認
  if (!sheet) {
    logToSheet("Error: processed_events sheet not found");
    throw new Error("processed_events sheet not found");
  }

  sheet.appendRow([eventId, new Date()]);
}

function getDataFromSheet() {
  logToSheet("Fetching data from sheet...");
  const sheet =
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);

  // シートが正しく取得できているか確認
  if (!sheet) {
    logToSheet("Error: Sheet not found");
    throw new Error("Sheet not found");
  }

  const data = sheet.getDataRange().getValues();
  logToSheet("Data fetched: " + JSON.stringify(data));
  return data.map((row) => ({ category: row[0], content: row[1] }));
}

function getAnswerFromGPT(question, data) {
  logToSheet("Generating answer using GPT...");
  const url = "https://api.openai.com/v1/chat/completions"; // エンドポイントを変更
  const prompt = generatePrompt(question, data);

  const payload = {
    model: "gpt-4", // 使用するモデル
    messages: [
      { role: "system", content: "You are a helpful assistant." },
      { role: "user", content: prompt },
    ],
    max_tokens: 150,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY,
    },
    payload: JSON.stringify(payload),
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  logToSheet("GPT response: " + JSON.stringify(json));
  return json.choices[0].message.content.trim();
}

function generatePrompt(question, data) {
  logToSheet("Generating prompt...");
  let prompt = `以下のデータに基づいて、質問に対する最適な回答を提供してください。\n\nデータ:\n`;
  data.forEach((item) => {
    prompt += `${item.category}: ${item.content}\n`;
  });
  prompt += `\n質問: ${question}\n回答: `;
  logToSheet("Prompt generated: " + prompt);
  return prompt;
}

function sendAnswerToSlack(channel, text, threadTs) {
  logToSheet("Sending answer to Slack...");
  const url = "https://slack.com/api/chat.postMessage";
  const payload = {
    channel: channel,
    text: text,
    thread_ts: threadTs, // スレッドのタイムスタンプを追加
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + SLACK_TOKEN,
    },
    payload: JSON.stringify(payload),
  };

  const response = UrlFetchApp.fetch(url, options);
  logToSheet("Slack response: " + response.getContentText());
}

function logToSheet(message) {
  Logger.log("Logging to sheet: " + message);
  try {
    const logSheet =
      SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("log");

    // シートが存在するか確認
    if (!logSheet) {
      Logger.log("Error: log sheet not found");
      throw new Error("log sheet not found");
    }

    logSheet.appendRow([new Date(), message]);
    Logger.log("Log successfully written to sheet.");
  } catch (error) {
    Logger.log("Error logging to sheet: " + error.message);
  }
}
