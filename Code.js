// 請填入您在 LINE Developers 後台取得的 Channel Access Token
const LINE_CHANNEL_ACCESS_TOKEN = PropertiesService
  .getScriptProperties()
  .getProperty('LINE_CHANNEL_ACCESS_TOKEN');

  
function doPost(e) {
  // 解析 LINE 傳來的 JSON 格式資料
  const event = JSON.parse(e.postData.contents).events[0];
  
  // 只處理文字訊息
  if (event.type !== 'message' || event.message.type !== 'text') {
    return ContentService.createTextOutput("Success");
  }

  const userMessage = event.message.text;
  const replyToken = event.replyToken;

  // 1. 解析訊息內容 (判斷早晚)
  let period = "早";
  if (userMessage.match(/(晚|Night|night)/i)) {
    period = "晚";
  }

  // 2. 擷取所有數字 (過濾掉文字，只取數值)
  const nums = userMessage.match(/\d+/g);

  // 如果成功擷取到 6 個以上的數字，代表有兩組完整的 收縮/舒張/脈搏
  if (nums && nums.length >= 6) {
    const sys1 = parseInt(nums[0], 10);
    const dia1 = parseInt(nums[1], 10);
    const pul1 = parseInt(nums[2], 10);
    const sys2 = parseInt(nums[3], 10);
    const dia2 = parseInt(nums[4], 10);
    const pul2 = parseInt(nums[5], 10);

    // 3. 計算平均值
    const avgSys = Math.round((sys1 + sys2) / 2);
    const avgDia = Math.round((dia1 + dia2) / 2);
    const avgPul = Math.round((pul1 + pul2) / 2);

    // 4. 狀態評估邏輯
    let status = "✅ 正常";
    if (avgSys < 100 || avgDia < 60) {
      status = "⚠️ 偏低";
    } else if (avgSys > 140 || avgDia > 90) {
      status = "🔴 偏高";
    }

    // 5. 取得當天日期 (格式：M/D)
    const today = new Date();
    // GAS 預設時區可能需調整，加上 8 小時為台灣時間
    const twTime = new Date(today.getTime() + (8 * 60 * 60 * 1000));
    const dateString = (twTime.getUTCMonth() + 1) + '/' + twTime.getUTCDate();

    // 6. 將資料 Append 到媽媽血壓的 Google Sheet 最後一行
const spreadsheetId = '1EZJzRoOBkWDnaD3hUEeZGHIWv5zh5slsgpI3hzQCDKM';
const targetGid = 2143792150;

const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
const sheet = spreadsheet
  .getSheets()
  .find(s => s.getSheetId() === targetGid);

if (!sheet) {
  throw new Error('找不到指定的工作表分頁 gid: ' + targetGid);
}

sheet.appendRow([
  dateString, period,
  sys1, dia1, pul1,
  sys2, dia2, pul2,
  avgSys, avgDia, avgPul,
  status
]);

    // 7. 回覆 LINE 訊息，讓 Susi 或您知道紀錄成功
    replyToLine(replyToken, `✅ 已成功紀錄 (${period})\n平均：${avgSys} / ${avgDia}\n狀態：${status}`);
  } else {
    // 數字不夠 6 個的錯誤提示
    replyToLine(replyToken, "格式似乎不對喔！請確認是否有兩組完整的血壓數據。");
  }

  return ContentService.createTextOutput("Success");
}

// 發送回覆訊息給 LINE 的函式
function replyToLine(replyToken, text) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text }]
  };
  const options = {
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + 'S1g4dEKt6pSTETqnpoFEaAt8cJ3QtJiITJgtuPz3gLuM8Uml4YCF4jIQVQSi2jNIAvApoRpmGT8s1h0X79UZz4YljCJPNH3Xw3j98H+4FalaU6pvA/8RsZfEuwfYzCM7C+J4I+a2PoGcoaq7yXiXcAdB04t89/1O/w1cDnyilFU='
    },
    method: 'post',
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}