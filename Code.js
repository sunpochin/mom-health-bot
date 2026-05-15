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

  // 檢查是否至少有一組 (3個數字) 或 兩組 (6個數字)
  if (nums && nums.length >= 3) {
    const sys1 = parseInt(nums[0], 10);
    const dia1 = parseInt(nums[1], 10);
    const pul1 = parseInt(nums[2], 10);
    
    let sys2 = "";
    let dia2 = "";
    let pul2 = "";
    
    let avgSys = sys1;
    let avgDia = dia1;
    let avgPul = pul1;

    // 若有兩組數據
    if (nums.length >= 6) {
      sys2 = parseInt(nums[3], 10);
      dia2 = parseInt(nums[4], 10);
      pul2 = parseInt(nums[5], 10);

      // 3. 計算平均值
      avgSys = Math.round((sys1 + sys2) / 2);
      avgDia = Math.round((dia1 + dia2) / 2);
      avgPul = Math.round((pul1 + pul2) / 2);
    }

    // 4. 根據 722 原則評估狀態
    let status = "✅ 正常";
    if (avgSys >= 130 || avgDia >= 80) {
      status = "🔴 偏高";
    } else if (avgSys < 110 && avgDia < 60) {
      status = "⚠️ 偏低";
    } else if (avgSys < 110 && avgDia >= 60) {
      status = "✅ 正常 (偏低點)";
    } else if (avgSys >= 110 && avgDia < 60) {
      status = "⚠️ 舒張壓偏低";
    } else {
      status = "✅ 正常";
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

    // 7. 取得最近三天的紀錄並回覆 LINE
    // Google Sheets 的 getDateRange().getValues() 取得的日期可能是 Date 物件
    const data = sheet.getDataRange().getValues();
    let historyMsg = "\n\n【最近三天紀錄】";
    let datesFound = [];
    let recentRecords = [];
    
    // 從最後一行往上找 (包含剛寫入的那行)
    // 假設第0行是標題，所以 i >= 1
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      let rowDate = row[0]; // A欄: 日期
      
      // 如果 Google Sheet 讀出來的是 Date 物件，將其轉為字串比較
      if (rowDate instanceof Date) {
        rowDate = (rowDate.getMonth() + 1) + '/' + rowDate.getDate();
      } else {
        rowDate = String(rowDate); // 確保轉為字串
      }
      
      if (!datesFound.includes(rowDate)) {
        if (datesFound.length >= 3) {
          break; // 已經收集到三個不同日期的紀錄
        }
        datesFound.push(rowDate);
      }
      
      const rPeriod = row[1];
      const rAvgSys = row[8];
      const rAvgDia = row[9];
      const rStatus = row[11];
      recentRecords.unshift(`${rowDate} ${rPeriod} ${rAvgSys}/${rAvgDia} ${rStatus}`);
    }

    if (recentRecords.length > 0) {
      historyMsg += "\n" + recentRecords.join("\n");
    } else {
      historyMsg = "";
    }

    const firstLine = `收到了！這筆血壓是 ${avgSys} / ${avgDia}`;
    replyToLine(replyToken, `${firstLine}\n✅ 已成功紀錄 (${period})\n狀態：${status}${historyMsg}`);
  } else {
    // 數字不夠 3 個的錯誤提示
    replyToLine(replyToken, "格式似乎不對喔！請確認是否有提供至少一組血壓數據。");
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
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
    },
    method: 'post',
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}
