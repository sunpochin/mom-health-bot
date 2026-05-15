// 從腳本屬性中安全地取得 LINE Channel Access Token
const LINE_CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');

/**
 * 處理 LINE Webhook 傳來的 POST 請求
 */
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

    // 4. 狀態評估邏輯 (中印雙語)
    const statusObj = getBpStatus(avgSys, avgDia);

    // 5. 取得當天日期 (格式：M/D)
    const today = new Date();
    const twTime = new Date(today.getTime() + (8 * 60 * 60 * 1000));
    const dateString = (twTime.getUTCMonth() + 1) + '/' + twTime.getUTCDate();

    // 6. 將資料 Append 到指定的 Google Sheet
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
      statusObj.zh // 試算表存中文即可
    ]);

    // 7. 取得最近三天的摘要 (分別取得中印版本)
    const summaries = getRecentSummary(sheet);

    // 8. 建構雙語訊息 (分開區塊顯示)
    const periodId = (period === "早") ? "Pagi" : "Malam";
    
    // 中文區塊
    const zhBlock = [
      `✅ 已成功紀錄 (${period})`,
      `平均：${avgSys} / ${avgDia}`,
      `狀態：${statusObj.zh}`,
      ``,
      `【最近三天紀錄】`,
      summaries.zh,
      ``,
      `💡 判斷參考：`,
      `🔴 偏高：>= 135 / 85`,
      `⚠️ 偏低：兩者都低於 90 / 60`,
      `⚠️ 收縮壓低：上面低於 90`,
      `⚠️ 舒張壓低：下面低於 60`
    ].join('\n');

    // 印尼文區塊
    const idBlock = [
      `✅ Berhasil dicatat (${periodId})`,
      `Rata-rata: ${avgSys} / ${avgDia}`,
      `Status: ${statusObj.id}`,
      ``,
      `[Catatan 3 Hari Terakhir]`,
      summaries.id,
      ``,
      `💡 Referensi Status:`,
      `🔴 Tinggi: >= 135 / 85`,
      `⚠️ Rendah: Keduanya < 90 / 60`,
      `⚠️ Sistolik Rendah: Atas < 90`,
      `⚠️ Diastolik Rendah: Bawah < 60`
    ].join('\n');

    const replyMsg = idBlock + '\n\n' + '--------------------\n\n' + zhBlock;

    replyToLine(replyToken, replyMsg);
  } else {
    replyToLine(replyToken, "✅ Berhasil dicatat (Pagi)\nFormat salah!\nPerlu 2 set data.\n\n--------------------\n\n已成功紀錄 (早)\n格式似乎不對喔！\n請確認是否有兩組完整的數據。");
  }

  return ContentService.createTextOutput("Success");
}

/**
 * 依照血壓數值評估狀態 (傳回中印雙語)
 * 判斷標準：
 * 1. 偏高：收縮壓 >= 135 或 舒張壓 >= 85 (居家血壓標準)
 * 2. 偏低門檻：收縮壓 < 90 或 舒張壓 < 60
 * 3. 正常 (偏低點)：未達偏低門檻，但收縮壓 < 110 或 舒張壓 < 70
 */
function getBpStatus(sys, dia) {
  // --- 判斷偏高 ---
  if (sys >= 135 || dia >= 85) {
    return { zh: "🔴 偏高", id: "🔴 Tinggi" };
  }
  
  // --- 判斷偏低 (低於 90/60) ---
  if (sys < 90 || dia < 60) {
    // 情況 A: 收縮壓與舒張壓「兩者都低於門檻」
    if (sys < 90 && dia < 60) {
      return { zh: "⚠️ 偏低", id: "⚠️ Rendah" };
    }
    // 情況 B: 只有「收縮壓」(上面的數字) 低於 90
    if (sys < 90) {
      return { zh: "⚠️ 收縮壓偏低", id: "⚠️ Sistolik Rendah" };
    }
    // 情況 C: 只有「舒張壓」(下面的數字) 低於 60
    return { zh: "⚠️ 舒張壓偏低", id: "⚠️ Diastolik Rendah" };
  }
  
  // --- 判斷正常但數值稍低 (介於 90~110 或 60~70 之間) ---
  if (sys < 110 || dia < 70) {
    return { zh: "✅ 正常 (偏低點)", id: "✅ Normal (Agak rendah)" };
  }
  
  // --- 判斷完全正常 ---
  return { zh: "✅ 正常", id: "✅ Normal" };
}

/**
 * 讀取試算表，取得並格式化最近三天的紀錄 (分別傳回中印版本)
 */
function getRecentSummary(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { zh: "尚無紀錄", id: "No record" };

  const startRow = Math.max(2, lastRow - 15);
  const data = sheet.getRange(startRow, 1, (lastRow - startRow + 1), 12).getValues();

  // 取得最近 3 個不重複的日期
  const uniqueDates = [];
  for (let i = data.length - 1; i >= 0; i--) {
    let dateVal = data[i][0];
    if (dateVal instanceof Date) {
      dateVal = (dateVal.getMonth() + 1) + '/' + dateVal.getDate();
    }
    if (!uniqueDates.includes(dateVal)) {
      uniqueDates.push(dateVal);
    }
    if (uniqueDates.length >= 3) break;
  }

  // 過濾出符合日期的列
  const filteredRows = data.filter(row => {
    let d = row[0];
    if (d instanceof Date) d = (d.getMonth() + 1) + '/' + d.getDate();
    return uniqueDates.includes(d);
  });

  // 格式化中文版
  const zhSummary = filteredRows.map(row => {
    let d = row[0];
    if (d instanceof Date) d = (d.getMonth() + 1) + '/' + d.getDate();
    const icon = (row[1] === "早") ? "☀️" : "🌙";
    return `${icon} ${d} ${row[1]}\n   ${row[2]}/${row[3]}/${row[4]} | ${row[5]}/${row[6]}/${row[7]}\n   ${row[11]}`;
  }).join('\n───\n');

  // 格式化印尼文版
  const idSummary = filteredRows.map(row => {
    let d = row[0];
    if (d instanceof Date) d = (d.getMonth() + 1) + '/' + d.getDate();
    const pId = (row[1] === "早") ? "Pagi" : "Malam";
    const icon = (row[1] === "早") ? "☀️" : "🌙";
    const st = getBpStatus(row[8], row[9]);
    return `${icon} ${d} ${pId}\n   ${row[2]}/${row[3]}/${row[4]} | ${row[5]}/${row[6]}/${row[7]}\n   ${st.id}`;
  }).join('\n───\n');

  return { zh: zhSummary, id: idSummary };
}

/**
 * 發送回覆訊息給 LINE 的函式
 */
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