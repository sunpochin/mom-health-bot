// 從腳本屬性中安全地取得 LINE Channel Access Token
const LINE_CHANNEL_ACCESS_TOKEN = (typeof PropertiesService !== 'undefined')
  ? PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN')
  : '';

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
  const parseResult = parseBpEntries(userMessage);
  if (!parseResult.entries.length) {
    if (parseResult.errorCode === 'missing_period_for_backfill') {
      replyToLine(replyToken, [
        "Perlu keterangan waktu: tulis Pagi/Malam atau ☀️/🌙 untuk data yang ada tanggal atau data lama.",
        "請補上時段：有日期的補登資料，請寫早/晚或 ☀️/🌙。"
      ]);
      return ContentService.createTextOutput("Success");
    }

    replyToLine(replyToken, [
      "Format salah.\nKirim 2 set data tekanan darah, misalnya:\n🌙 5/15\n128/65/75 | 123/63/73",
      "格式不對。\n請提供兩組血壓資料，例如：\n🌙 5/15\n128/65/75 | 123/63/73"
    ]);
    return ContentService.createTextOutput("Success");
  }
  const entries = parseResult.entries;

  // 1. 取得 Google Sheet 物件
  const spreadsheetId = '1EZJzRoOBkWDnaD3hUEeZGHIWv5zh5slsgpI3hzQCDKM';
  const targetGid = 2143792150;

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet
    .getSheets()
    .find(s => s.getSheetId() === targetGid);

  if (!sheet) {
    throw new Error('找不到指定的工作表分頁 gid: ' + targetGid);
  }

  const savedEntries = entries.map(entry => appendEntry(sheet, entry));
  const latestEntry = savedEntries[savedEntries.length - 1];
  const summaries = getRecentSummary(sheet);

  const idBlock = buildReplyBlock("id", savedEntries, latestEntry, summaries);
  const zhBlock = buildReplyBlock("zh", savedEntries, latestEntry, summaries);

  replyToLine(replyToken, [idBlock, zhBlock]);

  return ContentService.createTextOutput("Success");
}

function parseBpEntries(message) {
  const lines = message
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean);

  const entries = [];
  let pendingDate = null;
  let pendingPeriod = detectPeriod(message);
  let foundAnyDate = false;
  let bpLineCount = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const headerInfo = parseHeaderLine(line);
    if (headerInfo.dateString) {
      pendingDate = headerInfo.dateString;
      foundAnyDate = true;
    }
    if (headerInfo.period) {
      pendingPeriod = headerInfo.period;
    }

    const bpLine = extractBpLine(line);
    if (!bpLine) continue;
    bpLineCount++;

    if (!pendingPeriod && (pendingDate || foundAnyDate || bpLineCount > 1)) {
      return { entries: [], errorCode: 'missing_period_for_backfill' };
    }

    entries.push({
      dateString: pendingDate || getTodayDateString(),
      period: pendingPeriod || getDefaultPeriod(),
      sys1: bpLine.sys1,
      dia1: bpLine.dia1,
      pul1: bpLine.pul1,
      sys2: bpLine.sys2,
      dia2: bpLine.dia2,
      pul2: bpLine.pul2
    });
  }

  if (entries.length) return { entries: entries, errorCode: null };

  const fallbackLine = extractBpLine(message);
  if (!fallbackLine) return { entries: [], errorCode: 'invalid_format' };

  if (!pendingPeriod && foundAnyDate) {
    return { entries: [], errorCode: 'missing_period_for_backfill' };
  }

  return { entries: [{
    dateString: getTodayDateString(),
    period: detectPeriod(message) || getDefaultPeriod(),
    sys1: fallbackLine.sys1,
    dia1: fallbackLine.dia1,
    pul1: fallbackLine.pul1,
    sys2: fallbackLine.sys2,
    dia2: fallbackLine.dia2,
    pul2: fallbackLine.pul2
  }], errorCode: null };
}

function parseHeaderLine(line) {
  const dateMatch = line.match(/(?:^|\s)(\d{1,2}\/\d{1,2})(?:\s|$)/);
  return {
    dateString: dateMatch ? dateMatch[1] : null,
    period: detectPeriod(line)
  };
}

function detectPeriod(text) {
  if (/(🌙|晚|malam|night|pm)/i.test(text)) return "晚";
  if (/(☀️|早|pagi|morning|am)/i.test(text)) return "早";
  return null;
}

function getDefaultPeriod(date) {
  const hour = getTaipeiHour(date);
  return hour >= 12 ? "晚" : "早";
}

function extractBpLine(text) {
  const match = text.match(/(\d{2,3})\s*\/\s*(\d{2,3})\s*\/\s*(\d{2,3})\s*\|\s*(\d{2,3})\s*\/\s*(\d{2,3})\s*\/\s*(\d{2,3})/);
  if (!match) return null;

  return {
    sys1: parseInt(match[1], 10),
    dia1: parseInt(match[2], 10),
    pul1: parseInt(match[3], 10),
    sys2: parseInt(match[4], 10),
    dia2: parseInt(match[5], 10),
    pul2: parseInt(match[6], 10)
  };
}

function getTodayDateString() {
  if (typeof Utilities !== 'undefined') {
    return Utilities.formatDate(new Date(), 'Asia/Taipei', 'M/d');
  }

  return formatTaipeiParts(new Date()).dateString;
}

function getTaipeiHour(date) {
  if (typeof Utilities !== 'undefined') {
    return parseInt(Utilities.formatDate(date || new Date(), 'Asia/Taipei', 'H'), 10);
  }

  return formatTaipeiParts(date || new Date()).hour;
}

function formatTaipeiParts(date) {
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Taipei',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    hour12: false
  });
  const parts = formatter.formatToParts(date);
  const month = parts.find(part => part.type === 'month').value;
  const day = parts.find(part => part.type === 'day').value;
  const hour = parseInt(parts.find(part => part.type === 'hour').value, 10);

  return {
    dateString: `${month}/${day}`,
    hour: hour
  };
}

function appendEntry(sheet, entry) {
  const avgSys = Math.round((entry.sys1 + entry.sys2) / 2);
  const avgDia = Math.round((entry.dia1 + entry.dia2) / 2);
  const avgPul = Math.round((entry.pul1 + entry.pul2) / 2);
  const statusObj = getBpStatus(avgSys, avgDia);
  const hasReminder = isDuplicateEntry(sheet, entry);

  sheet.appendRow([
    entry.dateString, entry.period,
    entry.sys1, entry.dia1, entry.pul1,
    entry.sys2, entry.dia2, entry.pul2,
    avgSys, avgDia, avgPul,
    statusObj.zh
  ]);

  return {
    dateString: entry.dateString,
    period: entry.period,
    sys1: entry.sys1,
    dia1: entry.dia1,
    pul1: entry.pul1,
    sys2: entry.sys2,
    dia2: entry.dia2,
    pul2: entry.pul2,
    avgSys: avgSys,
    avgDia: avgDia,
    avgPul: avgPul,
    hasReminder: hasReminder,
    statusObj: statusObj
  };
}

function isDuplicateEntry(sheet, entry) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const lastValues = sheet.getRange(lastRow, 1, 1, 8).getValues()[0];
  let lastDate = lastValues[0];
  if (lastDate instanceof Date) {
    lastDate = (lastDate.getMonth() + 1) + '/' + lastDate.getDate();
  }

  return (
    lastDate == entry.dateString &&
    lastValues[1] == entry.period &&
    lastValues[2] == entry.sys1 &&
    lastValues[3] == entry.dia1 &&
    lastValues[4] == entry.pul1 &&
    lastValues[5] == entry.sys2 &&
    lastValues[6] == entry.dia2 &&
    lastValues[7] == entry.pul2
  );
}

function buildReplyBlock(lang, savedEntries, latestEntry, summaries) {
  const isZh = lang === "zh";
  const periodLabel = latestEntry.period === "早"
    ? (isZh ? "早" : "Pagi")
    : (isZh ? "晚" : "Malam");
  const duplicateWarning = latestEntry.hasReminder
    ? (isZh
      ? "\n⚠️ (此筆資料與上一筆完全相同)\n"
      : "\n⚠️ (Data ini sama dengan sebelumnya / 此資料與上一筆相同)\n")
    : "";
  const countLine = savedEntries.length > 1
    ? (isZh
      ? `本次共新增 ${savedEntries.length} 筆`
      : `Total ${savedEntries.length} catatan ditambahkan`)
    : null;
  const latestLabel = isZh ? "最新一筆" : "Catatan terbaru";
  const avgLabel = isZh ? "平均" : "Rata-rata";
  const statusLabel = isZh ? "狀態" : "Status";
  const summaryTitle = isZh ? "【最近三天紀錄】" : "[Catatan 3 Hari Terakhir]";
  const summaryBody = isZh ? summaries.zh : summaries.id;
  const statusText = isZh ? latestEntry.statusObj.zh : latestEntry.statusObj.id;
  const intro = isZh
    ? `✅ 已成功紀錄 (${periodLabel})${duplicateWarning}`
    : `✅ Berhasil dicatat (${periodLabel})${duplicateWarning}`;
  const helpLines = isZh
    ? [
        `💡 判斷參考：`,
        `🔴 偏高：>= 135 / 85`,
        `⚠️ 偏低：兩者都低於 90 / 60`,
        `⚠️ 收縮壓低：上面低於 90`,
        `⚠️ 舒張壓低：下面低於 60`
      ]
    : [
        `💡 Referensi Status:`,
        `🔴 Tinggi: >= 135 / 85`,
        `⚠️ Rendah: Keduanya < 90 / 60`,
        `⚠️ Sistolik Rendah: Atas < 90`,
        `⚠️ Diastolik Rendah: Bawah < 60`
      ];

  return [
    intro,
    countLine,
    `${latestLabel}: ${latestEntry.dateString} ${periodLabel}`,
    `${avgLabel}: ${latestEntry.avgSys} / ${latestEntry.avgDia}`,
    `${statusLabel}: ${statusText}`,
    ``,
    summaryTitle,
    summaryBody,
    ``,
    helpLines.join('\n')
  ].filter(Boolean).join('\n');
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
 * 發送回覆訊息給 LINE 的函式 (支援多則訊息)
 */
function replyToLine(replyToken, messages) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  
  // 確保傳入的是陣列
  const messageArray = Array.isArray(messages) ? messages : [messages];
  
  const payload = {
    replyToken: replyToken,
    messages: messageArray.map(msg => ({ type: 'text', text: msg }))
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

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    buildReplyBlock,
    detectPeriod,
    extractBpLine,
    getBpStatus,
    getDefaultPeriod,
    getTaipeiHour,
    getTodayDateString,
    parseBpEntries,
    parseHeaderLine
  };
}
