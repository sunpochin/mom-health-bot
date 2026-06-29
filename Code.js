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
      replyToLine(replyToken, "Perlu keterangan waktu: tulis Pagi/Malam atau ☀️/🌙 untuk data yang ada tanggal atau data lama.");
      return ContentService.createTextOutput("Success");
    }

    replyToLine(replyToken, "Format salah.\nKirim 2 set data tekanan darah, misalnya:\n🌙 5/15\n128/65/75 | 123/63/73");
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

  // 僅回傳印尼文的單條訊息
  replyToLine(replyToken, idBlock);

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
  const timeStr = getTaipeiTimeString(date);
  const hour = parseInt(timeStr.split(':')[0], 10);
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

function getTaipeiTimeString(date) {
  if (typeof Utilities !== 'undefined') {
    return Utilities.formatDate(date || new Date(), 'Asia/Taipei', 'HH:mm');
  }
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Taipei',
    hour: 'numeric',
    minute: 'numeric',
    hour12: false
  });
  const parts = formatter.formatToParts(date || new Date());
  const hour = parts.find(part => part.type === 'hour').value.padStart(2, '0');
  const minute = parts.find(part => part.type === 'minute').value.padStart(2, '0');
  return `${hour}:${minute}`;
}

function appendEntry(sheet, entry) {
  const avgSys = Math.round((entry.sys1 + entry.sys2) / 2);
  const avgDia = Math.round((entry.dia1 + entry.dia2) / 2);
  const avgPul = Math.round((entry.pul1 + entry.pul2) / 2);
  
  // 確保寫入日期包含當前時間戳記 (如果原本沒有包含時間且是今日登載)
  let dateToSave = entry.dateString;
  if (!dateToSave.includes(' ') && !dateToSave.includes(':')) {
    const todayString = getTodayDateString();
    if (dateToSave === todayString) {
      dateToSave = `${dateToSave} ${getTaipeiTimeString()}`;
    }
  }
  
  entry.dateString = dateToSave;
  
  const statusObj = getBpStatus(avgSys, avgDia);
  const hasReminder = isDuplicateEntry(sheet, entry);

  sheet.appendRow([
    dateToSave, entry.period,
    entry.sys1, entry.dia1, entry.pul1,
    entry.sys2, entry.dia2, entry.pul2,
    avgSys, avgDia, avgPul,
    statusObj.zh
  ]);

  return {
    dateString: dateToSave,
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

  // 使用 getDisplayValues() 避免 Date 物件轉換時區產生的偏差
  const lastValues = sheet.getRange(lastRow, 1, 1, 8).getDisplayValues()[0];
  let lastDate = lastValues[0];
  lastDate = lastDate.split(' ')[0]; // 取出日期部分，例如 "6/29"

  const entryDateOnly = entry.dateString.split(' ')[0];

  return (
    lastDate == entryDateOnly &&
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
  const summaryTitle = isZh ? "【最近四筆紀錄】" : "[4 Catatan Terakhir]";
  const summaryBody = isZh ? summaries.zh : summaries.id;
  const statusText = isZh ? latestEntry.statusObj.zh : latestEntry.statusObj.id;

  // 嘗試從 latestEntry.dateString 中解析出時間戳記
  const timeMatch = latestEntry.dateString.match(/\s(\d{2}:\d{2})/);
  const timeSuffix = timeMatch ? ` - ${timeMatch[1]}` : "";

  const intro = isZh
    ? `✅ 已成功紀錄 (${periodLabel}${timeSuffix})${duplicateWarning}`
    : `✅ Berhasil dicatat (${periodLabel}${timeSuffix})${duplicateWarning}`;

  const helpLines = isZh
    ? [
        `💡 判斷參考 (媽媽專用)：`,
        `🔴 極高危險：>= 180 / 120 (需立即複測)`,
        `🔴 明顯偏高：>= 160 / 100`,
        `⚠️ 偏高：>= 135 / 85`,
        `🟡 偏高觀察：130-134 或 80-84`,
        `✅ 正常：110-129 / 60-79`,
        `✅ 正常 (舒張壓偏低點)：110-129 / 55-59`,
        `✅ 正常 (偏低點)：100-109 / >= 55`,
        `⚠️ 偏低：90-99 或 50-54`,
        `🔴 明顯偏低：< 90 或 < 50`
      ]
    : [
        `💡 Referensi Status (Khusus Ibu):`,
        `🔴 Bahaya Sangat Tinggi: >= 180 / 120 (ukur ulang segera)`,
        `🔴 Cukup Tinggi: >= 160 / 100`,
        `⚠️ Tinggi: >= 135 / 85`,
        `🟡 Observasi (Agak tinggi): 130-134 atau 80-84`,
        `✅ Normal: 110-129 / 60-79`,
        `✅ Normal (Diastolik agak rendah): 110-129 / 55-59`,
        `✅ Normal (Agak rendah): 100-109 / >= 55`,
        `⚠️ Rendah: 90-99 atau 50-54`,
        `🔴 Sangat Rendah: < 90 atau < 50`
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
 * 針對長輩腦動脈瘤與防跌倒的客製化標準：
 * 1. 極高危險：SBP >= 180 或 DBP >= 120
 * 2. 明顯偏高：SBP >= 160 或 DBP >= 100
 * 3. 偏高：SBP >= 135 或 DBP >= 85
 * 4. 偏高觀察：SBP 130-134 或 DBP 80-84
 * 5. 明顯偏低：SBP < 90 或 DBP < 50
 * 6. 偏低：SBP 90-99 或 DBP 50-54
 * 7. 正常 (舒張壓偏低點)：SBP 110-129 且 DBP 55-59
 * 8. 正常 (偏低點)：SBP 100-109 且 DBP >= 55
 * 9. 正常：SBP 110-129 且 DBP 60-79
 */
function getBpStatus(sys, dia) {
  // 1. 高血壓端 (由高到低判斷)
  if (sys >= 180 || dia >= 120) {
    return { 
      zh: "🔴 極高危險，需立即複測", 
      id: "🔴 Bahaya Sangat Tinggi, ukur ulang segera" 
    };
  }
  if (sys >= 160 || dia >= 100) {
    return { 
      zh: "🔴 明顯偏高", 
      id: "🔴 Cukup Tinggi" 
    };
  }
  if (sys >= 135 || dia >= 85) {
    return { 
      zh: "⚠️ 偏高", 
      id: "⚠️ Tinggi" 
    };
  }
  if ((sys >= 130 && sys <= 134) || (dia >= 80 && dia <= 84)) {
    return { 
      zh: "🟡 偏高觀察", 
      id: "🟡 Observasi (Agak tinggi)" 
    };
  }

  // 2. 低血壓端 (由低到高判斷，此時 sys < 130 且 dia < 80)
  if (sys < 90 || dia < 50) {
    return { 
      zh: "🔴 明顯偏低", 
      id: "🔴 Sangat Rendah" 
    };
  }
  if ((sys >= 90 && sys <= 99) || (dia >= 50 && dia <= 54)) {
    return { 
      zh: "⚠️ 偏低", 
      id: "⚠️ Rendah" 
    };
  }

  // 3. 正常區與細分正常偏低點
  // 情況 A：收縮壓 110–129 且 舒張壓 55–59 (例如：117 / 58.5)
  if (sys >= 110 && sys <= 129 && dia >= 55 && dia <= 59) {
    return {
      zh: "✅ 正常 (舒張壓偏低點)",
      id: "✅ Normal (Diastolik agak rendah)"
    };
  }
  // 情況 B：收縮壓 100–109 且 舒張壓 >= 55 (例如：108 / 60)
  if (sys >= 100 && sys <= 109 && dia >= 55) {
    return {
      zh: "✅ 正常 (偏低點)",
      id: "✅ Normal (Agak rendah)"
    };
  }
  // 情況 C：理想/穩定狀態 SBP 110-129 且 DBP 60-79
  if (sys >= 110 && sys <= 129 && dia >= 60 && dia <= 79) {
    return { 
      zh: "✅ 正常", 
      id: "✅ Normal" 
    };
  }
  
  // 預防極端邊界情況的預設回傳值
  return { 
    zh: "✅ 正常", 
    id: "✅ Normal" 
  };
}

/**
 * 安全解析並格式化試算表中的日期與時間 (傳回 M/D HH:mm 格式字串)
 */
function formatDateTime(val) {
  if (val instanceof Date) {
    const m = val.getMonth() + 1;
    const d = val.getDate();
    const h = String(val.getHours()).padStart(2, '0');
    const min = String(val.getMinutes()).padStart(2, '0');
    return `${m}/${d} ${h}:${min}`;
  }
  return String(val); // 如果已經是字串，直接回傳
}

/**
 * 讀取試算表，取得並格式化最近四筆的紀錄 (分別傳回中印版本)
 */
function getRecentSummary(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { zh: "尚無紀錄", id: "No record" };

  // 直接取得最後 4 筆 (第一列是標題列，因此從 Row 2 開始)
  // 使用 getDisplayValues() 直接讀取儲存格純文字，完全繞過 Date 物件轉換時的時區偏差
  const startRow = Math.max(2, lastRow - 3);
  const data = sheet.getRange(startRow, 1, (lastRow - startRow + 1), 12).getDisplayValues();

  // 格式化中文版 (顯示包含時間戳記的日期)
  const zhSummary = data.map(row => {
    const d = formatDateTime(row[0]);
    const icon = (row[1] === "早") ? "☀️" : "🌙";
    return `${icon} ${d} (${row[1]})\n   ${row[2]}/${row[3]}/${row[4]} | ${row[5]}/${row[6]}/${row[7]}\n   ${row[11]}`;
  }).join('\n───\n');

  // 格式化印尼文版 (顯示包含時間戳記的日期)
  const idSummary = data.map(row => {
    const d = formatDateTime(row[0]);
    const pId = (row[1] === "早") ? "Pagi" : "Malam";
    const icon = (row[1] === "早") ? "☀️" : "🌙";
    const st = getBpStatus(parseInt(row[8], 10), parseInt(row[9], 10));
    return `${icon} ${d} (${pId})\n   ${row[2]}/${row[3]}/${row[4]} | ${row[5]}/${row[6]}/${row[7]}\n   ${st.id}`;
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
