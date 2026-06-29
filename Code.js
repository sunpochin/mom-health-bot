// 從腳本屬性中安全地取得 LINE Channel Access Token
const LINE_CHANNEL_ACCESS_TOKEN = (typeof PropertiesService !== 'undefined')
  ? PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN')
  : '';

/**
 * 處理 LINE Webhook 傳來的 POST 請求 (主要入口點)
 * @param {Object} e - Google Apps Script 傳入的 POST 請求物件，包含 postData 內容
 * @return {TextOutput} 回傳給 LINE 的成功訊息
 */
function doPost(e) {
  // 解析 LINE 傳來的 JSON 格式資料
  const event = JSON.parse(e.postData.contents).events[0];
  
  // 只處理文字訊息，其餘類型一律忽略以保持群組安靜
  if (event.type !== 'message' || event.message.type !== 'text') {
    return ContentService.createTextOutput("Success");
  }

  const userMessage = event.message.text;
  const replyToken = event.replyToken;
  
  // 解析訊息，提取出一筆或多筆血壓資料
  const parseResult = parseBpEntries(userMessage);
  
  // 若格式不符或缺少必要資訊，發送相對應的印尼文錯誤提示
  if (!parseResult.entries.length) {
    // 情況：補登歷史資料卻忘記註明早/晚時段
    if (parseResult.errorCode === 'missing_period_for_backfill') {
      replyToLine(replyToken, "Perlu keterangan waktu: tulis Pagi/Malam atau ☀️/🌙 untuk data yang ada tanggal atau data lama.");
      return ContentService.createTextOutput("Success");
    }

    // 情況：一般的格式錯誤提示
    replyToLine(replyToken, "Format salah.\nKirim 2 set data tekanan darah, misalnya:\n🌙 5/15\n128/65/75 | 123/63/73");
    return ContentService.createTextOutput("Success");
  }
  
  const entries = parseResult.entries;

  // 1. 取得 Google Sheet 試算表物件
  const spreadsheetId = '1EZJzRoOBkWDnaD3hUEeZGHIWv5zh5slsgpI3hzQCDKM';
  const targetGid = 2143792150;

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet
    .getSheets()
    .find(s => s.getSheetId() === targetGid);

  if (!sheet) {
    throw new Error('找不到指定的工作表分頁 gid: ' + targetGid);
  }

  // 2. 逐筆寫入試算表
  const savedEntries = entries.map(entry => appendEntry(sheet, entry));
  const latestEntry = savedEntries[savedEntries.length - 1]; // 取得最新寫入的那一筆
  
  // 3. 讀取最新 4 筆歷史紀錄
  const summaries = getRecentSummary(sheet);

  // 4. 建構印尼文的回覆訊息 (本系統只發送一條訊息給 Susi 閱讀)
  const idBlock = buildReplyBlock("id", savedEntries, latestEntry, summaries);

  // 5. 將訊息回傳給 LINE
  replyToLine(replyToken, idBlock);

  return ContentService.createTextOutput("Success");
}

/**
 * 解析使用者輸入的多行或單行訊息，抽離出日期、時段與兩組血壓數據
 * @param {string} message - 使用者傳入的完整文字訊息
 * @return {Object} 包含解析後的 entries 陣列與 errorCode 錯誤代碼
 */
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
    
    // 解析該行是否為標頭 (例如 "🌙 5/15")
    const headerInfo = parseHeaderLine(line);
    if (headerInfo.dateString) {
      pendingDate = headerInfo.dateString;
      foundAnyDate = true;
    }
    if (headerInfo.period) {
      pendingPeriod = headerInfo.period;
    }

    // 解析該行是否為血壓數據行 (例如 "128/65/75 | 123/63/73")
    const bpLine = extractBpLine(line);
    if (!bpLine) continue;
    bpLineCount++;

    // 歷史補登防呆：如果補登多筆或有指定日期，但卻沒有指定早/晚時段，判定為缺少時段
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

  // 若成功解析出多筆項目，直接回傳
  if (entries.length) return { entries: entries, errorCode: null };

  // 備用方案：如果使用者只傳送了單行血壓數值 (無換行)
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

/**
 * 解析單行文字，抽離出日期與時段標頭
 * @param {string} line - 單行文字
 * @return {Object} 包含 dateString (M/D 格式) 與 period (早/晚)
 */
function parseHeaderLine(line) {
  const dateMatch = line.match(/(?:^|\s)(\d{1,2}\/\d{1,2})(?:\s|$)/);
  return {
    dateString: dateMatch ? dateMatch[1] : null,
    period: detectPeriod(line)
  };
}

/**
 * 偵測文字中是否包含代表早晚時段的關鍵字
 * @param {string} text - 要偵測的文字
 * @return {string|null} 回傳 "早"、"晚" 或 null
 */
function detectPeriod(text) {
  if (/(🌙|晚|malam|night|pm)/i.test(text)) return "晚";
  if (/(☀️|早|pagi|morning|am)/i.test(text)) return "早";
  return null;
}

/**
 * 根據時間獲取預設的時段 (依據台北時間小時判定)
 * @param {Date} [date] - 基準時間，未傳入則預設為當前台北時間
 * @return {string} 回傳 "早" 或 "晚"
 */
function getDefaultPeriod(date) {
  const timeStr = getTaipeiTimeString(date);
  const hour = parseInt(timeStr.split(':')[0], 10);
  return hour >= 12 ? "晚" : "早";
}

/**
 * 從單行文字中提取兩組完整的血壓與脈搏數值
 * @param {string} text - 量測文字，如 "128/65/75 | 123/63/73"
 * @return {Object|null} 包含兩組 Sys/Dia/Pul 的物件，格式不符則傳回 null
 */
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

/**
 * 獲取當前台灣時間的 M/d 日期字串 (例如 "6/29")
 * @return {string} 日期字串
 */
function getTodayDateString() {
  if (typeof Utilities !== 'undefined') {
    return Utilities.formatDate(new Date(), 'Asia/Taipei', 'M/d');
  }

  return formatTaipeiParts(new Date()).dateString;
}

/**
 * 獲取台北時間的小時數 (24 小時制)
 * @param {Date} [date] - 基準時間，預設為當前時間
 * @return {number} 小時數
 */
function getTaipeiHour(date) {
  if (typeof Utilities !== 'undefined') {
    return parseInt(Utilities.formatDate(date || new Date(), 'Asia/Taipei', 'H'), 10);
  }

  return formatTaipeiParts(date || new Date()).hour;
}

/**
 * 本地開發/測試環境的備份時區轉換方案 (當無 Google Utilities 時使用)
 * @param {Date} date - 基準時間
 * @return {Object} 包含 dateString 與 hour
 */
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

/**
 * 獲取台北時間的 HH:mm 精準時間字串 (例如 "16:34")
 * @param {Date} [date] - 基準時間，預設為當前時間
 * @return {string} 時間字串
 */
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

/**
 * 將單筆血壓紀錄寫入試算表，並補全當下時間戳記
 * @param {Sheet} sheet - Google Sheet 工作表物件
 * @param {Object} entry - 解析後的單筆資料物件
 * @return {Object} 寫入成功且已更新時間欄位的資料物件
 */
function appendEntry(sheet, entry) {
  const avgSys = Math.round((entry.sys1 + entry.sys2) / 2);
  const avgDia = Math.round((entry.dia1 + entry.dia2) / 2);
  const avgPul = Math.round((entry.pul1 + entry.pul2) / 2);
  
  // 補全時間戳記：若為今日量測且原資料無時間，補上台北當下時間
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

  // 寫入 Google Sheet 試算表
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

/**
 * 比對試算表最後一筆紀錄，判斷是否與當次輸入完全重複
 * @param {Sheet} sheet - Google Sheet 工作表物件
 * @param {Object} entry - 當前準備寫入的資料物件
 * @return {boolean} 是否完全重複
 */
function isDuplicateEntry(sheet, entry) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  // 使用 getDisplayValues() 獲取純文字，避免 Date 時區轉換造成的 Bug
  const lastValues = sheet.getRange(lastRow, 1, 1, 8).getDisplayValues()[0];
  let lastDate = lastValues[0];
  lastDate = lastDate.split(' ')[0]; // 只提取日期部分，例如 "6/29"

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

/**
 * 依據指定語言 (中/印)，動態建置回覆給 LINE 的訊息字串
 * @param {string} lang - 語言代碼，"id" 或 "zh"
 * @param {Array} savedEntries - 本次新增的紀錄陣列
 * @param {Object} latestEntry - 最新一筆紀錄
 * @param {Object} summaries - 最近四筆歷史紀錄的格式化結果
 * @return {string} 完整的訊息內容
 */
function buildReplyBlock(lang, savedEntries, latestEntry, summaries) {
  const isZh = lang === "zh";
  const periodLabel = latestEntry.period === "早"
    ? (isZh ? "早" : "Pagi")
    : (isZh ? "晚" : "Malam");
  
  // 若為重複紀錄，產生相對應的雙語重複警告
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

  // 從日期中提取時間戳記
  const timeMatch = latestEntry.dateString.match(/\s(\d{2}:\d{2})/);
  const timeSuffix = timeMatch ? ` - ${timeMatch[1]}` : "";

  const intro = isZh
    ? `✅ 已成功紀錄 (${periodLabel}${timeSuffix})${duplicateWarning}`
    : `✅ Berhasil dicatat (${periodLabel}${timeSuffix})${duplicateWarning}`;

  // 客製化對照標準說明文字
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
 * @param {number} sys - 平均收縮壓
 * @param {number} dia - 平均舒張壓
 * @return {Object} 包含 zh 與 id 狀態的物件
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
 * @param {Date|string} val - 日期物件或字串
 * @return {string} 格式化後的時間字串
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
 * @param {Sheet} sheet - Google Sheet 工作表物件
 * @return {Object} 包含 zh 與 id 格式化歷史紀錄的物件
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
 * @param {string} replyToken - LINE 回覆用 Token
 * @param {string|Array} messages - 單則訊息字串或多則訊息字串陣列
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
