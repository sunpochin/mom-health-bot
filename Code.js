// Code.test.js 不應該出現在這個 GAS 專案裡：它用 Node 的 require()/
// 解構賦值宣告跟這裡同名的頂層變數 (doPost、buildReplyBlock...)，GAS
// 沒有模組隔離、所有檔案共用同一個頂層作用域，這會造成
// "Identifier 'xxx' has already been declared" 的 SyntaxError，讓整個
// 專案完全無法執行 —— 這才是這幾輪 doPost 一直顯示 Failed 的真正原因，
// 跟訊息內容、events 陣列、try/catch 都無關。已加 .claspignore 排除它，
// 這行註解只是為了讓這次 push 產生真正的內容差異，強制觸發實際同步。
/**
 * 【暫時診斷用】直接觸碰 SpreadsheetApp 跟 UrlFetchApp，用來手動觸發
 * Google 的授權同意畫面。連結/變更 GCP project 之後，之前的授權會失效，
 * 但用編輯器手動執行 doPost() 因為沒有真正的 e 參數，會在
 * e.postData 那行就先炸掉，永遠碰不到這兩個服務，也就永遠不會跳出
 * 同意畫面。這支函式不依賴任何參數，直接呼叫這兩個服務，執行它才會
 * 真正觸發授權提示。確認權限恢復正常後就會移除這支函式。
 */
function authorizeScopes() {
  const spreadsheetId = '1EZJzRoOBkWDnaD3hUEeZGHIWv5zh5slsgpI3hzQCDKM';
  SpreadsheetApp.openById(spreadsheetId);
  UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
    headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
    muteHttpExceptions: true
  });
  console.log('✅ authorizeScopes 執行完成，如果沒有跳出授權視窗代表已經有權限了');
}

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
  // 這個專案目前沒有連結任何看得到的 GCP project，Cloud Logging 讀不到、
  // clasp tail-logs 也讀不到，導致 doPost 內部噴的例外完全看不到內容
  // (執行紀錄只顯示 Failed，沒有訊息)。在正式接上 GCP project 之前，
  // 先用「把例外文字直接回覆到 LINE」當診斷手段——不吞掉錯誤，至少
  // 傳訊息的人自己看得到炸在哪裡。取得足夠診斷資訊、確認穩定後會拿掉
  // 回覆原始錯誤文字的部分，改成只留 console.error + sendOpsAlert。
  try {
    return handleDoPost(e);
  } catch (err) {
    const stackText = (err && err.stack) || (err && err.message) || String(err);
    console.error('❌ doPost 例外: ' + stackText);
    try {
      const events = JSON.parse(e.postData.contents).events;
      const replyToken = events && events[0] && events[0].replyToken;
      if (replyToken) {
        replyToLine(replyToken, '🐛 DEBUG (暫時診斷用，之後會移除):\n' + stackText);
      }
    } catch (nestedErr) {
      // 連取得 replyToken 都失敗 (例如 e.postData 本身就是壞的)，
      // 這裡已經沒有任何管道能回報，只能靠上面的 console.error。
    }
    return ContentService.createTextOutput("Success");
  }
}

/**
 * doPost 實際處理邏輯，拆出來讓外層 try/catch 能包住整個流程。
 * @param {Object} e - 同 doPost 的參數
 * @return {TextOutput}
 */
function handleDoPost(e) {
  // 解析 LINE 傳來的 JSON 格式資料
  const events = JSON.parse(e.postData.contents).events;

  // LINE 的 webhook「驗證」測試 ping 會送出 events 為空陣列的請求；
  // 舊版程式碼直接存取 events[0] 沒檢查，遇到這種請求會在下一行對
  // undefined 取屬性直接噴例外，執行紀錄顯示 Failed 且沒有任何 log
  // (連下面 console.log 那行的參數都還沒求值完成就已經噴掉)。
  if (!events || events.length === 0) {
    return ContentService.createTextOutput("Success");
  }

  const event = events[0];

  // 【暫時診斷用】為了取得 ALERT_LINE_USER_ID 要填的值而加的一行 log，
  // 只印 event.source（userId/groupId/type），不影響任何回覆或判讀邏輯。
  // 取得 userId、設定好 Script Property 後就會移除這行並重新部署。
  console.log('👤 event.source: ' + JSON.stringify(event.source));

  // 只處理文字訊息，其餘類型一律忽略以保持群組安靜
  if (event.type !== 'message' || event.message.type !== 'text') {
    return ContentService.createTextOutput("Success");
  }

  const userMessage = event.message.text;
  const replyToken = event.replyToken;
  
  // 1. 從文字中提取兩組完整的血壓與脈搏數值 (Susi 當下登錄的格式)
  const bpLine = extractBpLine(userMessage);
  
  // 若格式不符 (數字不足或格式不對)
  if (!bpLine) {
    // 只有在訊息「看起來像是在輸入血壓」(包含數字與斜線 /) 時，才給予錯誤提示
    // 若只是一般聊天或日常禮貌用語 (例如 "Terima kasih"、"hi")，則已讀不回，保持群組安靜
    if (/\d/.test(userMessage) && /\//.test(userMessage)) {
      replyToLine(replyToken, "Format salah.\nKirim 2 set data tekanan darah, misalnya:\n🌙\n128/65/75 | 123/63/73");
    } else {
      // 【暫時診斷用】這個專案沒有連結 GCP project，Cloud Logging 讀不到，
      // 所以原本印 event.source 的 console.log 完全看不到內容。這裡改用
      // 已經證實有效的管道 (LINE 回覆) 暫時取代「已讀不回」，直接把
      // event.source 回傳給傳訊息的人，讓他能親眼看到自己的 userId。
      // 拿到 userId、設定好 ALERT_LINE_USER_ID 後就會拿掉，恢復原本
      // 的靜默行為。
      replyToLine(replyToken, '🐛 DEBUG userId (暫時診斷用):\n' + JSON.stringify(event.source));
    }
    return ContentService.createTextOutput("Success");
  }

  // 2. 確定時段 (若訊息有註明以訊息為準，否則自動依台北當下時間判定)
  const period = detectPeriod(userMessage) || getDefaultPeriod();

  // 3. 取得 Google Sheet 試算表物件
  const spreadsheetId = '1EZJzRoOBkWDnaD3hUEeZGHIWv5zh5slsgpI3hzQCDKM';
  const targetGid = 2143792150;

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet
    .getSheets()
    .find(s => s.getSheetId() === targetGid);

  if (!sheet) {
    throw new Error('找不到指定的工作表分頁 gid: ' + targetGid);
  }

  // 4. 將資料寫入試算表
  const latestEntry = appendEntry(sheet, bpLine, period);

  // 同步寫入 Supabase 資料庫 (包裝在 try/catch 中以確保 LINE Bot 服務不中斷)
  // P0.3：寫入失敗不能只 console.error 默默吞掉——雙軌資料默默分岔，
  // dashboard 會變成謊言，所以失敗要另外推播通知開發者 (fail loudly)。
  try {
    const supabaseOk = writeToSupabase(latestEntry);
    if (!supabaseOk) {
      sendOpsAlert('Supabase 寫入失敗（回應非 2xx），紀錄時間：' + latestEntry.dateString);
    }
  } catch (supabaseErr) {
    console.error('❌ 同步 Supabase 失敗:', supabaseErr.message || supabaseErr);
    sendOpsAlert('Supabase 寫入發生例外，紀錄時間：' + latestEntry.dateString);
  }

  // 4.5 危險等級血壓推播給家屬 (P0.2)：Susi 半夜量到危險值時，兒子的
  // LINE reply 對話看不到——這是獨立的 push 通知，跟 Susi 的對話完全分開。
  sendDangerAlertToFamily(latestEntry);

  // 5. 讀取最近四筆的歷史紀錄摘要
  const summaries = getRecentSummary(sheet);

  // 6. 建構印尼文的回覆訊息 (本系統只發送一條訊息給 Susi 閱讀)
  const idBlock = buildReplyBlock("id", latestEntry, summaries);

  // 7. 將訊息回傳給 LINE
  replyToLine(replyToken, idBlock);

  return ContentService.createTextOutput("Success");
}

/**
 * 偵測文字中是否包含代表早晚時段的關鍵字
 * @param {string} text - 要偵測的文字
 * @return {string|null} 回傳 "早"、"晚" 或 null
 */
function detectPeriod(text) {
  if (/(🌙|晚|malam|night|\bpm\b)/i.test(text)) return "晚";
  if (/(☀️|早|pagi|morning|\bam\b)/i.test(text)) return "早";
  return null;
}

/**
 * 根據時間獲取預設的時段 (依據台北時間小時判定)
 * @return {string} 回傳 "早" 或 "晚"
 */
function getDefaultPeriod() {
  const timeStr = getTaipeiTimeString();
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

  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Taipei',
    month: 'numeric',
    day: 'numeric'
  });
  const parts = formatter.formatToParts(new Date());
  const month = parts.find(part => part.type === 'month').value;
  const day = parts.find(part => part.type === 'day').value;
  return `${month}/${day}`;
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
 * @param {Object} bpLine - 解析後的血壓數值物件
 * @param {string} period - 判定後的早/晚時段
 * @return {Object} 寫入成功且已更新時間欄位的資料物件
 */
function appendEntry(sheet, bpLine, period) {
  const avgSys = Math.round((bpLine.sys1 + bpLine.sys2) / 2);
  const avgDia = Math.round((bpLine.dia1 + bpLine.dia2) / 2);
  const avgPul = Math.round((bpLine.pul1 + bpLine.pul2) / 2);
  
  // 當下量測，日期自動補上台北當下時間 (例如 "6/29 16:34")
  const dateToSave = `${getTodayDateString()} ${getTaipeiTimeString()}`;
  
  const statusObj = getBpStatus(avgSys, avgDia);
  const hasReminder = isDuplicateEntry(sheet, bpLine, period, dateToSave);

  // 寫入 Google Sheet 試算表
  sheet.appendRow([
    dateToSave, period,
    bpLine.sys1, bpLine.dia1, bpLine.pul1,
    bpLine.sys2, bpLine.dia2, bpLine.pul2,
    avgSys, avgDia, avgPul,
    statusObj.zh
  ]);

  return {
    dateString: dateToSave,
    period: period,
    sys1: bpLine.sys1,
    dia1: bpLine.dia1,
    pul1: bpLine.pul1,
    sys2: bpLine.sys2,
    dia2: bpLine.dia2,
    pul2: bpLine.pul2,
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
 * @param {Object} bpLine - 當前準備寫入的血壓數值物件
 * @param {string} period - 當次時段
 * @param {string} dateToSave - 當次要寫入的日期時間字串
 * @return {boolean} 是否完全重複
 */
function isDuplicateEntry(sheet, bpLine, period, dateToSave) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  // 使用 getDisplayValues() 獲取純文字，避免 Date 時區轉換造成的 Bug
  const lastValues = sheet.getRange(lastRow, 1, 1, 8).getDisplayValues()[0];
  let lastDate = lastValues[0];
  lastDate = lastDate.split(' ')[0]; // 只提取日期部分，例如 "6/29"

  const entryDateOnly = dateToSave.split(' ')[0];

  return (
    lastDate == entryDateOnly &&
    lastValues[1] == period &&
    lastValues[2] == bpLine.sys1 &&
    lastValues[3] == bpLine.dia1 &&
    lastValues[4] == bpLine.pul1 &&
    lastValues[5] == bpLine.sys2 &&
    lastValues[6] == bpLine.dia2 &&
    lastValues[7] == bpLine.pul2
  );
}

/**
 * 依據指定語言 (中/印)，動態建置回覆給 LINE 的訊息字串
 * @param {string} lang - 語言代碼，"id" 或 "zh"
 * @param {Object} latestEntry - 最新一筆紀錄
 * @param {Object} summaries - 最近四筆歷史紀錄的格式化結果
 * @return {string} 完整的訊息內容
 */
function buildReplyBlock(lang, latestEntry, summaries) {
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
    
  const summaryTitle = isZh ? "【最近四筆紀錄】" : "[4 Catatan Terakhir]";
  const summaryBody = isZh ? summaries.zh : summaries.id;

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
    ``,
    helpLines.join('\n'),
    ``,
    summaryTitle,
    summaryBody
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

  // 格式化中文版 (顯示包含時間戳記的日期，直接顯示平均收縮壓/平均舒張壓/平均脈搏)
  const zhSummary = data.map(row => {
    const icon = (row[1] === "早") ? "☀️" : "🌙";
    return `${icon} ${row[0]} (${row[1]})\n   平均：${row[8]} / ${row[9]} / ${row[10]}\n   ${row[11]}`;
  }).join('\n───\n');

  // 格式化印尼文版 (顯示包含時間戳記的日期，直接顯示平均收縮壓/平均舒張壓/平均脈搏)
  const idSummary = data.map(row => {
    const pId = (row[1] === "早") ? "Pagi" : "Malam";
    const icon = (row[1] === "早") ? "☀️" : "🌙";
    const st = getBpStatus(parseInt(row[8], 10), parseInt(row[9], 10));
    return `${icon} ${row[0]} (${pId})\n   Rata-rata: ${row[8]} / ${row[9]} / ${row[10]}\n   ${st.id}`;
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
    payload: JSON.stringify(payload),
    // 沒有這個的話，GAS 的 UrlFetchApp.fetch() 遇到 LINE 回傳非 2xx
    // (例如 replyToken 已過期或被用過、額度限制) 會直接拋例外——這跟
    // Node 測試裡 fetch 的行為不同，是我們一直沒模擬到的測試盲點。
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    console.error('❌ LINE 回覆失敗，狀態碼 ' + code + ': ' + res.getContentText());
    return false;
  }
  return true;
}

/**
 * 將血壓數據寫入 Supabase 資料表 blood_pressure_records
 * @param {Object} latestEntry - 量測後的資料物件
 * @return {boolean} 是否成功寫入。未設定 Supabase 屬性視為刻意略過，回傳 true
 *   (不算失敗，呼叫端不需要為此發警示)；實際寫入失敗才回傳 false。
 */
function writeToSupabase(latestEntry) {
  const supabaseUrl = PropertiesService.getScriptProperties().getProperty('SUPABASE_URL');
  const supabaseAnonKey = PropertiesService.getScriptProperties().getProperty('SUPABASE_ANON_KEY');

  if (!supabaseUrl || !supabaseAnonKey) {
    console.warn('⚠️ 略過 Supabase 寫入：未設定 SUPABASE_URL 或 SUPABASE_ANON_KEY 指令碼屬性');
    return true;
  }

  const url = supabaseUrl.replace(/\/$/, '') + '/rest/v1/blood_pressure_records';

  // 解析日期與時間，組合成台北時區的 ISO string (例如 "6/29 16:34" -> "2026-06-29T16:34:00+08:00")
  const now = new Date();
  const year = now.getFullYear();
  const parts = latestEntry.dateString.split(' ');
  const datePart = parts[0];
  const timePart = parts[1];

  const [mStr, dStr] = datePart.split('/');
  const [hStr, minStr] = timePart.split(':');

  const pad = (n) => String(n).padStart(2, '0');
  const isoString = `${year}-${pad(mStr)}-${pad(dStr)}T${pad(hStr)}:${pad(minStr)}:00+08:00`;

  const payload = {
    systolic: latestEntry.avgSys,
    diastolic: latestEntry.avgDia,
    pulse: latestEntry.avgPul,
    measured_at: isoString,
    source: 'line_bot',
    created_by: 'pengasuh'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'apikey': supabaseAnonKey,
      'Authorization': 'Bearer ' + supabaseAnonKey,
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    if (code >= 200 && code < 300) {
      console.log('✅ 成功寫入 Supabase');
      return true;
    }
    console.error('❌ 寫入 Supabase 失敗，狀態碼 ' + code + ': ' + res.getContentText());
    return false;
  } catch (err) {
    console.error('❌ 呼叫 Supabase API 發生異常: ' + (err.message || err));
    return false;
  }
}

/**
 * 透過 LINE Push API 主動發送訊息給指定使用者 (不同於 replyToLine 的
 * reply API，push 不需要 replyToken，可在任何時候主動發送)。
 * @param {string} userId - 收件者的 LINE userId
 * @param {string} message - 要發送的文字內容
 * @return {boolean} 是否推播成功
 */
function pushToLine(userId, message) {
  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: userId,
    messages: [{ type: 'text', text: message }]
  };

  const options = {
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
    },
    method: 'post',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('LINE push 失敗，狀態碼 ' + code + ': ' + res.getContentText());
  }
  return true;
}

/**
 * 危險等級 (🔴) 血壓推播給家屬 (P0.2)。
 *
 * 為什麼需要獨立於 replyToLine 之外：Susi 半夜量到危險值時，回覆訊息只會
 * 出現在她跟 bot 的對話裡，兒子不會主動看到。這裡用 LINE Push API 主動
 * 通知一個獨立的收件人 (兒子的 LINE userId)，跟 Susi 的照護對話完全分開。
 *
 * 訊息格式刻意固定為「等級＋數值＋時間」，不夾帶其他內容 (參考
 * bahasa-tw-bot 的告警慣例：alert 訊息只帶固定分類標籤，不外洩多餘內容)。
 *
 * @param {Object} latestEntry - 量測後的資料物件 (含 statusObj、avgSys、avgDia、avgPul、dateString)
 */
function sendDangerAlertToFamily(latestEntry) {
  // 判定標準表 (appscript/blood_pressure_bot_docs.md) 裡，只有「極高危險」
  // 「明顯偏高」「明顯偏低」三個危險等級用 🔴，其餘 (⚠️ 偏高/偏低、🟡 觀察)
  // 都不推播——警報分級、寧缺勿濫，避免 alarm fatigue。
  const isDangerLevel = latestEntry.statusObj.zh.indexOf('🔴') === 0;
  if (!isDangerLevel) return;

  const alertUserId = PropertiesService.getScriptProperties().getProperty('ALERT_LINE_USER_ID');
  if (!alertUserId) {
    console.warn('⚠️ 略過危險血壓推播：未設定 ALERT_LINE_USER_ID 指令碼屬性');
    return;
  }

  const message = [
    '🔴 媽媽血壓警示',
    latestEntry.statusObj.zh,
    `${latestEntry.avgSys} / ${latestEntry.avgDia} / ${latestEntry.avgPul}`,
    latestEntry.dateString
  ].join('\n');

  try {
    pushToLine(alertUserId, message);
    console.log('✅ 已推播危險血壓警示給家屬');
  } catch (err) {
    console.error('❌ 推播危險血壓警示失敗: ' + (err.message || err));
  }
}

/**
 * 系統運作面失敗時通知開發者 (P0.3 fail loudly)，例如 Supabase 寫入失敗。
 * 訊息內容刻意只帶固定分類文字，不夾帶原始錯誤細節 (可能含 URL/金鑰片段)，
 * 完整錯誤還是要靠 console.error 進 Apps Script 執行紀錄查。
 * @param {string} summary - 固定分類的簡短描述
 */
function sendOpsAlert(summary) {
  const alertUserId = PropertiesService.getScriptProperties().getProperty('ALERT_LINE_USER_ID');
  if (!alertUserId) {
    console.warn('⚠️ 略過 ops 警示推播：未設定 ALERT_LINE_USER_ID 指令碼屬性');
    return;
  }

  try {
    pushToLine(alertUserId, '⚠️ 系統告警\n' + summary);
    console.log('✅ 已推播 ops 警示');
  } catch (err) {
    console.error('❌ 推播 ops 警示失敗: ' + (err.message || err));
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    doPost,
    buildReplyBlock,
    detectPeriod,
    extractBpLine,
    getBpStatus,
    getDefaultPeriod,
    getTodayDateString,
    replyToLine,
    writeToSupabase,
    pushToLine,
    sendDangerAlertToFamily,
    sendOpsAlert
  };
}
