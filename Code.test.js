const assert = require('assert');
const Code = require('./Code.js');
const {
  doPost,
  buildReplyBlock,
  detectPeriod,
  extractBpLine,
  getDefaultPeriod,
  getBpStatus,
  getTaipeiHour,
  parseBpEntries,
  parseHeaderLine,
  pushToLine,
  sendDangerAlertToFamily,
  sendOpsAlert,
  writeToSupabase
} = Code;

// Minimal Apps Script global mocks, only for the push/Supabase tests below.
// Code.js calls PropertiesService/UrlFetchApp directly (no `typeof` guard
// inside these functions), so they must exist as globals before calling in.
function withGasGlobals(props, fetchResponses, testFn) {
  const calls = [];
  global.PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => (Object.prototype.hasOwnProperty.call(props, key) ? props[key] : null)
    })
  };
  global.UrlFetchApp = {
    fetch: (url, options) => {
      calls.push({ url, options });
      const res = fetchResponses.shift() || { code: 200, body: '' };
      return {
        getResponseCode: () => res.code,
        getContentText: () => res.body
      };
    }
  };
  try {
    testFn(calls);
  } finally {
    delete global.PropertiesService;
    delete global.UrlFetchApp;
  }
}

function runTest(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
  } catch (error) {
    console.error(`FAIL ${name}`);
    console.error(error.stack);
    process.exitCode = 1;
  }
}

runTest('detectPeriod handles emoji and Indonesian labels', () => {
  assert.strictEqual(detectPeriod('🌙 5/13'), '晚');
  assert.strictEqual(detectPeriod('Pagi 5/13'), '早');
  assert.strictEqual(detectPeriod('random text'), null);
});

runTest('Taipei time fallback determines morning vs evening', () => {
  assert.strictEqual(getTaipeiHour(new Date('2026-05-15T01:30:00Z')), 9);
  assert.strictEqual(getTaipeiHour(new Date('2026-05-15T12:30:00Z')), 20);
  assert.strictEqual(getDefaultPeriod(new Date('2026-05-15T01:30:00Z')), '早');
  assert.strictEqual(getDefaultPeriod(new Date('2026-05-15T12:30:00Z')), '晚');
});

runTest('parseHeaderLine extracts date and period', () => {
  assert.deepStrictEqual(parseHeaderLine('🌙 5/14'), {
    dateString: '5/14',
    period: '晚'
  });
  assert.deepStrictEqual(parseHeaderLine('☀️ 5/15'), {
    dateString: '5/15',
    period: '早'
  });
});

runTest('extractBpLine only parses real blood pressure lines', () => {
  assert.deepStrictEqual(extractBpLine('128/65/75 | 123/63/73'), {
    sys1: 128,
    dia1: 65,
    pul1: 75,
    sys2: 123,
    dia2: 63,
    pul2: 73
  });
  assert.strictEqual(extractBpLine('🌙 5/13'), null);
});

runTest('parseBpEntries handles multiline multi-day message', () => {
  const sample = [
    '09:47 Susi 🌙 5/13',
    '128/65/75 | 123/63/73',
    '',
    '🌙 5/14',
    '116/60/70 | 118/57/69',
    '',
    '🌙 5/15',
    '128/61/68 | 129/59/71'
  ].join('\n');

  assert.deepStrictEqual(parseBpEntries(sample), {
    entries: [
      {
        dateString: '5/13',
        period: '晚',
        sys1: 128,
        dia1: 65,
        pul1: 75,
        sys2: 123,
        dia2: 63,
        pul2: 73
      },
      {
        dateString: '5/14',
        period: '晚',
        sys1: 116,
        dia1: 60,
        pul1: 70,
        sys2: 118,
        dia2: 57,
        pul2: 69
      },
      {
        dateString: '5/15',
        period: '晚',
        sys1: 128,
        dia1: 61,
        pul1: 68,
        sys2: 129,
        dia2: 59,
        pul2: 71
      }
    ],
    errorCode: null
  });
});

runTest('parseBpEntries falls back to today for simple single-line input', () => {
  const result = parseBpEntries('Pagi 128/65/75 | 123/63/73');
  const entries = result.entries;
  assert.strictEqual(result.errorCode, null);
  assert.strictEqual(entries.length, 1);
  assert.strictEqual(entries[0].period, '早');
  assert.deepStrictEqual(
    (({ sys1, dia1, pul1, sys2, dia2, pul2 }) => ({ sys1, dia1, pul1, sys2, dia2, pul2 }))(entries[0]),
    {
      sys1: 128,
      dia1: 65,
      pul1: 75,
      sys2: 123,
      dia2: 63,
      pul2: 73
    }
  );
  assert.match(entries[0].dateString, /^\d{1,2}\/\d{1,2}$/);
});

runTest('parseBpEntries uses Taipei-time fallback only for single immediate entry', () => {
  const result = parseBpEntries('128/65/75 | 123/63/73');
  assert.strictEqual(result.errorCode, null);
  assert.strictEqual(result.entries.length, 1);
  assert.match(result.entries[0].dateString, /^\d{1,2}\/\d{1,2}$/);
  assert.ok(['早', '晚'].includes(result.entries[0].period));
});

runTest('parseBpEntries rejects dated backfill without explicit period', () => {
  assert.deepStrictEqual(
    parseBpEntries('5/15\n128/65/75 | 123/63/73'),
    {
      entries: [],
      errorCode: 'missing_period_for_backfill'
    }
  );
});

runTest('parseBpEntries rejects multi-entry message without explicit period', () => {
  const sample = [
    '5/13',
    '128/65/75 | 123/63/73',
    '',
    '5/14',
    '116/60/70 | 118/57/69'
  ].join('\n');

  assert.deepStrictEqual(parseBpEntries(sample), {
    entries: [],
    errorCode: 'missing_period_for_backfill'
  });
});

runTest('getBpStatus classifies mild low diastolic as normal (diastolic slightly low)', () => {
  assert.deepStrictEqual(getBpStatus(117, 59), {
    zh: '✅ 正常 (舒張壓偏低點)',
    id: '✅ Normal (Diastolik agak rendah)'
  });
});

runTest('buildReplyBlock shows inserted count for batch messages', () => {
  const savedEntries = [
    {
      dateString: '5/13',
      period: '晚',
      avgSys: 126,
      avgDia: 64,
      hasReminder: false,
      statusObj: { zh: '✅ 正常', id: '✅ Normal' }
    },
    {
      dateString: '5/15',
      period: '晚',
      avgSys: 129,
      avgDia: 59,
      hasReminder: false,
      statusObj: { zh: '⚠️ 舒張壓偏低', id: '⚠️ Diastolik Rendah' }
    }
  ];
  const reply = buildReplyBlock('zh', savedEntries, savedEntries[1], {
    zh: 'mock summary zh',
    id: 'mock summary id'
  });

  assert.match(reply, /本次共新增 2 筆/);
  assert.match(reply, /最新一筆: 5\/15 晚/);
  assert.match(reply, /狀態: ⚠️ 舒張壓偏低/);
});

// P0.2 — 危險等級血壓推播給家屬

runTest('sendDangerAlertToFamily pushes a fixed-format message on a danger level', () => {
  withGasGlobals({ ALERT_LINE_USER_ID: 'U_SON' }, [{ code: 200 }], (calls) => {
    sendDangerAlertToFamily({
      statusObj: { zh: '🔴 明顯偏高', id: '🔴 Cukup Tinggi' },
      avgSys: 165, avgDia: 102, avgPul: 80,
      dateString: '7/2 03:15'
    });

    assert.strictEqual(calls.length, 1);
    const payload = JSON.parse(calls[0].options.payload);
    assert.strictEqual(payload.to, 'U_SON');
    assert.match(payload.messages[0].text, /🔴 明顯偏高/);
    assert.match(payload.messages[0].text, /165 \/ 102 \/ 80/);
    assert.match(payload.messages[0].text, /7\/2 03:15/);
  });
});

runTest('sendDangerAlertToFamily does not push for non-danger levels', () => {
  withGasGlobals({ ALERT_LINE_USER_ID: 'U_SON' }, [], (calls) => {
    sendDangerAlertToFamily({
      statusObj: { zh: '⚠️ 偏高', id: '⚠️ Tinggi' },
      avgSys: 140, avgDia: 88, avgPul: 75,
      dateString: '7/2 08:00'
    });
    assert.strictEqual(calls.length, 0);
  });
});

runTest('sendDangerAlertToFamily silently skips when ALERT_LINE_USER_ID is unset', () => {
  withGasGlobals({}, [], (calls) => {
    assert.doesNotThrow(() => sendDangerAlertToFamily({
      statusObj: { zh: '🔴 明顯偏低', id: '🔴 Sangat Rendah' },
      avgSys: 85, avgDia: 48, avgPul: 90,
      dateString: '7/2 03:15'
    }));
    assert.strictEqual(calls.length, 0);
  });
});

// P0.3 — Supabase 寫入失敗 fail loudly

runTest('writeToSupabase returns true and skips fetch when unconfigured', () => {
  withGasGlobals({}, [], (calls) => {
    const ok = writeToSupabase({ dateString: '7/2 03:15', avgSys: 120, avgDia: 70, avgPul: 70 });
    assert.strictEqual(ok, true);
    assert.strictEqual(calls.length, 0);
  });
});

runTest('writeToSupabase returns false on a non-2xx response', () => {
  withGasGlobals(
    { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_ANON_KEY: 'key' },
    [{ code: 500, body: 'server error' }],
    (calls) => {
      const ok = writeToSupabase({ dateString: '7/2 03:15', avgSys: 120, avgDia: 70, avgPul: 70 });
      assert.strictEqual(ok, false);
      assert.strictEqual(calls.length, 1);
    }
  );
});

runTest('writeToSupabase returns true on a 2xx response', () => {
  withGasGlobals(
    { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_ANON_KEY: 'key' },
    [{ code: 201, body: '' }],
    (calls) => {
      const ok = writeToSupabase({ dateString: '7/2 03:15', avgSys: 120, avgDia: 70, avgPul: 70 });
      assert.strictEqual(ok, true);
      assert.strictEqual(calls.length, 1);
    }
  );
});

runTest('sendOpsAlert pushes a fixed-category message without leaking raw error details', () => {
  withGasGlobals({ ALERT_LINE_USER_ID: 'U_SON' }, [{ code: 200 }], (calls) => {
    sendOpsAlert('Supabase 寫入失敗（回應非 2xx），紀錄時間：7/2 03:15');
    assert.strictEqual(calls.length, 1);
    const payload = JSON.parse(calls[0].options.payload);
    assert.match(payload.messages[0].text, /⚠️ 系統告警/);
    assert.match(payload.messages[0].text, /Supabase 寫入失敗/);
  });
});

runTest('sendOpsAlert silently skips when ALERT_LINE_USER_ID is unset', () => {
  withGasGlobals({}, [], (calls) => {
    assert.doesNotThrow(() => sendOpsAlert('anything'));
    assert.strictEqual(calls.length, 0);
  });
});

// doPost — LINE's webhook "Verify" ping sends `events: []`, and older code
// crashed on that (events[0] is undefined, next line throws on undefined.type)

runTest('doPost does not throw when LINE sends an empty events array (verify ping)', () => {
  global.ContentService = {
    createTextOutput: (text) => ({ text, getContent: () => text })
  };
  try {
    const result = doPost({ postData: { contents: JSON.stringify({ events: [] }) } });
    assert.strictEqual(result.getContent(), 'Success');
  } finally {
    delete global.ContentService;
  }
});

runTest('doPost does not throw when events key is missing entirely', () => {
  global.ContentService = {
    createTextOutput: (text) => ({ text, getContent: () => text })
  };
  try {
    const result = doPost({ postData: { contents: JSON.stringify({}) } });
    assert.strictEqual(result.getContent(), 'Success');
  } finally {
    delete global.ContentService;
  }
});

// doPost — top-level try/catch (temporary debug channel: reply the raw
// exception via LINE, since this project has no linked GCP project yet so
// Cloud Logging shows nothing). This must never let an exception escape
// doPost uncaught, since GAS then shows "Failed" with zero information.

runTest('doPost catches an internal exception and replies it back via LINE (temp debug)', () => {
  global.ContentService = { createTextOutput: (text) => ({ text, getContent: () => text }) };
  try {
    withGasGlobals({}, [{ code: 200 }], (calls) => {
      // A valid BP-format message reaches SpreadsheetApp.openById(), which
      // doesn't exist in this Node test environment — a real, unmocked
      // exception, not a hand-crafted one.
      const payload = {
        postData: {
          contents: JSON.stringify({
            events: [{
              type: 'message',
              replyToken: 'RT123',
              message: { type: 'text', text: '128/65/75 | 123/63/73' },
              source: { type: 'user', userId: 'U_TEST' }
            }]
          })
        }
      };

      const result = doPost(payload);
      assert.strictEqual(result.getContent(), 'Success');
      assert.strictEqual(calls.length, 1);
      const body = JSON.parse(calls[0].options.payload);
      assert.strictEqual(body.replyToken, 'RT123');
      assert.match(body.messages[0].text, /🐛 DEBUG/);
      assert.match(body.messages[0].text, /SpreadsheetApp/);
    });
  } finally {
    delete global.ContentService;
  }
});

runTest('doPost does not throw even if it cannot recover a replyToken to report the error', () => {
  global.ContentService = { createTextOutput: (text) => ({ text, getContent: () => text }) };
  try {
    withGasGlobals({}, [], (calls) => {
      // Malformed postData.contents (not valid JSON at all) — the nested
      // try/catch for extracting replyToken also fails; doPost must still
      // return normally instead of throwing all the way up to GAS.
      assert.doesNotThrow(() => {
        const result = doPost({ postData: { contents: 'not json' } });
        assert.strictEqual(result.getContent(), 'Success');
      });
      assert.strictEqual(calls.length, 0);
    });
  } finally {
    delete global.ContentService;
  }
});
