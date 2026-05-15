const assert = require('assert');
const {
  buildReplyBlock,
  detectPeriod,
  extractBpLine,
  getBpStatus,
  parseBpEntries,
  parseHeaderLine
} = require('./Code.js');

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

  assert.deepStrictEqual(parseBpEntries(sample), [
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
  ]);
});

runTest('parseBpEntries falls back to today for simple single-line input', () => {
  const entries = parseBpEntries('Pagi 128/65/75 | 123/63/73');
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

runTest('getBpStatus keeps low diastolic classification', () => {
  assert.deepStrictEqual(getBpStatus(117, 59), {
    zh: '⚠️ 舒張壓偏低',
    id: '⚠️ Diastolik Rendah'
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
