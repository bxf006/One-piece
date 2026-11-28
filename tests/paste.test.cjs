const { test } = require('node:test');
const assert = require('node:assert/strict');
const S = require('../src/sheet-reader.js');
const C = require('../src/core.js');
const data = (text, mode = 'auto') => C.readSheet(S.readPasted(text), '粘贴数据', S, mode);

test('Excel clipboard loads Chinese headers and actual time spacing into all four ranges', () => {
  const d = data('时间\t压力\t温度\t电流\t转速\r\n2026-01-01 08:00:10\t4.2\t26.5\t420\t7200\r\n2026-01-01 08:00:00\t4.5\t27.1\t450\t7350\r\n');
  assert.equal(d.kind, 'time');
  assert.equal(d.rows.length, 2);
  assert.equal(d.rows[1].x - d.rows[0].x, -10000);
  assert.deepEqual(d.rows[0].values, [4.2, 26.5, 420, 7200]);
  assert.deepEqual(d.series.map(s => s.range.id), ['small', 'thermal', 'medium', 'large']);
});

test('clipboard preserves blank middle and trailing cells, zero, duplicate columns and error gaps', () => {
  const d = data('时间\t压力\t压力\t温度\n0\t0\t\t\n1\t\t2.5\t#DIV/0!\n2\t3\t4\t20\n');
  assert.deepEqual(d.rows.map(r => r.values), [[0, null, null], [null, 2.5, null], [3, 4, 20]]);
  assert.equal(d.series[1].occurrence, 2);
  assert.equal(d.rows[0].x, 0);
});

test('quoted multiline cells, escaped quotes and punctuation stay within their columns', () => {
  const wb = S.readPasted('\uFEFF时间\t"压力,;\t名称\n第二行"\t"温度""传感器"\r\n08:00:00\t2.5\t20\r\n');
  const d = C.readSheet(wb, '粘贴数据', S);
  assert.equal(d.series[0].name, '压力,;\t名称\n第二行');
  assert.equal(d.series[1].name, '温度"传感器');
  assert.deepEqual(d.rows[0].values, [2.5, 20]);
  const singleColumn = '时间,温度\n0,20';
  assert.throws(() => S.readPasted(singleColumn), /至少两列/);
});

test('clipboard supports explicit numeric, Excel and Unix time modes and reusable configs', () => {
  assert.equal(data('时间\t值\n46023.5\t4', 'excel').rows[0].x, Date.UTC(2026, 0, 1, 12));
  assert.equal(data('时间\t值\n1767225600\t4', 'unix-s').rows[0].x, 1767225600000);
  const d = data('时间\t值\n1\t4');
  const config = C.validateConfig({ schema: 'curve-viewer', version: 1, name: '粘贴配置', selectedSeries: [{ name: '值', occurrence: 1 }], timeMode: 'auto' });
  assert.equal(d.series[0].key, C.keyOf(config.selectedSeries[0].name, config.selectedSeries[0].occurrence));
});

test('invalid clipboard data fails before replacing a chart', () => {
  for (const text of ['', '\t\r\n', '时间\t温度', '时间\n0\n1']) assert.throws(() => S.readPasted(text));
  assert.throws(() => S.readPasted('时间\t温度\n0\t20\t30'), /第 2 行/);
  assert.throws(() => S.readPasted('时间\t"温度\n0\t20'), /引号未闭合/);
  assert.throws(() => data('时间\t温度\n0\tbad'), /有效数值/);
  assert.throws(() => data('时间\t温度\n####\t20', 'date'), /没有可绘制/);
});

test('clipboard enforces row, column and text size limits', () => {
  assert.throws(() => S.readPasted(Array(513).fill('列').join('\t') + '\n0\t1'), /512 列/);
  assert.throws(() => S.readPasted('时间\t值\n' + '0\t1\n'.repeat(250001)), /250,000 行/);
  assert.throws(() => S.readPasted('x'.repeat(50 * 1024 * 1024 + 1)), /50 MB/);
  assert.doesNotThrow(() => S.readPasted(Array(512).fill('列').join('\t') + '\n' + Array(512).fill('1').join('\t')));
});
