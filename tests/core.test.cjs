const { test } = require('node:test');
const assert = require('node:assert/strict');
const C = require('../src/core.js');
const S = require('../src/sheet-reader.js');
function book(rows) { const wb = S.utils.book_new(); S.utils.book_append_sheet(wb, S.utils.aoa_to_sheet(rows), '数据'); return wb; }
test('device headers extract the final path name and bracket value without changing plain headers', () => {
  assert.equal(C.normalizeHeader('HPODU/设备/告警1 (ID 34000) [1]'), '告警1[1]');
  assert.equal(C.normalizeHeader(' HPODU/设备/压缩机频率  (ID 34016) [ Hz ]  '), '压缩机频率[Hz]');
  assert.equal(C.normalizeHeader('HPODU/设备/告警1 (ID 34000) [01]'), '告警1[01]');
  for (const name of ['时间', '告警1[1]', '设备/压力', '温度 (℃)', '告警1 (ID 34000) [1]', '设备/告警1 (ID 34000)', '设备/告警1 (ID 34000) []', '设备/ (ID 34000) [1]', '设备/告警1 (ID 34000 [1]']) {
    assert.equal(C.normalizeHeader(name), name);
  }
});
test('renamed headers keep duplicate columns, original names and data values intact', () => {
  const first = 'HPODU/设备/告警1 (ID 34000) [1]', second = 'HPODU/另一设备/告警1 (ID 35000) [1]';
  const wb = book([['HPODU/设备/时间 (ID 1) [s]', first, second, first], [0, 0, 1, 2], [1, 3, 4, 5]]);
  const d = C.readSheet(wb, '数据', S);
  assert.equal(d.timeName, '时间[s]');
  assert.deepEqual(d.series.map(s => s.label), ['告警1[1]', '告警1[1] (2)', '告警1[1] (3)']);
  assert.equal(new Set(d.series.map(s => s.key)).size, 3);
  assert.deepEqual(d.series.map(s => [s.sourceName, s.sourceOccurrence]), [[first, 1], [second, 1], [first, 2]]);
  assert.deepEqual(d.rows.map(r => r.values), [[0, 1, 2], [3, 4, 5]]);
  assert.equal(wb.Sheets['数据'].B1.v, first);
});
test('range boundaries and overflow never clip a sequence', () => {
  assert.equal(C.rangeFor(0, 10).id, 'small'); assert.equal(C.rangeFor(-40, 100).id, 'thermal');
  assert.equal(C.rangeFor(0, 1000).id, 'medium'); assert.equal(C.rangeFor(0, 10000).id, 'large');
  assert.equal(C.rangeFor(-41, 100), null); assert.equal(C.rangeFor(0, 10001), null);
  const a = C.axesFor([{ count: 3, min: -90, max: 24000, range: null }])[0]; assert.ok(a.min <= -90 && a.max >= 24000);
});
test('manual axis bounds preserve grouping and other axes, including overflow', () => {
  const series = [
    { count: 2, min: 1, max: 2, range: C.RANGES[0] },
    { count: 2, min: -20, max: 50, range: C.RANGES[1] },
    { count: 2, min: -90, max: 24000, range: null }
  ];
  const automatic = C.axesFor(series);
  const limits = C.validateAxisLimits({ small: { min: -0.5, max: 1.75 }, auto: { min: 100, max: 200 } });
  const axes = C.axesFor(series, limits);
  assert.deepEqual(axes.map(a => a.id), ['small', 'thermal', 'auto']);
  assert.equal(axes[0].min, -0.5); assert.equal(axes[0].max, 1.75); assert.equal(axes[0].manual, true);
  assert.equal(axes[0].label, '−0.5 ~ 1.75'.replace('−', '-'));
  assert.deepEqual(axes[1], automatic[1]);
  assert.equal(axes[2].min, 100); assert.equal(axes[2].max, 200);
  assert.deepEqual(C.axesFor(series), automatic); assert.equal(C.RANGES[0].min, 0);
});

test('automatic bounds pad grouped data extrema and round outward while keeping default bands', () => {
  const d = C.readSheet(book([
    ['时间', '压力A', '压力B', '温度', '电压', '转速', '超范围', '空列'],
    [0, 1.4, 2, -18.9, 226, 2100, -80, null],
    [1, 1.6, 3, 43, 320, 3500, -60, null]
  ]), '数据', S);
  const axes = C.axesFor(d.series);
  assert.deepEqual(axes.map(a => [a.id, a.min, a.max, a.band]), [
    ['small', 1, 4, [0.8, 1]], ['thermal', -22, 47, [0.25, 0.8]],
    ['medium', 221, 325, [0, 0.25]], ['large', 2030, 3570, [0, 1]], ['auto', -81, -59, [0, 1]]
  ]);
  assert.deepEqual(C.axesFor([d.series[0]]).map(a => [a.min, a.max]), [[1, 2]]);
  for (const value of [0, 1, 43, -80, 320]) {
    const axis = C.axesFor([{ count: 1, min: value, max: value, range: C.rangeFor(value, value) }])[0];
    assert.ok(axis.min < value && axis.max > value && Number.isFinite(axis.max - axis.min));
  }
});

test('automatic padding uses five percent of the data span, including negative and constant values', () => {
  for (const [min, max, expected] of [
    [-20, 30, [-23, 33]], [20, 80, [17, 83]], [-20, -10, [-21, -9]],
    [0, 10, [-1, 11]], [1.4, 1.6, [1, 2]], [0.01, 0.02, [0, 1]],
    [0, 0, [-1, 1]], [43, 43, [40, 46]], [-43, -43, [-46, -40]]
  ]) {
    const series = [{ count: 2, min, max, range: C.rangeFor(min, max) }];
    const axis = C.axesFor(series)[0];
    assert.deepEqual([axis.min, axis.max], expected, `range ${min} to ${max}`);
    assert.ok(axis.min < min && axis.max > max);
    const manual = C.axesFor(series, { [axis.id]: { min: -0.75, max: 2.25 } })[0];
    assert.deepEqual([manual.min, manual.max], [-0.75, 2.25]);
  }
});

test('custom region percentages round trip, preserve absent defaults and allow overlapping groups', () => {
  const axisRegions = { small: { start: 10.5, end: 40.5 }, thermal: { start: 20, end: 70 }, auto: { start: 30, end: 100 } };
  const config = { schema: 'curve-viewer', version: 1, name: '区域', timeMode: 'auto', selectedSeries: [], axisRegions };
  assert.deepEqual(C.validateConfig(JSON.parse(JSON.stringify(config))), config);
  assert.equal(C.validateConfig({ ...config, axisRegions: undefined }).axisRegions, undefined);
  assert.equal(C.validateConfig({ ...config, axisRegions: {} }).axisRegions, undefined);
  const d = C.readSheet(book([['时间', '压力', '温度', '电压', '超范围'], [0, 1, -20, 226, -80], [1, 3, 40, 320, -60]]), '数据', S);
  const defaults = C.axesFor(d.series), custom = C.axesFor(d.series, {}, {}, axisRegions);
  assert.deepEqual(custom.map(a => a.band), [[0.105, 0.405], [0.2, 0.7], [0, 0.25], [0.3, 1]]);
  assert.deepEqual(custom.map(a => [a.min, a.max]), defaults.map(a => [a.min, a.max]));
  assert.equal(custom[0].region, '10.5%～40.5%');
  for (const invalid of [null, [], { bogus: { start: 0, end: 100 } }, { small: null }, { small: { start: '', end: 100 } },
    { small: { start: 0, end: Infinity } }, { small: { start: -1, end: 50 } }, { small: { start: 0, end: 101 } },
    { small: { start: 50, end: 50 } }, { small: { start: 70, end: 20 } }]) {
    assert.throws(() => C.validateConfig({ ...config, axisRegions: invalid }));
  }
});

test('manual curve assignment changes grouping and extrema while axis limits still take precedence', () => {
  const d = C.readSheet(book([['时间', '压力', '温度', '电压'], [0, 1, -20, 226], [1, 3, 40, 320]]), '数据', S);
  const assignments = { [d.series[0].key]: 'medium', [d.series[1].key]: 'auto' };
  assert.equal(C.axisIdFor(d.series[0], assignments), 'medium');
  assert.deepEqual(C.axesFor(d.series, {}, assignments).map(a => [a.id, a.min, a.max, a.band]), [
    ['medium', -15, 336, [0, 0.25]], ['auto', -23, 43, [0, 1]]
  ]);
  const manual = C.axesFor(d.series, { medium: { min: 100, max: 250 } }, assignments);
  assert.deepEqual([manual[0].min, manual[0].max, manual[0].manual, manual[0].band], [100, 250, true, [0, 0.25]]);
  assert.deepEqual(C.axesFor([d.series[2]], {}, assignments).map(a => [a.min, a.max]), [[221, 325]]);
  assert.equal(d.series[0].range.id, 'small');
});

test('curve coordinate assignments round trip with duplicate names and reject malformed config entries', () => {
  const base = { schema: 'curve-viewer', version: 1, name: '分区', timeMode: 'auto', selectedSeries: [] };
  const one = { name: '温度', occurrence: 1, axisId: 'medium' };
  const config = { ...base, seriesAxes: [one, { name: '温度', occurrence: 2, axisId: 'auto' }] };
  assert.deepEqual(C.validateConfig(JSON.parse(JSON.stringify(config))), config);
  assert.deepEqual(C.validateConfig(base), base);
  for (const seriesAxes of [null, {}, [null], [{ ...one, axisId: 'custom' }], [{ ...one, occurrence: 0 }], [{ ...one, name: '' }], [one, one], Array(513).fill(one)]) {
    assert.throws(() => C.validateConfig({ ...base, seriesAxes }));
  }
});
test('axis settings round trip in v1 configs and reject invalid bounds atomically', () => {
  const config = { schema: 'curve-viewer', version: 1, name: '坐标', timeMode: 'auto', selectedSeries: [], axisLimits: { thermal: { min: -25.25, max: 45.5 }, small: { min: 0, max: 1e-13 } } };
  assert.deepEqual(C.validateConfig(JSON.parse(JSON.stringify(config))), config);
  assert.equal(C.validateConfig({ ...config, axisLimits: undefined }).axisLimits, undefined);
  for (const axisLimits of [null, [], { bogus: { min: 0, max: 1 } }, { small: { min: 0, max: 0 } }, { small: { min: 2, max: 1 } }, { small: { min: '', max: 1 } }, { small: { min: 0, max: Infinity } }, { small: { min: -1e308, max: 1e308 } }, { small: { min: null, max: 1 } }]) {
    assert.throws(() => C.validateConfig({ ...config, axisLimits }));
  }
});
test('real time spacing preserves source order, invalid rows, nulls, and duplicate headers', () => {
  const d = C.readSheet(book([['时间', '温度', '温度', '空列'], ['2026-01-01 08:00:10', 20, 8, null], ['2026-01-01 08:00:00', -10, 2, null], [null, 99, 3, null], ['无效', 30, 4, null], ['2026-01-01 08:01:00', '', '坏值', null]]), '数据', S);
  assert.equal(d.kind, 'time'); assert.equal(d.rows.length, 3); assert.equal(d.skipped, 2);
  assert.equal(d.rows[1].x - d.rows[0].x, -10000); assert.equal(d.rows[2].x - d.rows[0].x, 50000);
  assert.deepEqual(d.rows.map(r => r.sourceRow), [2, 3, 6]);
  assert.deepEqual(d.domain, [d.rows[1].x, d.rows[2].x]);
  assert.deepEqual(d.rows[0].values, [20, 8, null]); assert.equal(d.series[1].occurrence, 2); assert.notEqual(d.series[0].key, d.series[1].key);
  assert.equal(d.series[0].range.id, 'thermal'); assert.equal(d.series[1].range.id, 'small'); assert.equal(d.series[2].count, 0);
});
test('Excel serials respect 1900 and 1904 systems', () => {
  const wb = book([['时间', '数值'], [46023.5, 4], [46024, 6]]); wb.Sheets['数据'].A2.z = 'yyyy-mm-dd hh:mm:ss';
  const a = C.readSheet(wb, '数据', S); assert.equal(a.detected, 'excel'); assert.equal(new Date(a.rows[0].x).toISOString(), '2026-01-01T12:00:00.000Z');
  wb.Workbook = { WBProps: { date1904: true } }; const b = C.readSheet(wb, '数据', S); assert.equal(b.rows[0].x - a.rows[0].x, 1462 * 86400000);
});
test('numeric zero time is retained and unix units are explicit', () => {
  const d = C.readSheet(book([['时间', '值'], [2, 2], [0, 4], [1, 3]]), '数据', S);
  assert.deepEqual(d.rows.map(r => r.x), [2, 0, 1]); assert.deepEqual(d.domain, [0, 2]); assert.equal(d.kind, 'value');
  const unix = C.readSheet(book([['时间', '值'], [1767225600, 1]]), '数据', S, 'unix-s'); assert.equal(unix.rows[0].x, 1767225600000);
});
test('numeric parsing rejects booleans, errors, units and missing values', () => {
  for (const value of [null, undefined, '', ' ', true, '12℃', 'NaN', Infinity, '0x12']) assert.equal(C.numeric(value), null);
  assert.equal(C.numeric(' -1.5e2 '), -150); assert.equal(C.numeric('0'), 0);
});
test('date validation rejects impossible dates and preserves milliseconds', () => {
  assert.equal(C.parseTextTime('2026-02-30'), null); assert.equal(C.parseTextTime('24:61:00'), null);
  assert.equal(C.parseTextTime('08:05:01.125'), 29101125);
  assert.equal(C.formatTime(C.parseTextTime('2026/01/01 08:00:00.125')), '2026-01-01 08:00:00.125');
});
test('configuration round trip, empty selection and dangerous colors', () => {
  const config = { schema: 'curve-viewer', version: 1, name: '测试', selectedSeries: [{ name: '<温度>', occurrence: 1, color: '#5465e8' }], timeMode: 'auto' };
  assert.deepEqual(C.validateConfig(JSON.parse(JSON.stringify(config))), config);
  assert.deepEqual(C.validateConfig({ ...config, selectedSeries: [] }).selectedSeries, []);
  assert.throws(() => C.validateConfig({ ...config, version: 99 }));
  assert.throws(() => C.validateConfig({ ...config, selectedSeries: [config.selectedSeries[0], config.selectedSeries[0]] }));
  assert.equal(C.validateConfig({ ...config, selectedSeries: [{ name: '温度', occurrence: 1, color: 'red;position:fixed' }] }).selectedSeries[0].color, undefined);
});
test('nearest hover point is chosen by actual time distance', () => {
  const rows = [{ x: 0 }, { x: 10 }, { x: 1000 }]; assert.equal(C.nearest(rows, 500), 1); assert.equal(C.nearest(rows, 800), 2);
  const unsorted = [{ x: 20 }, { x: 0 }, { x: 10 }, { x: 10 }];
  assert.equal(C.nearest(unsorted, 9), 2); assert.equal(C.nearest(unsorted, 1), 1);
  assert.equal(C.nearest(unsorted, 10), 2); assert.equal(C.nearest(unsorted, 15), 0);
});
test('CSV supports Chinese headers, quoted commas, line breaks and escaping', () => {
  assert.deepEqual(S.parseCSV('时间,\"名称,一\",值\r\n08:00:00,\"两\"\"个\",2\r\n'), [['时间', '名称,一', '值'], ['08:00:00', '两"个', '2']]);
  assert.deepEqual(S.parseCSV('时间;值\n0;5'), [['时间', '值'], ['0', '5']]);
  assert.throws(() => S.parseCSV('a,b\n\"oops'));
});
test('blank, all-text and error workbooks fail clearly', () => {
  assert.throws(() => C.readSheet(book([['时间', '值']]), '数据', S));
  assert.throws(() => C.readSheet(book([['时间', '值'], [0, 'bad']]), '数据', S));
  const wb = book([['时间', '值'], [0, 7]]); wb.Sheets['数据'].B2.t = 'e'; assert.throws(() => C.readSheet(wb, '数据', S));
});
