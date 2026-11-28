const { test } = require('node:test');
const assert = require('node:assert/strict');
const S = require('../src/sheet-reader.js');
const C = require('../src/core.js');
const bytes = text => new TextEncoder().encode(text);
const read = (text, name) => C.readSheet(S.readDataFile(bytes(text), name), '数据', S);

test('all upload formats keep descending and repeated source rows while domain covers every time', () => {
  const rows = [['时间', '值'], [8, 80], [20, 200], [0, 0], [8, 81], [5, 50]];
  for (const [text, name] of [
    [rows.map(row => row.join(',')).join('\n'), 'order.csv'],
    [rows.map(row => row.join('\t')).join('\n'), 'order.txt'],
    [JSON.stringify(rows), 'order.json'],
    [JSON.stringify({ columns: rows[0], data: rows.slice(1) }), 'order.json'],
    [JSON.stringify(rows.slice(1).map(([时间, 值]) => ({ 时间, 值 }))), 'order.json']
  ]) {
    const d = read(text, name);
    assert.deepEqual(d.rows.map(row => [row.x, row.values[0]]), rows.slice(1));
    assert.deepEqual(d.rows.map(row => row.sourceRow), [2, 3, 4, 5, 6]);
    assert.deepEqual(d.domain, [0, 20]);
  }
});

test('CSV, TXT and all supported JSON shapes use the same shortened device headers', () => {
  const name = 'HPODU/设备/告警1 (ID 34000) [1]';
  const rows = [['时间', name], [0, 2], [1, 3]];
  for (const [text, file] of [
    [rows.map(row => row.join(',')).join('\n'), 'headers.csv'],
    [rows.map(row => row.join('\t')).join('\n'), 'headers.txt'],
    [JSON.stringify(rows), 'headers.json'],
    [JSON.stringify({ columns: rows[0], data: rows.slice(1) }), 'headers.json'],
    [JSON.stringify([{ 时间: 0, [name]: 2 }, { 时间: 1, [name]: 3 }]), 'headers.json']
  ]) {
    const d = read(text, file);
    assert.equal(d.series[0].name, '告警1[1]');
    assert.deepEqual(d.rows.map(r => r.values), [[2], [3]]);
    assert.deepEqual(d.rows.map(r => r.x), [0, 1]);
  }
});

test('CSV/TXT decode common delimiters, quotes and actual time intervals', () => {
  for (const [separator, name] of [[',', 'sample.CSV'], ['\t', 'sample.txt'], [';', 'sample.TXT']]) {
    const table = [['时间', '压力', '温度'], ['2026-11-25 18:00:40.964', 1.4, -18.9], ['2026-11-25 18:00:41.601', 0, '']].map(row => row.join(separator)).join('\r\n');
    const d = read(table, name);
    assert.equal(d.rows[1].x - d.rows[0].x, 637);
    assert.deepEqual(d.rows.map(row => row.values), [[1.4, -18.9], [0, null]]);
  }
  const d = read('时间,"名称;;;,一"\n0,2\n1,3', 'quoted.csv');
  assert.equal(d.series[0].name, '名称;;;,一');
  assert.equal(read('sep=;\r\n时间;值\r\n0;2', 'sample.csv').rows[0].values[0], 2);
});

test('text imports support UTF-8 BOM, UTF-16 LE/BE and GBK', () => {
  const text = '时间,值\n0,2';
  const le = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]);
  const be = Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from(text, 'utf16le').swap16()]);
  const gbk = Buffer.from('cab1bce42cd6b50a302c32', 'hex');
  for (const data of [bytes('\uFEFF' + text), le, be, gbk]) {
    const d = C.readSheet(S.readDataFile(data, 'sample.txt'), '数据', S);
    assert.equal(d.timeName, '时间'); assert.equal(d.series[0].name, '值'); assert.equal(d.rows[0].values[0], 2);
  }
});

test('JSON object rows find the time field, union columns and preserve missing values', () => {
  const d = read(JSON.stringify([{ '温度': -18.9, '时间': '2026-11-25 18:00:40.964' }, { '时间': '2026-11-25 18:00:41.601', '压力': 1.4 }]), 'data.json');
  assert.equal(d.timeName, '时间');
  assert.deepEqual(d.series.map(s => s.name), ['温度', '压力']);
  assert.deepEqual(d.rows.map(r => r.values), [[-18.9, null], [null, 1.4]]);
  assert.equal(d.rows[1].x - d.rows[0].x, 637);
});

test('JSON matrices and columns/data support duplicate headers and numeric zero times', () => {
  const rows = [['时间', '值', '值'], [0, 0, 2], [1, 3, null]];
  for (const input of [rows, { columns: rows[0], data: rows.slice(1) }]) {
    const d = read(JSON.stringify(input), 'data.json');
    assert.equal(d.rows[0].x, 0); assert.equal(d.series[1].occurrence, 2);
    assert.deepEqual(d.rows[1].values, [3, null]);
  }
  const d = read('[{"time":0,"__proto__":2,"constructor":3}]', 'names.json');
  assert.deepEqual(d.rows[0].values, [2, 3]);
});

test('display configuration JSON is distinguished from curve data', () => {
  const config = { schema: 'curve-viewer', version: 1, name: '配置', selectedSeries: [], timeMode: 'auto' };
  for (const input of [config, [config]]) assert.throws(() => S.readDataFile(bytes(JSON.stringify(input)), 'config.json'), error => error.code === 'DISPLAY_CONFIG');
});

test('unsupported extensions, renamed spreadsheets, invalid JSON and nested cells fail clearly', () => {
  for (const name of ['file.pdf', 'file.exe', 'file.docx']) assert.throws(() => read('时间,值\n0,2', name), /支持 Excel、CSV/);
  assert.throws(() => S.readDataFile(new Uint8Array([0x50, 0x4b, 3, 4]), 'renamed.csv'), /不能只修改扩展名/);
  assert.throws(() => read('{oops}', 'bad.json'), /JSON 格式无效/);
  for (const input of [[], {}, [2, 3], [{ time: 0, value: { nested: 2 } }], [['时间', '值'], [0, [2]]]]) assert.throws(() => read(JSON.stringify(input), 'bad.json'));
  assert.throws(() => read('', 'empty.txt'), /文件为空/);
  assert.throws(() => read('时间,值\n0,2,3', 'ragged.csv'), /列数多于列名/);
});

test('all data formats enforce the same row and column limits', () => {
  assert.throws(() => read(Array(513).fill('列').join(',') + '\n0,1', 'wide.csv'), /512 列/);
  assert.throws(() => read('时间\t值\n' + '0\t1\n'.repeat(250001), 'long.txt'), /250,000 行/);
  assert.throws(() => read(JSON.stringify({ columns: Array(513).fill('列'), data: [[0, 1]] }), 'wide.json'), /512 列/);
});

test('downloaded CSV round trips quoted headers, missing cells and data', () => {
  const workbook = S.utils.book_new();
  S.utils.book_append_sheet(workbook, S.utils.aoa_to_sheet([['时间', '值,"一"'], [0, 2], [1, null], [2, -1.5]]), '数据');
  const d = read(S.writeCSV(workbook), 'sample.csv');
  assert.equal(d.series[0].name, '值,"一"');
  assert.deepEqual(d.rows.map(r => r.values[0]), [2, null, -1.5]);
});
