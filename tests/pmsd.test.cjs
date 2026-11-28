const { test } = require('node:test');
const assert = require('node:assert/strict');
const S = require('../src/sheet-reader.js');
const C = require('../src/core.js');
const { makePMSD, crc32, repairHeader, START_MS, FILENAME } = require('./pmsd-fixture.cjs');
const load = (bytes, name = FILENAME) => { const workbook = S.readDataFile(bytes, name); return { workbook, data: C.readSheet(workbook, '数据', S) }; };

test('PMSD v2 reads engineering values, units and millisecond times matching the original filename zone', () => {
  const { workbook, data } = load(makePMSD());
  assert.deepEqual(data.series.map(s => s.name), ['压力[bar]', '温度[℃]']);
  assert.deepEqual(data.rows.map(r => r.values), [[1.4, -18.9], [1.6, 43]]);
  assert.equal(C.formatTime(data.rows[0].x), '2026-11-25 18:00:40.964');
  assert.equal(C.formatTime(data.rows[1].x), '2026-11-25 18:00:41.601');
  assert.equal(data.rows[1].x - data.rows[0].x, 637);
  assert.equal(workbook.ImportInfo.timeZone, 'UTC+08:00'); assert.equal(workbook.ImportInfo.timeZoneSource, 'filename');
  assert.equal(workbook.ImportInfo.records, 2);
  const enclosed = Buffer.concat([Buffer.alloc(17), makePMSD(), Buffer.alloc(3)]);
  assert.deepEqual(load(enclosed.subarray(17, -3)).data.rows, data.rows);
});

test('PMSD preserves source order, repeated timestamps and duplicate names; validity bits create gaps', () => {
  const fields = [{ name: '告警', unit: '1' }, { name: '告警', unit: '1' }, { name: '压力', unit: 'bar' }];
  const rows = [
    { timeMs: START_MS + 2000, values: [0, 1, 1.4] }, { timeMs: START_MS, values: [77, 2, 2.5], valid: [false, true, true] },
    { timeMs: START_MS, values: [99, 99, 99], valid: [false, false, false] }, { timeMs: START_MS + 1000, values: [3, NaN, Infinity] }
  ];
  const { data, workbook } = load(makePMSD({ fields, rows }));
  assert.deepEqual(data.series.map(s => s.label), ['告警[1]', '告警[1] (2)', '压力[bar]']);
  assert.deepEqual(data.rows.map(r => r.values), [[0, 1, 1.4], [null, 2, 2.5], [null, null, null], [3, null, null]]);
  assert.deepEqual(data.rows.map(r => r.sourceRow), [2, 3, 4, 5]);
  assert.deepEqual(data.rows.map(r => r.x - data.rows[1].x), [2000, 0, 0, 1000]);
  assert.equal(C.nearest(data.rows, data.rows[1].x), 1); assert.equal(workbook.ImportInfo.emptyRecords, 1); assert.equal(data.missing, 6);
});

test('PMSD handles negative fractional-hour zones and gives explicit local fallback after renaming', () => {
  const input = makePMSD();
  const western = load(input, 'sample_20261125_043040_865_sid_1.pmsd');
  assert.equal(C.formatTime(western.data.rows[0].x), '2026-11-25 04:30:40.964');
  assert.equal(western.workbook.ImportInfo.timeZone, 'UTC−05:30');
  for (const name of ['renamed.PMSD', 'sample_20261325_180040_865_sid_1.pmsd', 'sample_20250101_180040_865_sid_1.pmsd']) {
    const { workbook, data } = load(input, name), date = new Date(START_MS + 99);
    assert.equal(workbook.ImportInfo.timeZoneSource, 'local');
    assert.equal(data.rows[0].x, Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds()));
  }
});

test('PMSD rejects unsupported formats, invalid headers and broken metadata instead of guessing', () => {
  assert.throws(() => load(Buffer.from('时间,值\n0,2')), /PMSAMPLE/);
  assert.throws(() => load(makePMSD({ version: 1 })), /版本/);
  assert.throws(() => load(makePMSD({ flags: 0 })), /存储方式/);
  for (const [change, message] of [
    [b => b.writeUInt32LE(0, 16), /511/], [b => b.writeUInt32LE(512, 16), /512/],
    [b => b.writeUInt32LE(b.length + 1, 12), /文件头长度/], [b => { b[60] ^= 1; }, /字段信息校验/],
    [b => { b.writeUInt16LE(4, 52); repairHeader(b); }, /字段索引/],
    [b => { b.writeUInt16LE(65535, 58); repairHeader(b); }, /文本长度/],
    [b => { b[60] = 0xff; repairHeader(b); }, /UTF-8/]
  ]) { const input = makePMSD(); change(input); assert.throws(() => load(input), message); }
  assert.throws(() => load(makePMSD({ rows: [] })), /至少一行/);
  assert.throws(() => load(new Uint8Array(50 * 1024 * 1024 + 1)), /50 MB/);
});

test('PMSD rejects damaged or truncated samples and trailing bytes atomically', () => {
  for (const [change, message] of [
    [(b, start) => { b[start] = 0; }, /类型不支持/],
    [(b, start) => b.writeUInt32LE(1, start + 4), /记录长度/],
    [(b, start) => b.writeUInt32LE(2, start + 24), /标记长度/],
    [(b, start) => { b[start + 30] ^= 1; }, /记录校验/]
  ]) { const input = makePMSD(), start = input.readUInt32LE(12); change(input, start); assert.throws(() => load(input), message); }
  const input = makePMSD(), header = input.readUInt32LE(12);
  for (const end of [header + 1, header + 5, header + 30, input.length - 1]) assert.throws(() => load(input.subarray(0, end)), /不完整|截断/);
  assert.throws(() => load(Buffer.concat([input, Buffer.from([1])])), /不完整/);
});

test('PMSD END1 trailers validate counts and sample CRC without adding a data row', () => {
  const plain = load(makePMSD()), completed = load(makePMSD({ endRecord: true }));
  assert.deepEqual(completed.data, plain.data);
  assert.equal(plain.workbook.ImportInfo.endRecord, null);
  assert.equal(completed.workbook.ImportInfo.records, 2);
  assert.deepEqual(completed.workbook.ImportInfo.endRecord, { records: 2, time: '2026-11-25 18:00:42.101' });
  const rows = [{ timeMs: START_MS + 2000, values: [1,20] }, { timeMs: START_MS, values: [2,30] }];
  assert.deepEqual(load(makePMSD({ rows, endRecord: true })).data.rows.map(r => r.x), load(makePMSD({ rows })).data.rows.map(r => r.x));
  const input = makePMSD({ endRecord: true }), enclosed = Buffer.concat([Buffer.alloc(7), input, Buffer.alloc(11)]);
  assert.deepEqual(load(enclosed.subarray(7,-11)).data, completed.data);
});

test('PMSD rejects damaged END1 counts and invalid aggregate CRC while sample CRCs remain valid', () => {
  for (const count of [0n, 1n, 3n, (1n << 32n) + 2n, (1n << 63n) + 2n]) {
    const bytes = makePMSD({ endRecord: true }); bytes.writeBigUInt64LE(count, bytes.length - 20);
    assert.throws(() => load(bytes), /结束块.*计数/);
  }
  const checksum = makePMSD({ endRecord: true }); checksum[checksum.length-1] ^= 1;
  assert.throws(() => load(checksum), /结束块.*采样区校验/);
  const wrongCoverage = makePMSD({ endRecord: true }), start = wrongCoverage.readUInt32LE(12), end = wrongCoverage.length - 28;
  wrongCoverage.writeUInt32LE(crc32(wrongCoverage.subarray(start+8,end-4)),end+24);
  assert.throws(() => load(wrongCoverage), /结束块.*采样区校验/);
  assert.deepEqual(load(wrongCoverage.subarray(0,-28)).data, load(makePMSD()).data);
});

test('PMSD requires a complete terminal END1 and continues rejecting unknown record types', () => {
  const bytes = makePMSD({ endRecord: true }), start = bytes.length - 28;
  for (let remaining=1;remaining<28;remaining++) assert.throws(() => load(bytes.subarray(0,start+remaining)), /不完整|截断/);
  for (const length of [0, 8, 27, 29, 0xffffffff]) {
    const bad = Buffer.from(bytes); bad.writeUInt32LE(length,start+4); assert.throws(() => load(bad), /结束块长度/);
  }
  for (const extra of [Buffer.from([0]), bytes.subarray(start), bytes.subarray(bytes.readUInt32LE(12),start)]) assert.throws(() => load(Buffer.concat([bytes,extra])), /结束块后.*多余数据/);
  const unknown = Buffer.from(bytes); unknown.write('END2',start,'ascii'); assert.throws(() => load(unknown), /类型不支持/);
  const damaged = Buffer.from(bytes); damaged[damaged.readUInt32LE(12)+30] ^= 1; assert.throws(() => load(damaged), /第 1 条记录校验/);
});

test('PMSD completed captures also parse in browser globals without external libraries', () => {
  const vm = require('node:vm'), fs = require('node:fs'), path = require('node:path');
  const context = vm.createContext({ TextDecoder, TextEncoder, Uint8Array, DataView, ArrayBuffer });
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../src/sheet-reader.js'),'utf8'),context);
  const workbook = context.SheetReader.readDataFile(new Uint8Array(makePMSD({ endRecord: true })),FILENAME);
  assert.equal(workbook.ImportInfo.records,2); assert.equal(workbook.ImportInfo.endRecord.records,2);
  assert.deepEqual(C.readSheet(workbook,'数据',context.SheetReader).rows.map(row => Array.from(row.values)),[[1.4,-18.9],[1.6,43]]);
});
