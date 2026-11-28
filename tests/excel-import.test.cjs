const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const S = require('../src/sheet-reader.js');
const C = require('../src/core.js');
const { X, JSZip, excelBook, excelBytes } = require('./excel-fixture.cjs');

for (const extension of ['xlsx', 'xls', 'xlsm', 'xlsb', 'xltx', 'xltm', 'xlt', 'xml']) {
  test(`Excel .${extension} imports Chinese fields, sheets, dates and gaps without sorting`, async () => {
    const bytes = await excelBytes(extension), book = await S.readDataFile(bytes, `采样.${extension.toUpperCase()}`);
    assert.deepEqual(book.SheetNames, ['说明', '运行数据']);
    const data = C.readSheet(book, '运行数据', S);
    assert.deepEqual(data.series.map(s => s.name), ['温度[℃]', '压力']);
    assert.equal(data.detected, 'excel'); assert.equal(C.formatTime(data.rows[1].x), '2026-01-01 12:00:00');
    assert.deepEqual(data.rows.map(row => row.values), [[30, 3], [-20, 1.4], [40, null]]);
    assert.deepEqual(data.rows.map(row => row.x - data.rows[1].x), [43200000, 0, 43200000]);
    assert.equal(book.ImportInfo.format, 'Excel');
  });
}

test('legacy BIFF2/3/4/5 files including XLW retain numeric time and data', async () => {
  for (const type of ['biff2', 'biff3', 'biff4', 'biff5']) {
    const book = X.utils.book_new(); X.utils.book_append_sheet(book, X.utils.aoa_to_sheet([['Time', 'Value'], [2, 30], [0, -20], [1, 40]]), 'Data');
    const parsed = await S.readDataFile(X.write(book, { type: 'buffer', bookType: type }), type === 'biff4' ? 'old.xlw' : 'old.xls');
    const data = C.readSheet(parsed, parsed.SheetNames[0], S);
    assert.equal(data.detected, 'number'); assert.deepEqual(data.rows.map(r => [r.x, r.values[0]]), [[2, 30], [0, -20], [1, 40]]);
  }
});

test('Excel dates honor the 1904 calendar, while formulas use cached values and errors remain gaps', async () => {
  for (const extension of ['xlsx', 'xls', 'xlsb']) {
    const book = excelBook(); book.Workbook = { WBProps: { date1904: true } };
    const sheet = book.Sheets['运行数据']; sheet.B2 = { t: 'n', f: '10+5', v: 15 }; sheet.B3 = { t: 'e', v: 7 }; sheet.B4 = { t: 'n', f: '1+2' };
    const parsed = await S.readDataFile(await excelBytes(extension, book), `formulas.${extension}`);
    const data = C.readSheet(parsed, '运行数据', S);
    assert.equal(data.rows[1].x, Date.UTC(2026, 0, 1, 12) + 1462 * 86400000);
    assert.equal(data.rows[0].values[0], 15); assert.equal(data.rows[1].values[0], null);
    if (extension === 'xlsx') assert.equal(data.rows[2].values[0], null); // No cached result: do not calculate or invent zero.
    assert.equal(parsed.Sheets['运行数据'].B2.f, undefined); assert.equal(parsed.vbaraw, undefined);
  }
});

test('Excel input checks reject disguised text, arbitrary XML, invalid ZIPs and non-workbook archives', async () => {
  for (const [name, content] of [['fake.xls', '时间,值\n0,2'], ['fake.xml', '<document>hello</document>'], ['fake.xlsx', 'not a workbook']]) {
    await assert.rejects(() => S.readDataFile(Buffer.from(content), name), /不是可识别的 Excel/);
  }
  await assert.rejects(() => S.readDataFile(Buffer.from([0x50, 0x4b, 3, 4]), 'broken.xlsx'), /压缩结构/);
  const zip = new JSZip(); zip.file('readme.txt', 'not a workbook');
  await assert.rejects(() => S.readDataFile(Buffer.alloc(0), 'empty.xls'), /为空/);
  await assert.rejects(() => S.readDataFile(new Uint8Array(50 * 1024 * 1024 + 1), 'large.xlsx'), /50 MB/);
  const archive = await zip.generateAsync({ type: 'nodebuffer' });
  await assert.rejects(() => S.readDataFile(archive, 'archive.xlsx'), /没有 Excel 工作簿/);
});

test('Excel does not silently truncate oversized worksheet ranges and enforces archive expansion limits', async () => {
  const bytes = await excelBytes('xlsx');
  for (const range of ['A1:C250003', 'A1:SS4']) {
    const zip = await JSZip.loadAsync(bytes), sheet = 'xl/worksheets/sheet2.xml';
    const xml = await zip.file(sheet).async('string'); zip.file(sheet, xml.replace(/<dimension[^>]*\/>/, `<dimension ref="${range}"/>`));
    const parsed = await S.readDataFile(await zip.generateAsync({ type: 'nodebuffer' }), 'oversized.xlsx');
    assert.throws(() => C.readSheet(parsed, '运行数据', S), /250,000 行和 512 列/);
  }
  const inflated = Buffer.from(bytes), entry = inflated.indexOf(Buffer.from([0x50, 0x4b, 1, 2]));
  assert.ok(entry >= 0); inflated.writeUInt32LE(201 * 1024 * 1024, entry + 24);
  await assert.rejects(() => S.readDataFile(inflated, 'large.xlsx'), /解压后超过 200 MB/);
});

test('the same Excel adapter runs with browser globals and local vendor libraries without a network', async () => {
  const context = { TextEncoder, TextDecoder, Uint8Array, ArrayBuffer, DataView, setTimeout, clearTimeout, console };
  vm.createContext(context);
  for (const file of ['vendor/jszip.min.js', 'vendor/xlsx.full.min.js', 'src/sheet-reader.js']) vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
  assert.equal(context.XLSX.version, '0.20.3');
  const book = await context.SheetReader.readDataFile(new Uint8Array(await excelBytes('xlsb')), 'browser.xlsb');
  const data = C.readSheet(book, '运行数据', context.SheetReader);
  assert.equal(data.series[0].name, '温度[℃]'); assert.equal(data.rows[0].values[0], 30);
});
