// Integration checks for real ZIP/XML workbook bytes. Needs xml-js for a Node XML DOM shim.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const xml = require('xml-js');
global.JSZip = require('../vendor/jszip.min.js');
class XMLNode {
  constructor(node) { this.name = node.name || ''; this.attributes = node.attributes || {}; this.children = (node.elements || []).map(n => new XMLNode(n)); this.content = node.text ?? node.cdata ?? ''; }
  get textContent() { return this.content + this.children.map(n => n.textContent).join(''); }
  getAttribute(name) { return this.attributes[name] ?? null; }
  getAttributeNS(_, name) { return Object.entries(this.attributes).find(([key]) => key.split(':').at(-1) === name)?.[1] ?? null; }
  getElementsByTagNameNS(_, name) { return this.children.flatMap(child => [...(child.name.split(':').at(-1) === name ? [child] : []), ...child.getElementsByTagNameNS('*', name)]); }
  getElementsByTagName(name) { return this.getElementsByTagNameNS('*', name); }
}
global.DOMParser = class { parseFromString(s) { return new XMLNode(xml.xml2js(s)); } };
const S = require('../src/sheet-reader.js'), C = require('../src/core.js');
test('XLSX writer/reader round trip with Chinese names, XML escapes and missing cells', async () => {
  const wb = S.utils.book_new();
  S.utils.book_append_sheet(wb, S.utils.aoa_to_sheet([['时间', '压力<&>', '温度'], ['2026-01-01 08:00:00', 2.25, -20], ['2026-01-01 08:00:07', null, 90]]), '运行数据');
  S.utils.book_append_sheet(wb, S.utils.aoa_to_sheet([['时间', '转速'], [0, 6500]]), '转速表');
  const bytes = await S.write(wb), read = await S.read(bytes);
  assert.deepEqual(read.SheetNames, ['运行数据', '转速表']);
  const d = C.readSheet(read, '运行数据', S); assert.equal(d.rows[1].x - d.rows[0].x, 7000); assert.equal(d.series[0].name, '压力<&>'); assert.deepEqual(d.rows[1].values, [null, 90]);
});
test('Excel shared strings, numeric dates, styles, cached formulas and error values', async () => {
  const wb = S.utils.book_new(); S.utils.book_append_sheet(wb, S.utils.aoa_to_sheet([['时间', '温度'], [1, 1]]), 'Sheet1');
  const zip = await JSZip.loadAsync(await S.write(wb));
  zip.file('xl/sharedStrings.xml', '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>时间</t></si><si><r><t>环境</t></r><r><t>温度</t></r></si></sst>');
  zip.file('xl/styles.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2" s="1"><v>46023.5</v></c><c r="B2"><f>10+5</f><v>15</v></c></row><row r="3"><c r="A3" s="1"><v>46024</v></c><c r="B3" t="e"><v>#DIV/0!</v></c></row></sheetData></worksheet>');
  const read = await S.read(await zip.generateAsync({ type: 'uint8array' })); const d = C.readSheet(read, 'Sheet1', S);
  assert.equal(d.detected, 'excel'); assert.equal(d.series[0].name, '环境温度'); assert.equal(C.formatTime(d.rows[0].x), '2026-01-01 12:00:00'); assert.deepEqual(d.rows.map(r => r.values[0]), [15, null]);
});
test('malformed ZIP and legacy binary workbooks produce clear errors', async () => {
  await assert.rejects(() => S.read(new Uint8Array([0xd0, 0xcf, 0, 0])), /旧版/);
  await assert.rejects(() => S.read(new Uint8Array([0x50, 0x4b, 0, 0])));
});
