/* Local data adapters. Excel parsing: SheetJS CE (Apache-2.0); ZIP handling: JSZip (MIT). */
(function (root, factory) {
  const api = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SheetReader = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function (root) {
  'use strict';
  function encode_cell({ r, c }) { let s = ''; for (c++; c; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s + (r + 1); }
  function decode_cell(s) { const m = /^([A-Z]+)(\d+)$/i.exec(s); if (!m) throw new Error('Excel 单元格地址无效'); let c = 0; for (const ch of m[1].toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; return { r: +m[2] - 1, c: c - 1 }; }
  function decode_range(s) { const [a, b = a] = s.split(':'); return { s: decode_cell(a), e: decode_cell(b) }; }
  function aoa_to_sheet(rows) {
    const sheet = {}; let maxCol = 0;
    rows.forEach((row, r) => row.forEach((v, c) => { maxCol = Math.max(maxCol, c); if (v !== null && v !== undefined) sheet[encode_cell({ r, c })] = { v, t: typeof v === 'number' ? 'n' : typeof v === 'boolean' ? 'b' : 's' }; }));
    if (rows.length) sheet['!ref'] = `A1:${encode_cell({ r: rows.length - 1, c: maxCol })}`;
    return sheet;
  }
  function sheet_to_json(sheet, options = {}) {
    if (!sheet['!ref']) return [];
    const range = options.range || decode_range(sheet['!ref']), rows = [];
    for (let r = range.s.r; r <= range.e.r; r++) { const row = []; for (let c = range.s.c; c <= range.e.c; c++) row.push(sheet[encode_cell({ r, c })]?.v ?? options.defval ?? null); rows.push(row); }
    return rows;
  }
  const utils = { encode_cell, decode_cell, decode_range, aoa_to_sheet, sheet_to_json, book_new: () => ({ SheetNames: [], Sheets: {} }), book_append_sheet: (wb, sheet, name) => { wb.SheetNames.push(name); wb.Sheets[name] = sheet; } };
  function parse_date_code(n, options = {}) {
    if (!Number.isFinite(n) || n < 0 || n > 2958465) return null;
    let whole = Math.floor(n), ms = Math.round((n - whole) * 86400000);
    const epoch = options.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    if (!options.date1904 && whole >= 60) whole--; // Excel's historical 1900 leap-year convention.
    const d = new Date(epoch + whole * 86400000 + ms);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(), M: d.getUTCMinutes(), S: d.getUTCSeconds(), u: d.getUTCMilliseconds() / 1000 };
  }
  const SSF = { parse_date_code, is_date: z => /[ymdhis]/i.test(String(z).replace(/"[^"]*"|\\.|\[(?![hms]+\])[^\]]*\]/gi, '')) };
  function xml(text) { const doc = new DOMParser().parseFromString(text, 'application/xml'); if (doc.getElementsByTagName('parsererror').length) throw new Error('Excel XML 内容损坏'); return doc; }
  function elements(parent, name) { return Array.from(parent.getElementsByTagNameNS('*', name)); }
  function textRuns(parent) { return elements(parent, 't').map(t => t.textContent).join(''); }
  function normalizePath(base, target) {
    const parts = (target.startsWith('/') ? target.slice(1) : base + target).split('/'), out = [];
    for (const part of parts) { if (part === '..') out.pop(); else if (part && part !== '.') out.push(part); }
    return out.join('/');
  }
  function detectDelimiter(text) {
    const counts = new Map([['\t', 0], [',', 0], [';', 0]]);
    let quoted = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (ch === '"') { if (quoted && text[i + 1] === '"') i++; else quoted = !quoted; }
      else if (!quoted && (ch === '\n' || ch === '\r')) break;
      else if (!quoted && counts.has(ch)) counts.set(ch, counts.get(ch) + 1);
    }
    return counts.get('\t') ? '\t' : counts.get(';') > counts.get(',') ? ';' : ',';
  }
  function parseCSV(text, delimiter, limits = {}) {
    text = text.replace(/^\uFEFF/, '');
    delimiter ??= detectDelimiter(text);
    const rows = []; let row = [], field = '', quoted = false;
    function pushField() {
      if (row.length >= (limits.columns ?? Infinity)) throw new Error('最多支持 512 列，请缩小数据范围。');
      row.push(field); field = '';
    }
    function pushRow() {
      if (rows.length >= (limits.rows ?? Infinity)) throw new Error('最多支持 250,000 行数据，请拆分数据。');
      rows.push(row); row = [];
    }
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (ch === '"') { if (quoted && text[i + 1] === '"') { field += '"'; i++; } else if (!field || quoted) quoted = !quoted; else field += ch; }
      else if (ch === delimiter && !quoted) pushField();
      else if ((ch === '\r' || ch === '\n') && !quoted) { if (ch === '\r' && text[i + 1] === '\n') i++; pushField(); pushRow(); }
      else field += ch;
    }
    if (quoted) throw new Error('表格中的引号未闭合，请重新复制完整的单元格区域。');
    if (field || row.length) { pushField(); pushRow(); }
    return rows;
  }
  function readPasted(text) {
    if (typeof text !== 'string' || !text.trim()) throw new Error('请先粘贴从 Excel 复制的表格数据。');
    const maxBytes = 50 * 1024 * 1024;
    if (text.length > maxBytes || new TextEncoder().encode(text).length > maxBytes) throw new Error('粘贴内容超过 50 MB，请缩小复制范围。');
    // Excel clipboard text is TSV. Force tabs so punctuation in headers is preserved.
    const rows = parseCSV(text, '\t', { rows: 250001, columns: 512 });
    if (rows.length < 2 || rows[0].length < 2) throw new Error('请复制包含列名和数据的至少两列：第一列为时间，其余列为数值。');
    const extra = rows.findIndex((row, index) => index > 0 && row.length > rows[0].length);
    if (extra >= 0) throw new Error(`第 ${extra + 1} 行的列数多于首行，请连同全部列名重新复制。`);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, aoa_to_sheet(rows), '粘贴数据');
    return wb;
  }
  function decodeText(bytes) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (data.length > 50 * 1024 * 1024) throw new Error('文件超过 50 MB，请拆分数据。');
    if ((data[0] === 0x50 && data[1] === 0x4b) || (data[0] === 0xd0 && data[1] === 0xcf)) throw new Error('此文件含二进制表格或压缩内容，请保留原始 Excel 扩展名上传，不能只修改扩展名。');
    let text;
    if (data[0] === 0xff && data[1] === 0xfe) text = new TextDecoder('utf-16le').decode(data);
    else if (data[0] === 0xfe && data[1] === 0xff) text = new TextDecoder('utf-16be').decode(data);
    else {
      try { text = new TextDecoder('utf-8', { fatal: true }).decode(data); }
      catch { text = new TextDecoder('gb18030', { fatal: true }).decode(data); }
    }
    text = text.replace(/^\uFEFF/, '');
    if (text.includes('\u0000')) throw new Error('文件不是可识别的文本，请使用 UTF-8、GBK 或带 BOM 的 UTF-16 编码。');
    return text;
  }
  function tableWorkbook(rows) {
    if (!Array.isArray(rows) || rows.length < 2 || !Array.isArray(rows[0]) || rows[0].length < 2) throw new Error('数据需要首行列名和至少一行记录，第一列为时间，其余列为数值。');
    if (rows.length > 250001) throw new Error('最多支持 250,000 行数据，请拆分文件。');
    const width = rows[0].length;
    if (width > 512) throw new Error('最多支持 512 列，请缩小数据范围。');
    rows.forEach((row, index) => {
      if (!Array.isArray(row)) throw new Error(`第 ${index + 1} 行必须是数组。`);
      if (row.length > width) throw new Error(`第 ${index + 1} 行的列数多于列名，请补齐首行列名。`);
      for (const cell of row) if (cell !== null && !['string', 'number', 'boolean'].includes(typeof cell)) throw new Error(`第 ${index + 1} 行含有嵌套对象或数组，单元格应为文本、数字或 null。`);
    });
    const workbook = utils.book_new();
    utils.book_append_sheet(workbook, aoa_to_sheet(rows), '数据');
    return workbook;
  }
  function jsonTable(value) {
    if (value?.schema === 'curve-viewer' || value?.schema === 'curve-viewer-bundle' || (Array.isArray(value) && value.length && value.every(item => item?.schema === 'curve-viewer'))) {
      const error = new Error('这是显示配置 JSON，请导入到配置列表。'); error.code = 'DISPLAY_CONFIG'; throw error;
    }
    if (value && !Array.isArray(value) && Array.isArray(value.columns) && Array.isArray(value.data)) {
      if (value.data.length > 250000) throw new Error('最多支持 250,000 行数据，请拆分文件。');
      return [value.columns].concat(value.data);
    }
    if (!Array.isArray(value) || !value.length) throw new Error('JSON 需为非空对象数组、带列名的二维数组，或包含 columns 和 data 的对象。');
    if (Array.isArray(value[0])) return value;
    if (value.length > 250000) throw new Error('最多支持 250,000 行数据，请拆分文件。');
    const names = new Set();
    for (const row of value) {
      if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error('JSON 对象数组中的每一项都必须是一条记录对象。');
      for (const key of Object.keys(row)) { names.add(key); if (names.size > 512) throw new Error('最多支持 512 列，请缩小数据范围。'); }
    }
    const columns = [...names];
    const timeIndex = columns.findIndex(name => /^(时间|time|timestamp|datetime|date)$/i.test(name.trim()));
    if (timeIndex > 0) columns.unshift(...columns.splice(timeIndex, 1));
    return [columns].concat(value.map(row => columns.map(name => Object.prototype.hasOwnProperty.call(row, name) ? row[name] : null)));
  }
  const pmsdCRCTable = Uint32Array.from({ length: 256 }, (_, value) => {
    for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    return value >>> 0;
  });
  function pmsdCRC(data, start, end) {
    let crc = 0xffffffff;
    for (let i = start; i < end; i++) crc = pmsdCRCTable[(crc ^ data[i]) & 255] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }
  function pmsdTimeZone(filename, startMs) {
    // The observed format stores Unix nanoseconds but no explicit time-zone field.
    // Original filenames contain the local recording start, independently checked against the header.
    const match = /_(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})_(\d{3})(?:_|\.pmsd$)/i.exec(filename);
    if (match) {
      const [, y, mo, d, h, m, s, ms] = match.map(Number);
      const local = Date.UTC(y, mo - 1, d, h, m, s, ms), date = new Date(local);
      const offset = Math.round((local - startMs) / 60000);
      if (date.getUTCFullYear() === y && date.getUTCMonth() === mo - 1 && date.getUTCDate() === d && h < 24 && m < 60 && s < 60 &&
        Math.abs(offset) <= 840 && Math.abs(local - startMs - offset * 60000) <= 1000) return { offset, source: 'filename' };
    }
    return { offset: null, source: 'local' };
  }
  function pmsdTime(ms, zone) {
    if (zone.offset !== null) return new Date(ms + zone.offset * 60000).toISOString().slice(0, 23).replace('T', ' ');
    const date = new Date(ms), pad = (n, length = 2) => String(n).padStart(length, '0');
    return `${pad(date.getFullYear(), 4)}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}.${pad(date.getMilliseconds(), 3)}`;
  }
  function readPMSD(bytes, filename) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (data.length > 50 * 1024 * 1024) throw new Error('文件超过 50 MB，请拆分数据。');
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const signature = (at, text) => Array.from(text).every((char, i) => data[at + i] === char.charCodeAt(0));
    if (data.length < 52 || !signature(0, 'PMSAMPLE')) throw new Error('不是支持的 PMSD 采样文件，文件头应为 PMSAMPLE。');
    if (view.getUint16(8, true) !== 2 || view.getUint16(10, true) !== 3) throw new Error('暂不支持此 PMSD 版本或存储方式，目前支持 PMSAMPLE v2 / 标志 3。');
    const headerEnd = view.getUint32(12, true), count = view.getUint32(16, true);
    if (count < 1 || count > 511) throw new Error('PMSD 需要 1～511 个数据列（含时间列最多 512 列）。');
    if (headerEnd < 52 || headerEnd > data.length) throw new Error('PMSD 文件头长度无效或文件不完整。');
    if (pmsdCRC(data, 32, headerEnd) !== view.getUint32(28, true)) throw new Error('PMSD 字段信息校验失败，文件可能损坏。');
    const decoder = new TextDecoder('utf-8', { fatal: true }), fields = [];
    let position = 52;
    function readString() {
      if (position + 2 > headerEnd) throw new Error('PMSD 字段信息不完整。');
      const length = view.getUint16(position, true); position += 2;
      if (position + length > headerEnd) throw new Error('PMSD 字段文本长度无效。');
      let value;
      try { value = decoder.decode(data.subarray(position, position + length)); } catch { throw new Error('PMSD 字段名称不是有效的 UTF-8 文本。'); }
      position += length; return value;
    }
    for (let i = 0; i < count; i++) {
      if (position + 6 > headerEnd) throw new Error('PMSD 字段信息不完整。');
      const index = view.getUint16(position, true), id = view.getUint32(position + 2, true); position += 6;
      if (index !== i) throw new Error('暂不支持此 PMSD 字段索引结构。');
      fields.push({ id, model: readString(), device: readString(), name: readString(), unit: readString(), type: readString() });
    }
    if (position !== headerEnd) throw new Error('PMSD 文件头包含未识别的信息，未载入数据。');
    const zone = pmsdTimeZone(filename, Number(view.getBigInt64(20, true) / 1000000n));
    const rows = [['时间', ...fields.map(field => field.unit ? `${field.model}/${field.device}/${field.name} (ID ${field.id}) [${field.unit}]` : field.name)]];
    const maskBytes = Math.ceil(count / 8), recordBytes = 32 + maskBytes + count * 8;
    let missing = 0, emptyRecords = 0, endRecord = null;
    while (position < data.length) {
      const rowNumber = rows.length;
      // Completed captures may end with END1: length, uint64 sample count,
      // int64 end time (ns), and CRC32 of the entire SMP1 byte region.
      // It is a trailer, not an additional sample. Check it before the row limit.
      if (signature(position, 'END1')) {
        if (position + 8 > data.length) throw new Error('PMSD END1 结束块不完整，文件可能被截断。');
        const length = view.getUint32(position + 4, true);
        if (length !== 28) throw new Error('PMSD END1 结束块长度无效，当前支持 28 字节结构。');
        if (position + length > data.length) throw new Error('PMSD END1 结束块不完整，文件可能被截断。');
        if (position + length !== data.length) throw new Error('PMSD END1 结束块后仍有多余数据，未载入文件。');
        if (view.getBigUint64(position + 8, true) !== BigInt(rows.length - 1)) throw new Error('PMSD END1 结束块的采样计数与实际记录数不一致。');
        if (pmsdCRC(data, headerEnd, position) !== view.getUint32(position + 24, true)) throw new Error('PMSD END1 结束块的采样区校验失败，文件可能损坏。');
        endRecord = { records: rows.length - 1, time: pmsdTime(Number(view.getBigInt64(position + 16, true) / 1000000n), zone) };
        position += length; break;
      }
      if (rowNumber > 250000) throw new Error('最多支持 250,000 行数据，请拆分文件。');
      if (position + 8 > data.length || !signature(position, 'SMP1')) throw new Error(`PMSD 第 ${rowNumber} 条记录不完整或类型不支持。`);
      const length = view.getUint32(position + 4, true);
      if (length !== recordBytes || position + length > data.length) throw new Error(`PMSD 第 ${rowNumber} 条记录长度无效或文件被截断。`);
      if (view.getUint32(position + 24, true) !== maskBytes) throw new Error(`PMSD 第 ${rowNumber} 条记录的有效性标记长度无效。`);
      if (pmsdCRC(data, position + 8, position + length - 4) !== view.getUint32(position + length - 4, true)) throw new Error(`PMSD 第 ${rowNumber} 条记录校验失败，文件可能损坏。`);
      const ms = Number(view.getBigInt64(position + 16, true) / 1000000n);
      const row = [pmsdTime(ms, zone)], valuesStart = position + 28 + maskBytes;
      let valid = 0;
      for (let column = 0; column < count; column++) {
        const present = data[position + 28 + (column >> 3)] & (1 << (column & 7));
        const value = view.getFloat64(valuesStart + column * 8, true);
        if (present && Number.isFinite(value)) { row.push(value); valid++; }
        else { row.push(null); missing++; }
      }
      if (!valid) emptyRecords++;
      rows.push(row); position += length;
    }
    const workbook = tableWorkbook(rows);
    const offset = zone.offset, sign = offset < 0 ? '−' : '+';
    const timeZone = offset === null ? '浏览器本地时区' : `UTC${sign}${String(Math.floor(Math.abs(offset) / 60)).padStart(2, '0')}:${String(Math.abs(offset) % 60).padStart(2, '0')}`;
    workbook.ImportInfo = { format: 'PMSD', version: 2, timeZone, timeZoneSource: zone.source, records: rows.length - 1, missing, emptyRecords, endRecord };
    return workbook;
  }
  const EXCEL_EXTENSIONS = ['xlsx', 'xls', 'xlsm', 'xlsb', 'xltx', 'xltm', 'xlt', 'xlw', 'xml'];
  const DATA_EXTENSIONS = ['csv', 'json', 'txt', 'pmsd', ...EXCEL_EXTENSIONS];
  function supportsDataFile(filename) { return DATA_EXTENSIONS.includes(/\.([^.]+)$/.exec(filename)?.[1].toLowerCase()); }
  async function readExcelFile(bytes) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (data.length > 50 * 1024 * 1024) throw new Error('文件超过 50 MB，请拆分数据。');
    if (data.length < 4) throw new Error('Excel 文件为空或不完整。');
    const signature = Array.from(data.subarray(0, 8)).map(byte => byte.toString(16).padStart(2, '0')).join('');
    const zipped = signature.startsWith('504b0304'), compound = signature === 'd0cf11e0a1b11ae1';
    const biff = [0x0009, 0x0209, 0x0409, 0x0809].includes(data[0] | (data[1] << 8));
    let payload = data, type = 'array';
    if (zipped) {
      const Zip = root.JSZip || (typeof module === 'object' && module.exports ? require('../vendor/jszip.min.js') : null);
      if (!Zip) throw new Error('Excel 解压组件未能加载，请重新打开完整 HTML。');
      let zip;
      try { zip = await Zip.loadAsync(data); } catch { throw new Error('Excel 压缩结构损坏或文件不完整。'); }
      const entries = Object.values(zip.files);
      if (entries.length > 10000) throw new Error('工作簿包含过多文件，请拆分数据。');
      const expandedSize = entries.reduce((sum, entry) => sum + (entry._data?.uncompressedSize || 0), 0);
      if (expandedSize > 200 * 1024 * 1024) throw new Error('工作簿解压后超过 200 MB，请拆分数据。');
      if (!zip.file('xl/workbook.xml') && !zip.file('xl/workbook.bin')) throw new Error('压缩文件中没有 Excel 工作簿，请上传原始 Excel 文件。');
    } else if (!compound && !biff) {
      let text;
      try { text = decodeText(data); } catch { throw new Error('不是可识别的 Excel 工作簿，文件可能损坏。'); }
      const spreadsheetML = /<(?:[\w.-]+:)?Workbook\b/i.test(text) && /urn:schemas-microsoft-com:office:spreadsheet/i.test(text);
      if (!spreadsheetML) throw new Error('不是可识别的 Excel 工作簿；文本表格请使用 CSV 或 TXT 扩展名，XML 需为 Excel 2003 XML 格式。');
      payload = text; type = 'string';
    }
    const Excel = root.XLSX || (typeof module === 'object' && module.exports ? require('../vendor/xlsx.full.min.js') : null);
    if (!Excel?.read) throw new Error('Excel 解析组件未能加载，请重新打开完整 HTML。');
    let workbook;
    try {
      workbook = Excel.read(payload, { type, cellDates: false, cellNF: true, cellText: false, cellHTML: false, cellFormula: false, bookVBA: false, bookDeps: false, dense: false, sheetRows: 250002 });
    } catch (error) {
      if (/password|encrypt|filepass/i.test(error.message || '')) throw new Error('Excel 文件已加密，请另存为不带打开密码的副本后上传。');
      throw new Error('Excel 文件解析失败，请确认文件完整，并尝试用 Excel 另存后重新上传。');
    }
    if (!workbook.SheetNames?.length) throw new Error('Excel 文件中没有可读取的工作表。');
    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      // Retain the true dimensions so readSheet rejects oversized tables instead of plotting a truncated prefix.
      if (sheet?.['!fullref']) sheet['!ref'] = sheet['!fullref'];
    }
    workbook.ImportInfo = { format: 'Excel' };
    return workbook;
  }
  function readDataFile(bytes, filename) {
    const extension = /\.([^.]+)$/.exec(filename)?.[1].toLowerCase();
    if (!supportsDataFile(filename)) throw new Error('支持 Excel、CSV、JSON、TXT、PMSD 文件，请选择支持的扩展名。');
    if (EXCEL_EXTENSIONS.includes(extension)) return readExcelFile(bytes);
    if (extension === 'pmsd') return readPMSD(bytes, filename);
    let text = decodeText(bytes);
    if (!text.trim()) throw new Error('文件为空，请选择包含表格数据的文件。');
    if (extension === 'json') {
      let value;
      try { value = JSON.parse(text); } catch { throw new Error('JSON 格式无效，请检查逗号、引号和括号。'); }
      return tableWorkbook(jsonTable(value));
    }
    let delimiter;
    const separator = /^sep=([,;\t])\r?\n/i.exec(text);
    if (separator) { delimiter = separator[1]; text = text.slice(separator[0].length); }
    return tableWorkbook(parseCSV(text, delimiter, { rows: 250001, columns: 512 }));
  }
  function writeCSV(workbook, sheetName = workbook.SheetNames[0]) {
    const rows = sheet_to_json(workbook.Sheets[sheetName]);
    const field = value => { const text = String(value ?? ''); return /[",\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text; };
    return '\uFEFF' + rows.map(row => row.map(field).join(',')).join('\r\n') + '\r\n';
  }
  async function read(bytes) {
    const data = new Uint8Array(bytes);
    if (data[0] === 0xd0 && data[1] === 0xcf) throw new Error('旧版 .xls 或加密工作簿暂不支持，请另存为未加密的 .xlsx 后导入。');
    if (data[0] !== 0x50 || data[1] !== 0x4b) {
      let text;
      if (data[0] === 0xff && data[1] === 0xfe) text = new TextDecoder('utf-16le').decode(data);
      else { text = new TextDecoder('utf-8').decode(data); if (text.includes('\uFFFD')) text = new TextDecoder('gb18030').decode(data); }
      if (/^\s*</.test(text)) throw new Error('请将此文件另存为 .xlsx 或 CSV 格式。');
      const wb = utils.book_new(); utils.book_append_sheet(wb, aoa_to_sheet(parseCSV(text)), '数据'); return wb;
    }
    const zip = await root.JSZip.loadAsync(data);
    const zipEntries = Object.values(zip.files);
    let size = 0;
    for (const entry of zipEntries) { size += entry._data?.uncompressedSize || 0; if (size > 200 * 1024 * 1024) throw new Error('工作簿解压后超过 200 MB，请拆分数据。'); }
    if (zipEntries.length > 10000) throw new Error('工作簿包含过多文件，请拆分数据。');
    async function get(file, required = false) {
      const entry = zip.file(file);
      if (!entry) { if (required) throw new Error('不是有效的 .xlsx 工作簿，缺少 ' + file); return null; }
      const value = await entry.async('string'); return xml(value);
    }
    const book = await get('xl/workbook.xml', true), rels = await get('xl/_rels/workbook.xml.rels', true);
    const relationships = new Map(elements(rels, 'Relationship').filter(r => r.getAttribute('TargetMode') !== 'External').map(r => [r.getAttribute('Id'), normalizePath('xl/', r.getAttribute('Target'))]));
    const sharedPath = elements(rels, 'Relationship').find(r => /\/sharedStrings$/.test(r.getAttribute('Type')))?.getAttribute('Target');
    const sharedDoc = await get(sharedPath ? normalizePath('xl/', sharedPath) : 'xl/sharedStrings.xml');
    const strings = sharedDoc ? elements(sharedDoc, 'si').map(textRuns) : [];
    const styleDoc = await get('xl/styles.xml');
    const formats = new Map([[14, 'yyyy-mm-dd'], [15, 'd-mmm-yy'], [16, 'd-mmm'], [17, 'mmm-yy'], [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'], [20, 'h:mm'], [21, 'h:mm:ss'], [22, 'yyyy-mm-dd h:mm'], [45, 'mm:ss'], [46, '[h]:mm:ss'], [47, 'mmss.0']]);
    for (let id = 27; id <= 36; id++) formats.set(id, 'yyyy-mm-dd');
    for (let id = 50; id <= 58; id++) formats.set(id, 'yyyy-mm-dd');
    const styles = [];
    if (styleDoc) {
      elements(styleDoc, 'numFmt').forEach(f => formats.set(+f.getAttribute('numFmtId'), f.getAttribute('formatCode')));
      const xfs = elements(styleDoc, 'cellXfs')[0]; if (xfs) elements(xfs, 'xf').forEach(xf => styles.push(formats.get(+xf.getAttribute('numFmtId')) || 'General'));
    }
    const date1904 = ['1', 'true'].includes(elements(book, 'workbookPr')[0]?.getAttribute('date1904'));
    const wb = { SheetNames: [], Sheets: {}, Workbook: { WBProps: { date1904 } } };
    for (const spec of elements(book, 'sheet')) {
      const name = spec.getAttribute('name'), id = spec.getAttribute('r:id') || spec.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      const target = relationships.get(id);
      if (!target || !/worksheets\//.test(target)) continue;
      const doc = await get(target, true), sheet = {}; let maxR = 0, maxC = 0, any = false;
      for (const cell of elements(doc, 'c')) {
        const ref = cell.getAttribute('r'); if (!ref) continue;
        const pos = decode_cell(ref);
        if (pos.r > 250000 || pos.c > 511) throw new Error('工作表超出 250,000 行或 512 列，请先缩小数据范围。');
        const type = cell.getAttribute('t') || 'n'; const raw = elements(cell, 'v')[0]?.textContent;
        let value = null, t = type;
        if (type === 'inlineStr') { value = textRuns(cell); t = 's'; }
        else if (type === 's') { value = strings[+raw] ?? ''; t = 's'; }
        else if (type === 'b') value = raw === '1';
        else if (type === 'str' || type === 'e' || type === 'd') { value = raw ?? null; if (type === 'd') t = 's'; }
        else if (raw !== undefined && raw !== '') { value = Number(raw); if (!Number.isFinite(value)) value = null; }
        if (value === null) continue;
        sheet[ref] = { v: value, t, z: styles[+(cell.getAttribute('s') || 0)] || 'General' };
        maxR = Math.max(maxR, pos.r); maxC = Math.max(maxC, pos.c); any = true;
      }
      if (any) sheet['!ref'] = 'A1:' + encode_cell({ r: maxR, c: maxC });
      utils.book_append_sheet(wb, sheet, name);
    }
    return wb;
  }
  const esc = value => String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
  async function write(wb) {
    const zip = new root.JSZip(), ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    zip.file('[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${wb.SheetNames.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
    zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${wb.SheetNames.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wb.SheetNames.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}</Relationships>`);
    wb.SheetNames.forEach((name, i) => {
      const sheet = wb.Sheets[name], rows = sheet_to_json(sheet);
      const content = rows.map((row, r) => `<row r="${r + 1}">${row.map((v, c) => v == null ? '' : typeof v === 'number' ? `<c r="${encode_cell({ r, c })}"><v>${v}</v></c>` : `<c r="${encode_cell({ r, c })}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`).join('')}</row>`).join('');
      zip.file(`xl/worksheets/sheet${i + 1}.xml`, `<?xml version="1.0"?><worksheet xmlns="${ns}"><sheetData>${content}</sheetData></worksheet>`);
    });
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  return { utils, SSF, DATA_EXTENSIONS, EXCEL_EXTENSIONS, supportsDataFile, read, write, parseCSV, readPasted, readDataFile, writeCSV };
});
