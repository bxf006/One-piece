const X = require('../vendor/xlsx.full.min.js');
const JSZip = require('../vendor/jszip.min.js');
function excelBook() {
  const book = X.utils.book_new();
  X.utils.book_append_sheet(book, X.utils.aoa_to_sheet([['说明'], ['首个工作表没有曲线数据']]), '说明');
  const sheet = X.utils.aoa_to_sheet([
    ['时间', 'HPODU/设备/温度 (ID 34001) [℃]', '压力'],
    [46024, 30, 3], [46023.5, -20, 1.4], [46024, 40, null]
  ]);
  for (const cell of ['A2', 'A3', 'A4']) sheet[cell].z = 'yyyy-mm-dd hh:mm:ss';
  X.utils.book_append_sheet(book, sheet, '运行数据');
  return book;
}
async function excelBytes(extension, book = excelBook()) {
  const type = ({ xls: 'biff8', xlt: 'biff8', xltx: 'xlsx', xltm: 'xlsm', xml: 'xlml' })[extension] || extension;
  const bytes = X.write(book, { type: 'buffer', bookType: type });
  if (extension !== 'xltx' && extension !== 'xltm') return bytes;
  const zip = await JSZip.loadAsync(bytes), file = '[Content_Types].xml';
  const content = await zip.file(file).async('string');
  const updated = extension === 'xltx'
    ? content.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml')
    : content.replace('application/vnd.ms-excel.sheet.macroEnabled.main+xml', 'application/vnd.ms-excel.template.macroEnabled.main+xml');
  if (updated === content) throw Error('Template content type was not changed');
  zip.file(file, updated); return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
module.exports = { X, JSZip, excelBook, excelBytes };
