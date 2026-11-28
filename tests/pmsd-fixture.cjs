// Synthetic files follow the byte layout independently observed in the supplied PMSD sample.
// No device identifier or original recordings are included in the delivery package.
const crc32 = buffer => {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
};
const START_MS = Date.UTC(2026, 10, 25, 10, 0, 40, 865);
const FILENAME = 'capture_20261125_180040_865_sid_1_part0001.pmsd';
function makePMSD(options = {}) {
  const fields = options.fields || [{ name: '压力', unit: 'bar', type: 'UINT16' }, { name: '温度', unit: '℃', type: 'INT16' }];
  const rows = options.rows || [{ timeMs: START_MS + 99, values: [1.4, -18.9] }, { timeMs: START_MS + 736, values: [1.6, 43] }];
  const text = value => { const bytes = Buffer.from(value, 'utf8'), length = Buffer.alloc(2); length.writeUInt16LE(bytes.length); return Buffer.concat([length, bytes]); };
  const entries = fields.map((field, i) => {
    const prefix = Buffer.alloc(6); prefix.writeUInt16LE(i, 0); prefix.writeUInt32LE(field.id ?? 34000 + i, 2);
    return Buffer.concat([prefix, ...[field.model ?? 'HPODU', field.device ?? '设备', field.name, field.unit ?? '1', field.type ?? 'UINT16'].map(text)]);
  });
  const header = Buffer.concat([Buffer.alloc(52), ...entries]);
  header.write('PMSAMPLE', 0, 'ascii'); header.writeUInt16LE(options.version ?? 2, 8); header.writeUInt16LE(options.flags ?? 3, 10);
  header.writeUInt32LE(header.length, 12); header.writeUInt32LE(fields.length, 16);
  header.writeBigInt64LE(BigInt(options.startMs ?? START_MS) * 1000000n + 325568n, 20);
  header.writeBigUInt64LE(1n, 32); header.writeUInt32LE(1, 40); header.writeUInt32LE(crc32(header.subarray(32)), 28);
  const maskBytes = Math.ceil(fields.length / 8);
  const records = rows.map((row, i) => {
    const record = Buffer.alloc(32 + maskBytes + fields.length * 8);
    record.write('SMP1', 0, 'ascii'); record.writeUInt32LE(record.length, 4); record.writeBigUInt64LE(BigInt(row.sequence ?? i), 8);
    record.writeBigInt64LE(BigInt(row.timeMs) * 1000000n + (row.extraNs ?? 585472n), 16); record.writeUInt32LE(maskBytes, 24);
    for (let c = 0; c < fields.length; c++) {
      const value = row.values[c], valid = row.valid ? row.valid[c] : value !== null && value !== undefined;
      if (valid) record[28 + (c >> 3)] |= 1 << (c & 7);
      record.writeDoubleLE(value ?? 999, 28 + maskBytes + c * 8);
    }
    record.writeUInt32LE(crc32(record.subarray(8, record.length - 4)), record.length - 4); return record;
  });
  const samples = Buffer.concat(records);
  if (!options.endRecord) return Buffer.concat([header, samples]);
  const trailer = Buffer.alloc(28);
  trailer.write('END1', 0, 'ascii'); trailer.writeUInt32LE(28, 4); trailer.writeBigUInt64LE(BigInt(rows.length), 8);
  trailer.writeBigInt64LE(BigInt((rows.at(-1)?.timeMs ?? options.startMs ?? START_MS) + 500) * 1000000n, 16);
  trailer.writeUInt32LE(crc32(samples), 24);
  return Buffer.concat([header, samples, trailer]);
}
function repairHeader(bytes) { bytes.writeUInt32LE(crc32(bytes.subarray(32, bytes.readUInt32LE(12))), 28); }
module.exports = { makePMSD, crc32, repairHeader, START_MS, FILENAME };
