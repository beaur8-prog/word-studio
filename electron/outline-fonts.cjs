'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const families = new Map([
  ['Gulim', ['굴림', 'Gulim']], ['GulimChe', ['굴림체', 'GulimChe']],
  ['Dotum', ['돋움', 'Dotum']], ['DotumChe', ['돋움체', 'DotumChe']],
]);
function range(data, offset, size) {
  if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(size) || offset < 0 || size < 0 || offset + size > data.length)
    throw new Error('Invalid font table range');
  return data.subarray(offset, offset + size);
}
function checksum(data) {
  let value = 0;
  for (let i = 0; i < data.length; i += 4) value = (value + data.readUInt32BE(i)) >>> 0;
  return value;
}
/** Use the installed face's outline advances instead of its integer bitmap
 * strikes. Font files remain on the user's PC; none are shipped with the app. */
function extractOutlineFaces(data) {
  range(data, 0, 12);
  if (data.length > 64 * 1024 * 1024 || data.toString('ascii', 0, 4) !== 'ttcf') throw new Error('Expected a font collection');
  const count = data.readUInt32BE(8);
  if (!count || count > 32) throw new Error('Invalid font face count');
  range(data, 12, count * 4);
  const result = [];
  for (let face = 0; face < count; face++) {
    const offset = data.readUInt32BE(12 + face * 4);
    range(data, offset, 12);
    if (data.readUInt32BE(offset) !== 0x10000) continue;
    const n = data.readUInt16BE(offset + 4);
    if (!n || n > 256) throw new Error('Invalid font table count');
    range(data, offset + 12, n * 16);
    const tables = new Map();
    for (let i = 0; i < n; i++) {
      const p = offset + 12 + i * 16, tag = data.toString('ascii', p, p + 4);
      tables.set(tag, range(data, data.readUInt32BE(p + 8), data.readUInt32BE(p + 12)));
    }
    const names = tables.get('name'); if (!names) continue;
    range(names, 0, 6);
    const namesCount = names.readUInt16BE(2), start = names.readUInt16BE(4);
    range(names, 6, namesCount * 12);
    let family;
    for (let i = 0; i < namesCount; i++) {
      const p = 6 + i * 12;
      if (names.readUInt16BE(p) !== 3 || names.readUInt16BE(p + 6) !== 1) continue;
      const bytes = Buffer.from(range(names, start + names.readUInt16BE(p + 10), names.readUInt16BE(p + 8)));
      if (bytes.length % 2) continue;
      const name = bytes.swap16().toString('utf16le');
      family = [...families.entries()].find(([, aliases]) => aliases.includes(name));
      if (family) break;
    }
    if (!family || !tables.has('glyf') || !tables.has('head')) continue;
    for (const tag of ['EBDT', 'EBLC', 'EBSC', 'CBDT', 'CBLC', 'DSIG']) tables.delete(tag);
    const entries = [...tables].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0);
    let length = 12 + entries.length * 16;
    const output = Buffer.alloc(length + entries.reduce((size, [, table]) => size + Math.ceil(table.length / 4) * 4, 0));
    output.writeUInt32BE(0x10000); output.writeUInt16BE(entries.length, 4);
    const power = 2 ** Math.floor(Math.log2(entries.length));
    output.writeUInt16BE(power * 16, 6); output.writeUInt16BE(Math.log2(power), 8);
    output.writeUInt16BE(entries.length * 16 - power * 16, 10);
    let head = 0;
    entries.forEach(([tag, table], i) => {
      const p = 12 + i * 16;
      output.write(tag, p, 4, 'ascii'); output.writeUInt32BE(length, p + 8); output.writeUInt32BE(table.length, p + 12);
      table.copy(output, length);
      if (tag === 'head') { range(table, 0, 12); head = length; output.writeUInt32BE(0, head + 8); }
      const end = length + Math.ceil(table.length / 4) * 4;
      output.writeUInt32BE(checksum(output.subarray(length, end)), p + 4); length = end;
    });
    output.writeUInt32BE((0xb1b0afba - checksum(output)) >>> 0, head + 8);
    result.push({ families: family[1], data: output });
  }
  return result;
}
let cached;
function installedOutlineFonts() {
  if (process.platform !== 'win32') return Promise.resolve([]);
  // Fixed OS file, no renderer-provided path or font data.
  return cached ??= fs.readFile(path.join(process.env.SystemRoot || 'C:\\Windows', 'Fonts', 'gulim.ttc'))
    .then(extractOutlineFaces).catch(() => []);
}
module.exports = { extractOutlineFaces, installedOutlineFonts };
