'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { extractOutlineFaces } = require('../electron/outline-fonts.cjs');

function collection() {
  const name = Buffer.from('DotumChe', 'utf16le').swap16(), names = Buffer.alloc(18 + name.length);
  names.writeUInt16BE(1, 2); names.writeUInt16BE(18, 4); names.writeUInt16BE(3, 6);
  names.writeUInt16BE(1, 8); names.writeUInt16BE(1, 12); names.writeUInt16BE(name.length, 14); name.copy(names, 18);
  const tables = [['head', Buffer.alloc(54)], ['name', names], ['glyf', Buffer.from([1, 2, 3, 4])], ['EBDT', Buffer.alloc(5)], ['EBLC', Buffer.alloc(7)]];
  const bytes = Buffer.alloc(512); bytes.write('ttcf'); bytes.writeUInt32BE(0x10000, 4); bytes.writeUInt32BE(1, 8); bytes.writeUInt32BE(16, 12);
  bytes.writeUInt32BE(0x10000, 16); bytes.writeUInt16BE(tables.length, 20);
  let offset = 28 + tables.length * 16;
  tables.forEach(([tag, data], i) => { const p = 28 + i * 16; bytes.write(tag, p); bytes.writeUInt32BE(offset, p + 8); bytes.writeUInt32BE(data.length, p + 12); data.copy(bytes, offset); offset += data.length; });
  return bytes;
}
test('extracts an allowlisted outline face with valid checksums without mutating its collection', () => {
  const input = collection(), before = Buffer.from(input), [font] = extractOutlineFaces(input);
  assert.deepEqual(font.families, ['돋움체', 'DotumChe']); assert.deepEqual(input, before);
  const output = font.data, count = output.readUInt16BE(4), tags = [];
  let sum = 0; for (let i = 0; i < output.length; i += 4) sum = (sum + output.readUInt32BE(i)) >>> 0;
  assert.equal(sum, 0xb1b0afba);
  for (let i = 0; i < count; i++) tags.push(output.toString('ascii', 12 + i * 16, 16 + i * 16));
  assert.deepEqual(tags, ['glyf', 'head', 'name']);
});
test('rejects truncated collections and out-of-bounds table data', () => {
  assert.throws(() => extractOutlineFaces(Buffer.alloc(4)), /range/);
  const input = collection(); input.writeUInt32BE(0xfffffff0, 36);
  assert.throws(() => extractOutlineFaces(input), /range/);
});
