'use strict';
// Optional manual CI validation uses public reproduction documents only.
// Treat payload contents as data, bound them, verify bytes and generate paths.
const fs = require('node:fs'), path = require('node:path'), crypto = require('node:crypto');
let raw = process.env.PUBLIC_DOCX_VALIDATION_PAYLOAD || '';
// Read dispatch inputs from the runner's event file. A step-level environment
// variable would make Actions print the entire document payload in its header.
if (!raw && process.env.GITHUB_EVENT_PATH) {
  const value = JSON.parse(fs.readFileSync(process.env.GITHUB_EVENT_PATH, 'utf8')).inputs?.public_docx_validation_payload;
  if (value !== undefined && typeof value !== 'string') throw Error('Invalid public DOCX validation input');
  raw = value || '';
}
// 10 textbox, 10 conditional-spacing, 2 edited-list, 6 list-line/level/suffix,
// 6 list-marker, 4 header/footer visibility, 2 table-style copies,
// 12 drawing-group and 4 short-paragraph fixtures with two saved copies each,
// plus7 row-policy and15 paragraph-keep inputs, each with two saved copies.
let count = 154;
if (raw) {
  if (raw.length > 60000) throw Error('Public DOCX validation payload is too large');
  const copies = JSON.parse(raw);
  if (!Array.isArray(copies) || !copies.length || copies.length > 2) throw Error('Expected one or two public DOCX copies');
  const folder = path.join('test-results', 'public-schema-inputs');
  fs.mkdirSync(folder, { recursive: true });
  for (const [index, copy] of copies.entries()) {
    if (typeof copy.base64 !== 'string' || !/^[A-Za-z0-9+/]+={0,2}$/.test(copy.base64) ||
        typeof copy.sha256 !== 'string' || !/^[a-f0-9]{64}$/.test(copy.sha256)) throw Error('Invalid public DOCX payload');
    const bytes = Buffer.from(copy.base64, 'base64');
    if (bytes.length < 4 || bytes.length > 45000 || bytes.readUInt32LE(0) !== 0x04034b50 ||
        crypto.createHash('sha256').update(bytes).digest('hex') !== copy.sha256) throw Error('Public DOCX bytes do not match their digest');
    fs.writeFileSync(path.join(folder, `textbox-save-public-${index + 1}.docx`), bytes);
    console.log(JSON.stringify({ publicCopy: index + 1, bytes: bytes.length, sha256: copy.sha256 }));
    count++;
  }
}
if (process.env.GITHUB_OUTPUT) fs.appendFileSync(process.env.GITHUB_OUTPUT, `expected=${count}\n`);
console.log(`Expected exported DOCX files: ${count}`);
