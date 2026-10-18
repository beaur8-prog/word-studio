'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const root = path.resolve(__dirname, '..');

// Test actual cmap coverage: a successful FontFace load alone does not prove
// the font contains a glyph and can still render a missing-character box.
function codepoints(buffer) {
  const u16 = o => buffer.readUInt16BE(o), u32 = o => buffer.readUInt32BE(o);
  let cmap = 0;
  for (let i = 0; i < u16(4); i++) if (buffer.toString('ascii', 12+i*16, 16+i*16) === 'cmap') cmap = u32(20+i*16);
  assert.ok(cmap); const points = new Set();
  for (let i = 0; i < u16(cmap+2); i++) {
    const platform = u16(cmap+4+i*8); if (![0,3].includes(platform)) continue;
    const o = cmap + u32(cmap+8+i*8), format = u16(o);
    if (format === 12) for (let j = 0; j < u32(o+12); j++) {
      const first = u32(o+16+j*12), last = u32(o+20+j*12), gid = u32(o+24+j*12);
      for (let cp = first; cp <= last; cp++) if (gid+cp-first) points.add(cp);
    }
    if (format === 4) {
      const n = u16(o+6)/2, ends = o+14, starts = ends+2*n+2, deltas = starts+2*n, ranges = deltas+2*n;
      for (let j = 0; j < n; j++) for (let cp = u16(starts+2*j); cp <= u16(ends+2*j) && cp < 65535; cp++) {
        const range = u16(ranges+2*j), glyph = range ? u16(ranges+2*j+range+2*(cp-u16(starts+2*j))) : cp;
        if (glyph && (glyph+u16(deltas+2*j))%65536) points.add(cp);
      }
    }
  }
  return points;
}
test('bundled licensed fonts contain every Unicode glyph used by the five legacy symbol mappings', () => {
  const fonts = path.join(root,'src/assets/fonts'), covered = new Set();
  for (const source of JSON.parse(fs.readFileSync(path.join(fonts,'SOURCES.json')))) {
    const buffer = fs.readFileSync(path.join(fonts,source.file));
    assert.equal(createHash('sha256').update(buffer).digest('hex'),source.sha256);
    assert.ok(fs.readFileSync(path.join(root,source.licenseFile),'utf8').includes('SIL OPEN FONT LICENSE'));
    for (const cp of codepoints(buffer)) covered.add(cp);
  }
  const mappings = require('../src/symbol-font-maps.json');
  const required = new Set(Object.values(mappings).flatMap(map=>Object.values(map)).flatMap(text=>[...text].map(c=>c.codePointAt(0))));
  assert.equal(required.size,990);
  assert.deepEqual([...required].filter(cp=>!covered.has(cp)),[]);
});
