const { test } = require('node:test'), assert = require('node:assert/strict'), fixture = require('./hwp-inline-wrap-fixture.cjs');
let pending; const samples = () => pending ??= fixture();
for (const format of ['hwp', 'hwpx']) test(`${format}: narrow leading inline tables preserve multiline wrapping without inheriting the object's line height`, async () => {
  const source = (await samples())[format], html = source.editable.html;
  assert.ok(source.lines.lineCount > 1, 'the native paragraph spans more than one visual line');
  assert.equal((html.match(/<table\b/g) || []).length, 1, 'no additional two-column layout table');
  assert.match(html, /data-table-anchor="[^"]*&quot;wrap&quot;:&quot;right&quot;/, 'the source leading inline label has a live wrapping anchor');
  const p = [...html.matchAll(/<p\b[^>]*>[^]*?<\/p>/g)].find(m => m[0].includes('TAIL START'))?.[0];
  assert.ok(p, 'the complete tail remains a single editable paragraph');
  const leading = source.char.fontSize / 75 * source.properties.lineSpacing / 100;
  assert.match(p, new RegExp('line-height:' + leading + 'px'), 'the text-only fallback excludes the removed inline table height');
  assert.equal(p.replace(/<[^>]*>/g, ''), source.tail);
});
