const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { createRequire } = require('node:module');
const ts = require('typescript');
const createFixture = require('./hwp-layout-fixture.cjs');
const JSZip = require('jszip');
const setup = (async () => {
  const source = path.resolve(__dirname, '../src/hwp/layout-metadata.ts'), module = { exports: {} };
  new Function('require', 'module', 'exports', ts.transpileModule(await fs.readFile(source, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, esModuleInterop: true },
  }).outputText)(createRequire(source), module, module.exports);
  return module.exports;
})();
test('original HWP and HWPX paragraph headers retain explicit column/page breaks and physical column gaps', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-layout-'));
  try {
    const files = await createFixture(dir), { readHwpLayoutMetadata } = await setup;
    for (const format of ['hwp', 'hwpx']) {
      const result = await readHwpLayoutMetadata(await fs.readFile(files[format]), format === 'hwpx');
      assert.deepEqual(result.warnings, [], format);
      assert.equal(Object.keys(result.paragraphs).length, 3, format);
      assert.deepEqual(result.paragraphs['0:0'].columns, { count: 2, space: 24 }, format);
      assert.equal(result.paragraphs['0:1'].columnBreak, true, format);
      assert.equal(result.paragraphs['0:2'].pageBreak, true, format);
    }
  } finally { await fs.rm(dir, { recursive: true, force: true }); }
});
test('nested paragraphs and fake XML inside comments do not change the main paragraph index', async () => {
  const { parseHwpxSection } = await setup;
  const result = parseHwpxSection('<s:sec><p:p pageBreak="0"><p:colPr colCount="2" sameGap="1500"/><p:p pageBreak="1"/></p:p><!-- <p:p pageBreak="1"/> --><p:p columnBreak="1"/></s:sec>', 0);
  assert.deepEqual(Object.keys(result), ['0:0', '0:1']);
  assert.equal(result['0:0'].pageBreak, false); assert.equal(result['0:1'].columnBreak, true);
});
test('truncated binary metadata fails safely, without guessing paragraph indices', async () => {
  const { parseHwpSection, readHwpLayoutMetadata } = await setup;
  assert.throws(() => parseHwpSection(new Uint8Array([66, 0, 16, 0]), 0), /Truncated/);
  const fallback = await readHwpLayoutMetadata(new Uint8Array([1, 2, 3]));
  assert.deepEqual(fallback.paragraphs, {}); assert.equal(fallback.warnings.length, 1);
});

test('HWPX layout follows OPF-relative subdirectories and spine order, ignoring root decoys', async () => {
  const { readHwpLayoutMetadata } = await setup;
  const zip = new JSZip();
  zip.file('Contents/parts/section-one.xml', '<s:sec><p:p columnBreak="1"><p:colPr colCount="2" sameGap="1800"/></p:p><p:p pageBreak="1"/></s:sec>');
  zip.file('Contents/section9.xml', '<s:sec><p:p/><p:p columnBreak="1"/></s:sec>');
  zip.file('section9.xml', '<s:sec><p:p pageBreak="1"/></s:sec>');
  zip.file('Contents/content.hpf', '<opf:package><opf:manifest><opf:item id="nine" href="section9.xml" media-type="application/xml"/><opf:item id="one" href="./parts/x/../section-one.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="one"/><opf:itemref idref="nine"/></opf:spine></opf:package>');
  let result = await readHwpLayoutMetadata(await zip.generateAsync({ type: 'uint8array' }), true);
  assert.deepEqual(result.warnings, []);
  assert.deepEqual(Object.keys(result.paragraphs), ['0:0', '0:1', '1:0', '1:1']);
  assert.deepEqual(result.paragraphs['0:0'], { columnBreak: true, pageBreak: false, columns: { count: 2, space: 24 } });
  assert.equal(result.paragraphs['0:1'].pageBreak, true);
  assert.equal(result.paragraphs['1:0'].pageBreak, false);
  assert.equal(result.paragraphs['1:1'].columnBreak, true);
  // Legacy packages use archive-root hrefs; resolve those only if the
  // OPF-relative candidate does not exist.
  zip.file('Contents/content.hpf', (await zip.file('Contents/content.hpf').async('string')).replace('href="section9.xml"', 'href="Contents/section9.xml"'));
  assert.deepEqual(await readHwpLayoutMetadata(await zip.generateAsync({ type: 'uint8array' }), true), result);
});

test('HWPX metadata never silently succeeds with missing sections or invalid package paths', async () => {
  const { readHwpLayoutMetadata } = await setup;
  for (const href of ['missing-section.xml', '../../section.xml', '/section.xml', 'https://example.com/section.xml', 'C:\\section.xml']) {
    const zip = new JSZip();
    zip.file('Contents/content.hpf', `<p:package><p:item id="s" href="${href}" media-type="application/xml"/><p:itemref idref="s"/></p:package>`);
    const result = await readHwpLayoutMetadata(await zip.generateAsync({ type: 'uint8array' }), true);
    assert.deepEqual(result.paragraphs, {}, href); assert.equal(result.warnings.length, 1, href);
  }
});

test('HWP and HWPX preserve signed line leading without reading nested cell lines as body lines', async () => {
  const { parseHwpSection, parseHwpxSection } = await setup;
  const record = (tag, level, bytes) => {
    const b = Buffer.alloc(4 + bytes.length); b.writeUInt32LE(tag | (level << 10) | (bytes.length << 20)); bytes.copy(b, 4); return b;
  };
  const lines = Buffer.alloc(72); lines.writeInt32LE(900, 20); lines.writeInt32LE(-150, 56);
  const nested = Buffer.alloc(36); nested.writeInt32LE(7500, 20);
  const binary = parseHwpSection(Buffer.concat([record(66, 0, Buffer.alloc(12)), record(69, 1, lines), record(69, 3, nested)]), 0);
  assert.deepEqual(binary['0:0'].lineSpacings, [12, -2]);
  const xml = parseHwpxSection('<hp:sec><hp:p><hp:p><hp:lineseg spacing="7500"/></hp:p><hp:lineseg spacing="900"/><hp:lineseg spacing="-150"/></hp:p></hp:sec>', 0);
  assert.deepEqual(xml['0:0'].lineSpacings, [12, -2]);
  assert.throws(() => parseHwpSection(Buffer.concat([record(66, 0, Buffer.alloc(12)), record(69, 1, Buffer.alloc(37))]), 0), /line segments/);
});
