const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const { readFileSync } = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { createRequire } = require('node:module');
const ts = require('typescript');
const createFixture = require('./hwp-layout-fixture.cjs');
const JSZip = require('jszip');
function loadSource(source) {
  const module = { exports: {} };
  const load = createRequire(source);
  new Function('require', 'module', 'exports', ts.transpileModule(readFileSync(source, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, esModuleInterop: true },
  }).outputText)(name => name.startsWith('.') && !path.extname(name)
    ? loadSource(path.resolve(path.dirname(source), name + '.ts')) : load(name), module, module.exports);
  return module.exports;
}
const setup = loadSource(path.resolve(__dirname, '../src/hwp/layout-metadata.ts'));
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

test('tracking is an explicit conversion restriction even if optional layout parsing fails; other FileHeader bits do not trigger it', async () => {
  const CFB = require('cfb'), { readHwpLayoutMetadata } = await setup;
  for (const flags of [1, 0x41, 0x4001]) {
    const cfb = CFB.utils.cfb_new(), header = Buffer.alloc(256);
    header.writeUInt32LE(flags, 36);
    CFB.utils.cfb_add(cfb, 'FileHeader', header);
    CFB.utils.cfb_add(cfb, 'BodyText/Section0', Buffer.from([1, 2, 3]));
    const result = await readHwpLayoutMetadata(CFB.write(cfb, { type: 'buffer' }));
    assert.equal(!!result.conversionBlock, flags === 0x4001);
    if (result.conversionBlock) assert.match(result.conversionBlock, /변경 내용 추적/);
    assert.equal(result.warnings.length, 1, 'optional layout failure must not erase the restriction');
  }
});


test('tracking flags use the case-insensitive root FileHeader, not a nested decoy', async () => {
  const CFB = require('cfb'), { readHwpLayoutMetadata } = await setup;
  for (const name of ['FileHeader', 'fileheader', 'FILEHEADER']) {
    const cfb = CFB.utils.cfb_new(), tracked = Buffer.alloc(256), ordinary = Buffer.alloc(256);
    tracked.writeUInt32LE(0x4000, 36);
    CFB.utils.cfb_add(cfb, 'Preview/FileHeader', ordinary);
    CFB.utils.cfb_add(cfb, '/' + name, tracked);
    const result = await readHwpLayoutMetadata(CFB.write(cfb, { type: 'buffer' }));
    assert.match(result.conversionBlock, /변경 내용 추적/);
  }
});

test('ordinary space policy follows the root HWP character records even with compressed DocInfo', async () => {
  const CFB = require('cfb'), { deflateRawSync } = require('node:zlib');
  const { readHwpLayoutMetadata } = await setup;
  const shape = value => {
    const record = Buffer.alloc(54);
    record.writeUInt32LE(21 | (50 << 20));
    record.writeUInt32LE(value ? 1 << 25 : 0, 50);
    return record;
  };
  const records = Buffer.concat([shape(false), shape(true)]);
  for (const compressed of [false, true]) {
    const cfb = CFB.utils.cfb_new(), header = Buffer.alloc(256);
    if (compressed) header[36] = 1;
    CFB.utils.cfb_add(cfb, 'FileHeader', header);
    CFB.utils.cfb_add(cfb, 'Preview/DocInfo', shape(true));
    CFB.utils.cfb_add(cfb, '/docinfo', compressed ? deflateRawSync(records) : records);
    const result = await readHwpLayoutMetadata(CFB.write(cfb, {type:'buffer'}));
    assert.deepEqual(result.fontSpace, [false, true]);
    assert.deepEqual(result.warnings, []);
  }
});

test('HWPX uses declared character-space flags only at verified shape indices', async () => {
  const { readHwpLayoutMetadata } = await setup;
  const zip = new JSZip();
  zip.file('Contents/content.hpf', '<p:package><p:item id="header" href="./styles/default.xml" media-type="application/xml"/><p:item id="s" href="section0.xml" media-type="application/xml"/><p:itemref idref="s"/></p:package>');
  zip.file('Contents/section0.xml', '<s:sec><p:p pageBreak="1"/></s:sec>');
  zip.file('Contents/header.xml', '<h:head><h:charPr id="0" useFontSpace="1"/></h:head>');
  const first = '<h:charPr id="0" useFontSpace="0"/>';
  const second = '<h:charPr useFontSpace="true" id="1"/>';
  const variants = [
    { xml: first + '<!-- <h:charPr id="1" useFontSpace="0"/> -->' + second, flags: [false, true] },
    { xml: first + '<h:charPr id="0" useFontSpace="1"/>' },
    { xml: first + '<h:charPr id="2" useFontSpace="0"/>' },
    { xml: first + '<h:charPr id="1"/>' },
  ];
  for (const { xml, flags } of variants) {
    zip.file('Contents/styles/default.xml', '<h:head>' + xml + '</h:head>');
    const result = await readHwpLayoutMetadata(await zip.generateAsync({ type: 'uint8array' }), true);
    assert.deepEqual(result.hwpxCellFontSpace, flags);
    assert.equal(result.fontSpace, undefined);
    assert.equal(result.paragraphs['0:0'].pageBreak, true);
    assert.deepEqual(result.warnings, []);
  }
});

test('HWPX retains local picture fills and authored brightness beside invalid optional references', async () => {
  const { readHwpLayoutMetadata } = await setup;
  const zip = new JSZip();
  const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+s1gUAAAAASUVORK5CYII=';
  zip.file('Contents/content.hpf', '<p:package>' +
    '<p:item id="header" href="header.xml" media-type="application/xml"/>' +
    '<p:item id="s" href="section0.xml" media-type="application/xml"/>' +
    '<p:item id="local" href="../BinData/brush.png" media-type="image/png"/>' +
    '<p:item id="remote" href="https://example.invalid/brush.png" media-type="image/png"/>' +
    '<p:itemref idref="s"/></p:package>');
  zip.file('Contents/section0.xml', '<s:sec><p:p pageBreak="1"/></s:sec>');
  zip.file('BinData/brush.png', png, { base64: true });
  const brush = (id, image, alpha, bright = 0) => `<h:borderFill id="${id}"><c:fillBrush><c:imgBrush mode="TOTAL">` +
    `<c:img binaryItemIDRef="${image}" bright="${bright}" contrast="0" effect="REAL_PIC" alpha="${alpha}"/>` +
    '</c:imgBrush></c:fillBrush></h:borderFill>';
  zip.file('Contents/header.xml', '<h:head>' + brush(1, 'local', 0) + brush(2, 'remote', 0) + brush(3, 'local', 20) +
    brush(4, 'local', 0, 8) + brush(5, 'local', 0, 101) + brush(6, 'local', 0, -8) + '</h:head>');
  const result = await readHwpLayoutMetadata(await zip.generateAsync({ type: 'uint8array' }), true);
  const source = 'data:image/png;base64,' + png;
  assert.deepEqual(result.hwpxCellRasterFills, {
    1: { source, brightness: 0 }, 4: { source, brightness: 8 }, 6: { source, brightness: -8 },
  });
  assert.equal(result.paragraphs['0:0'].pageBreak, true);
  assert.deepEqual(result.warnings, []);
});
