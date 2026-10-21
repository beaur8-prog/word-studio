'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { createRequire } = require('node:module');
const ts = require('typescript');
const JSZip = require('jszip');
const setup = (async () => {
  const directory = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(directory, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(directory, 'rhwp_bg.wasm')) });
  // Exercise the same DOM-free conversion module used by the browser Worker.
  const sourceFile = path.join(__dirname, '../src/hwp/conversion.ts');
  const source = await fs.readFile(sourceFile, 'utf8');
  const compiled = ts.transpileModule(source, { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true } }).outputText;
  const bridge = { exports: {} };
  new Function('require', 'module', 'exports', compiled)(createRequire(sourceFile), bridge, bridge.exports);
  return { HwpDocument, ...bridge.exports };
})();
const paper = { size: 'Letter', orientation: 'landscape', margin: 20, margins: { top: 15, right: 18, bottom: 21, left: 24 }, header: '머리말', footer: '꼬리말', pageNumbers: false };
const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';

test('measured conversion progress reaches the actual paragraph count without changing document content', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const doc = new HwpDocument(await fs.readFile(path.join(__dirname, 'fixtures/hwp/reader.hwp')));
  try {
    const baseline = editableHwp(doc), events = [];
    const converted = editableHwp(doc, progress => events.push(progress));
    assert.deepEqual(converted, baseline);
    const counts = events.filter(event => event.total !== undefined);
    const total = Array.from({ length: doc.getSectionCount() }, (_, i) => doc.getParagraphCount(i)).reduce((a, b) => a + b, 0);
    assert.equal(counts[0].completed, 0);
    assert.equal(counts.at(-1).completed, total);
    assert.ok(counts.every(event => event.total === total && event.completed >= 0 && event.completed <= total));
    assert.ok(events.some(event => event.total === undefined && /표/.test(event.label)));
  } finally { doc.free(); }
});

for (const format of ['hwp', 'hwpx']) {
  test(`${format}: edit bridge writes real Hancom files with Unicode, formatting, merged/nested tables, image and page settings`, async () => {
    const { HwpDocument, serializeHwp, editableHwp } = await setup;
    const html = `<p style="text-align:right;line-height:150%"><strong><span style="font-size:16pt;color:#123456">한글 😀 EDITED</span></strong></p><table><tr><td colspan="2"><p>병합 셀</p></td></tr><tr><td>A</td><td><table><tr><td>NESTED</td></tr></table></td></tr></table><p><img src="data:image/png;base64,${png}" width="24" height="24"></p><div data-page-break="true"></div><p>마지막 쪽</p>`;
    const output = serializeHwp(HwpDocument, html, paper, format);
    if (format === 'hwp') assert.equal(Buffer.from(output.bytes).subarray(0, 8).toString('hex'), 'd0cf11e0a1b11ae1');
    else { const zip = await JSZip.loadAsync(output.bytes); assert.ok(zip.file('Contents/section0.xml')); }
    const doc = new HwpDocument(output.bytes);
    try {
      assert.equal(doc.pageCount(), 2);
      const content = editableHwp(doc);
      for (const text of ['한글 😀 EDITED', '병합 셀', 'NESTED', '마지막 쪽']) assert.ok(content.html.includes(text), text);
      assert.equal((content.html.match(/<table/g) || []).length, 2);
      assert.match(content.html, /colspan="2"/); assert.match(content.html, /data:image\/png;base64/);
      assert.match(content.html, /font-weight:bold/); assert.match(content.html, /font-size:16.0pt/);
      assert.match(content.html, /text-align:right/); assert.match(content.html, /data-page-break/);
      assert.equal(content.page.size, 'Letter'); assert.equal(content.page.orientation, 'landscape');
      assert.ok(Math.abs(content.page.margins.left - 24) < .02);
      assert.ok(Math.abs(content.page.margins.top - 15) < .02);
      assert.ok(Math.abs(content.page.margins.bottom - 21) < .02);
      assert.equal(content.page.header, '머리말'); assert.equal(content.page.footer, '꼬리말');
      assert.deepEqual(output.warnings, []);
    } finally { doc.free(); }
  });
}

test('HWP and standard OPF-relative HWPX import all table cells rather than only selection text', async () => {
  const { HwpDocument, editableHwp, serializeHwp } = await setup;
  for (const name of ['reader.hwp', 'reader.hwpx', 'manifest-relative.hwpx', 'landscape-pagination.hwpx']) {
    const doc = new HwpDocument(await fs.readFile(path.join(__dirname, 'fixtures/hwp', name)));
    try {
      const source = editableHwp(doc);
      if (name.startsWith('reader')) assert.match(source.html, /테이블 속 한글/);
      if (name.startsWith('manifest')) { assert.match(source.html, /첫 번째 문단/); assert.match(source.html, /두 번째 문단/); }
      if (name.startsWith('landscape')) for (let row = 1; row <= 24; row++) for (let col = 1; col <= 2; col++) {
        const label = `ROW-${String(row).padStart(2, '0')}-COLUMN-${col}`;
        assert.equal(source.html.split(label).length - 1, 1, label);
      }
      const out = serializeHwp(HwpDocument, source.html.replace('테이블 속 한글', '수정한 표 내용'), source.page, 'hwp');
      const reopen = new HwpDocument(out.bytes);
      try { if (name.startsWith('reader')) assert.match(editableHwp(reopen).html, /수정한 표 내용/); }
      finally { reopen.free(); }
    } finally { doc.free(); }
  }
});

for (const fixture of [
  { name: 'regular', heights: [32, 32, 32, 32] },
  { name: 'merged', heights: [32, 32, 32, 32], merged: true },
  { name: 'unequal', heights: [24, 40, 28, 36] },
]) {
  test(`Hancom ${fixture.name} rows survive a 72pt outer table with 128pt of cells`, async () => {
    const { HwpDocument, serializeHwp, editableHwp } = await setup;
    const cell = (row, column, span = 1) => `<td rowspan="${span}" style="width:180px;height:${fixture.heights.slice(row, row + span).reduce((a, b) => a + b, 0) * 4 / 3}px"><p>표 ${row + 1}-${column + 1}</p></td>`;
    const html = '<table>' + fixture.heights.map((_, row) => `<tr>${fixture.merged && row === 1 ? '' : cell(row, 0, fixture.merged && row === 0 ? 2 : 1)}${cell(row, 1)}</tr>`).join('') + '</table>';
    const output = serializeHwp(HwpDocument, html, paper, 'hwpx');
    const zip = await JSZip.loadAsync(output.bytes);
    const xml = await zip.file('Contents/section0.xml').async('string');
    // 100 HWPUNIT/pt. Alter only the outer size, preserving every cell size.
    const short = xml.replace(/(<hp:tbl\b[^>]*><hp:sz\b[^>]*\bheight=")\d+("[\s\S]*?\/?>)/, (_, before, after) => before + '7200' + after);
    assert.notEqual(short, xml);
    zip.file('Contents/section0.xml', short);
    const source = new HwpDocument(await zip.generateAsync({ type: 'uint8array' }));
    try {
      const control = JSON.parse(source.getControls()).find(c => c.ctrlId === 'tbl');
      assert.equal(control.props.Height, 7200);
      for (const format of ['hwpx', 'hwp']) {
        const doc = format === 'hwpx' ? source : new HwpDocument(source.exportHwp());
        try {
          const converted = editableHwp(doc);
          const heights = [...converted.html.matchAll(/<tr style="height:([\d.]+)px;/g)].map(match => Number(match[1]));
          assert.equal(heights.length, 4, format);
          heights.forEach((height, row) => assert.ok(Math.abs(height - fixture.heights[row] * 4 / 3) < .02, `${format} row ${row}: ${height}`));
          assert.ok(heights.reduce((a, b) => a + b, 0) >= 128 * 4 / 3 - .02);
          assert.match(converted.html, /표 4-2/);
          if (fixture.merged) assert.match(converted.html, /rowspan="2"/);
        } finally { if (doc !== source) doc.free(); }
      }
    } finally { source.free(); }
  });
}

test('unsupported equations remain visible as text and are reported before conversion', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument(); doc.insertEquation(0, 0, 0, 'a over b', 1100, 0);
    const result = editableHwp(doc);
    assert.match(result.html, /수식: a over b/);
    assert.ok(result.warnings.some(warning => warning.includes('수식 개체')));
  } finally { doc.free(); }
});

test('single-paragraph style survives repeated saves without accumulating sentinel paragraphs', async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  let html = '<p style="text-align:center;line-height:150%"><b>SINGLE 문단</b></p>';
  for (let i = 0; i < 3; i++) {
    const { bytes } = serializeHwp(HwpDocument, html, { ...paper, header: '', footer: '' }, 'hwp');
    const doc = new HwpDocument(bytes);
    try { assert.equal(doc.getParagraphCount(0), 1); assert.equal(JSON.parse(doc.getParaPropertiesAt(0, 0)).alignment, 'center'); html = editableHwp(doc).html; }
    finally { doc.free(); }
  }
});

test('conversion refuses oversized input and invalid format', async () => {
  const { HwpDocument, serializeHwp, HWP_HTML_LIMIT, checkHwpHtmlSize, hwpHtmlBytes } = await setup;
  assert.doesNotThrow(() => checkHwpHtmlSize(HWP_HTML_LIMIT));
  assert.throws(() => checkHwpHtmlSize(HWP_HTML_LIMIT + 1), /변환 데이터.*256 MiB.*원본 파일 크기가 아니라/);
  for (const value of ['ASCII', '한글 문서 😀', '\ud800', '\udc00', 'a😀b가']) assert.equal(hwpHtmlBytes(value), Buffer.byteLength(value));
  assert.throws(() => serializeHwp(HwpDocument, '가'.repeat(Math.floor(HWP_HTML_LIMIT / 3) + 1), paper, 'hwpx'), /변환 데이터.*256 MiB/);
  assert.throws(() => serializeHwp(HwpDocument, '<p>x</p>', paper, 'exe'), /형식/);
});
