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
  const nativeRequire = createRequire(sourceFile);
  const helperFile = path.join(__dirname, '../src/hwp/cell-content.ts'), helper = { exports: {} };
  const helperCode = ts.transpileModule(await fs.readFile(helperFile, 'utf8'), { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true } }).outputText;
  new Function('require', 'module', 'exports', helperCode)(nativeRequire, helper, helper.exports);
  new Function('require', 'module', 'exports', compiled)(name => name === './cell-content' ? helper.exports : nativeRequire(name), bridge, bridge.exports);
  return { HwpDocument, ...bridge.exports };
})();
const paper = { size: 'Letter', orientation: 'landscape', margin: 20, margins: { top: 15, right: 18, bottom: 21, left: 24 }, header: '머리말', footer: '꼬리말', pageNumbers: false };
const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';

test('custom physical paper and column spacing survive both Hancom writers and the edit bridge', async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  const page = { ...paper, size: 'Custom', width: 272, height: 394, orientation: 'portrait', columns: { count: 2, space: 18 } };
  for (const format of ['hwp', 'hwpx']) {
    const output = serializeHwp(HwpDocument, '<p>CUSTOM PAPER 한글</p>', page, format);
    const doc = new HwpDocument(output.bytes);
    try {
      const result = editableHwp(doc);
      assert.equal(result.page.size, 'Custom');
      assert.ok(Math.abs(result.page.width - 272) < .02);
      assert.ok(Math.abs(result.page.height - 394) < .02);
      assert.equal(result.page.columns.count, 2);
      assert.ok(Math.abs(result.page.columns.space - 18) < .02);
      assert.match(result.html, /data-column-count="2"/);
      assert.ok(!result.warnings.some(w => /A4로|단일 본문/.test(w)));
    } finally { doc.free(); }
  }
});

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

test('equations retain their SVG appearance and original script before raster preparation', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument(); doc.insertEquation(0, 0, 0, 'a over b', 1100, 0);
    const result = editableHwp(doc);
    assert.match(result.html, /수식: a over b/);
    assert.match(result.html, /data-hwp-svg="&lt;svg/);
    assert.ok(result.warnings.some(warning => warning.includes('수식은 원래 모양')));
  } finally { doc.free(); }
});

test('cell equations preserve multiple controls in one paragraph and nested cells in HWP and HWPX', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const directory = await fs.mkdtemp(path.resolve('.cache/hwp-cell-content-'));
  try {
    const files = await require('./hwp-cell-content-fixture.cjs')(directory);
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(await fs.readFile(files[format]));
      try {
        const converted = editableHwp(doc);
        for (const script of ['a over b', 'x^2', 'y over z', 'sqrt {q}']) assert.ok(converted.html.includes('수식: ' + script), `${format}: ${script}`);
        assert.equal((converted.html.match(/data-hwp-svg=/g) || []).length, 4);
        assert.equal((converted.html.match(/data-hwp-object=/g) || []).length, 3);
        assert.match(converted.html, /BEFORE[\s\S]*수식: x\^2[\s\S]*MIDDLE[\s\S]*수식: y over z[\s\S]*AFTER/);
        assert.match(converted.html, /UNCHANGED 한글/);
        assert.ok(!converted.html.includes('변환하지 못한 개체'), format);
      } finally { doc.free(); }
    }
  } finally { await fs.rm(directory, { recursive: true }); }
});

test('nonvisual cell controls warn without inserting visible text or dropping real unsupported content', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument(); doc.pasteHtml(0, 0, 0, '<table><tr><td>KEEP WORDS</td></tr></table>');
    const original = doc.exportControlHtml.bind(doc);
    doc.exportControlHtml = (...args) => original(...args).replace('KEEP WORDS', 'KEEP WORDS<!-- rhwp: 셀 안 Bookmark omitted --><!-- rhwp: 셀 안 PageHide omitted --><!-- rhwp: 셀 안 Unknown omitted -->');
    // Keep the selection bridge consistent with the deliberately injected
    // unsupported controls without inventing native control indexes.
    const selection = doc.exportSelectionInCellHtmlByPath.bind(doc);
    doc.exportSelectionInCellHtmlByPath = (...args) => selection(...args).replace('KEEP WORDS', 'KEEP WORDS<!-- rhwp: 셀 안 Bookmark omitted --><!-- rhwp: 셀 안 PageHide omitted --><!-- rhwp: 셀 안 Unknown omitted -->');
    const converted = editableHwp(doc);
    assert.match(converted.html, /KEEP WORDS/);
    assert.equal((converted.html.match(/변환하지 못한 개체/g) || []).length, 1);
    assert.ok(converted.warnings.some(w => w.includes('책갈피')));
    assert.ok(converted.warnings.some(w => w.includes('쪽 숨김')));
  } finally { doc.free(); }
});

test('floating shapes keep isolated drawing subtrees and body anchors despite colliding header indexes', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty();
  try {
    require('./hwp-shape-fixture.cjs').populateShapes(source);
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      try {
        const layout = doc.getPageControlLayout.bind(doc);
        doc.getPageControlLayout = page => {
          const body = JSON.parse(layout(page));
          body.controls.push(...body.controls.map(c => ({ ...c, x: 1, y: 1, zOrder: 999, headerFooter: { kind: 'header' } })));
          return JSON.stringify(body);
        };
        const result = editableHwp(doc);
        assert.equal((result.html.match(/data-hwp-object=/g) || []).length, 2, format);
        assert.match(result.html, /TITLE 한글/); assert.match(result.html, /BODY TEXT/);
        assert.match(result.html, /&quot;relative&quot;:&quot;page&quot;/);
        assert.ok(!result.html.includes('원본 보기에서 확인]'));
        assert.ok(result.warnings.some(w => w.includes('개체 내부는 편집할 수 없습니다')));
      } finally { doc.free(); }
    }
  } finally { source.free(); }
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

test('queried HWP paragraph values normalize mixed units without double-scaling margins or fixed leading', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty();
  try {
    source.createBlankDocument(); source.insertText(0, 0, 0, 'HANGING 한글');
    source.applyParaFormat(0, 0, JSON.stringify({ marginLeft: 2400, marginRight: 1200, indent: -1200,
      spacingBefore: 1800, spacingAfter: 600, lineSpacingType: 'Fixed', lineSpacing: 3000 }));
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      try {
        const html = editableHwp(doc).html;
        assert.match(html, /margin-left:24px;margin-right:8px;text-indent:-8px/);
        assert.match(html, /margin-top:12px;margin-bottom:4px/);
        assert.match(html, /line-height:20px/);
        assert.match(html, /HANGING 한글/);
      } finally { doc.free(); }
    }
  } finally { source.free(); }
});

test('imported Hancom tables retain their outer bottom gap instead of the editor table default', async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  const created = serializeHwp(HwpDocument, '<table><tr><td><p>TABLE GAP</p></td></tr></table><p>AFTER</p>', paper, 'hwpx');
  const zip = await JSZip.loadAsync(created.bytes);
  const xml = await zip.file('Contents/section0.xml').async('string');
  const changed = xml.replace(/(<hp:outMargin\b[^>]*\bbottom=")[^"]*(")/, '$1900$2');
  assert.notEqual(changed, xml);
  zip.file('Contents/section0.xml', changed);
  const doc = new HwpDocument(await zip.generateAsync({ type: 'uint8array' }));
  try {
    for (const format of ['hwp', 'hwpx']) {
      const reopened = new HwpDocument(format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
      try {
        const html = editableHwp(reopened).html;
        assert.match(html, /<table[^>]*margin-bottom:12px/);
        assert.match(html, /TABLE GAP/); assert.match(html, /AFTER/);
      } finally { reopened.free(); }
    }
  } finally { doc.free(); }
});

test('zero-leading titles identify a coincident spacer, and Hancom cut-line glyphs retain their half-em advance', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty();
  try {
    require('./hwp-textflow-fixture.cjs').populate(source);
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      try {
        const { html } = editableHwp(doc);
        assert.match(html, /data-hwp-zero-line="53\.3,26\.7"/);
        assert.match(html, /letter-spacing:-6\.666666666666666px/);
        assert.equal((html.match(/┈/g) || []).length, 81);
        assert.match(html, /NORMAL ┈/); assert.match(html, /UPPER TITLE/); assert.match(html, /BIG TITLE/);
      } finally { doc.free(); }
    }
  } finally { source.free(); }
});

test('only observed cell continuations enable row splitting, rather than repeated header occurrences', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty();
  try {
    const c = require('./hwp-row-flow-fixture.cjs').populate(source);
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      try {
        const boxes = JSON.parse(doc.getTableCellBboxes(0, c.paraIdx, c.controlIdx));
        assert.ok(new Set(boxes.filter(b => b.cellIdx === 2).map(b => b.pageIndex)).size > 1);
        const { html } = editableHwp(doc);
        assert.equal((html.match(/data-row-split="allow"/g) || []).length, 1);
        assert.match(html, /<th\b/); assert.match(html, /LEFT 109/); assert.match(html, /RIGHT 74/);
      } finally { doc.free(); }
    }
  } finally { source.free(); }
});

test('inline cell pictures keep their paragraph slot and order instead of adding a second block height', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty();
  try {
    source.createBlankDocument();
    source.pasteHtml(0, 0, 0, `<table><tr><td><p>BEFORE <img src="data:image/png;base64,${png}" width="24" height="24"> AFTER</p></td></tr></table>`);
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      try {
        const html = editableHwp(doc).html, cell = html.match(/<td\b[^>]*>([\s\S]*?)<\/td>/)?.[1];
        assert.ok(cell, format); assert.equal((cell.match(/<img\b/g) || []).length, 1);
        assert.match(cell, /<p\b[^>]*>(?:(?!<\/p>)[\s\S])*BEFORE(?:(?!<\/p>)[\s\S])*<img\b(?:(?!<\/p>)[\s\S])*AFTER(?:(?!<\/p>)[\s\S])*<\/p>/);
      } finally { doc.free(); }
    }
  } finally { source.free(); }
});

test('hard line breaks before cell pictures survive the clipboard exporter omitting the break', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty();
  try {
    source.createBlankDocument();
    source.pasteHtml(0, 0, 0, `<table><tr><td><p>BEFORE<br>LINE <img src="data:image/png;base64,${png}" width="24" height="24"> AFTER<br>END</p></td></tr></table>`);
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      try {
        const html = editableHwp(doc).html;
        assert.match(html, /BEFORE<br>LINE <img\b[^>]*> AFTER<br>END/, format);
        assert.equal((html.match(/<img\b/g) || []).length, 1);
      } finally { doc.free(); }
    }
  } finally { source.free(); }
});
