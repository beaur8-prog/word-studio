'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');
const { createRequire } = require('node:module');
const ts = require('typescript');
const JSZip = require('jszip');
const setup = (async () => {
  const directory = process.env.WORD_STUDIO_TEST_HWP_CORE || path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(directory, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(directory, 'rhwp_bg.wasm')) });
  // Exercise the same DOM-free conversion module used by the browser Worker.
  const sourceFile = path.join(__dirname, '../src/hwp/conversion.ts');
  const cache = new Map();
  const load = file => {
    if (cache.has(file)) return cache.get(file).exports;
    const module = { exports: {} }; cache.set(file, module);
    const nativeRequire = createRequire(file);
    const code = ts.transpileModule(require('node:fs').readFileSync(file, 'utf8'), { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true } }).outputText;
    new Function('require', 'module', 'exports', code)(name => name.startsWith('.') && !path.extname(name) ? load(path.resolve(path.dirname(file), name + '.ts')) : nativeRequire(name), module, module.exports);
    return module.exports;
  };
  return { HwpDocument, ...load(sourceFile), ...load(path.join(path.dirname(sourceFile), 'paragraph-layout.ts')),
    ...load(path.join(path.dirname(sourceFile), 'export-grid-rows.ts')),
    ...load(path.join(path.dirname(sourceFile), 'script-fonts.ts')),
    ...load(path.join(path.dirname(sourceFile), '../hwp-text-advances.ts')) };
})();
const paper = { size: 'Letter', orientation: 'landscape', margin: 20, margins: { top: 15, right: 18, bottom: 21, left: 24 }, header: '머리말', footer: '꼬리말', pageNumbers: false };
const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';

test('inline picture alternative text retains literal XML and replacement characters in both native writers', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  for (const format of ['hwp', 'hwpx']) {
    const html = `<p><img src="data:image/png;base64,${png}" width="20" height="20" alt="Picture $&amp; &lt;tag&gt; &quot;한글&quot;"></p>`;
    const out = serializeHwp(HwpDocument, html, { ...paper, header: '', footer: '' }, format), doc = new HwpDocument(out.bytes);
    try {
      const image = JSON.parse(doc.getControls()).find(c => c.ctrlId === 'gso');
      assert.ok(image);
      assert.equal(JSON.parse(doc.getPictureProperties(0, image.para, image.controlIndex)).description, 'Picture $& <tag> "한글"');
    } finally { doc.free(); }
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: floating pictures keep paper coordinates, wrapping, margins and alternative text through two saves`, async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  let html = [0, 1, 3].map((type, i) => {
    const g = { horizontalPosition: { relative: 'page', offset: (120 + 100 * i) * 9525 }, verticalPosition: { relative: 'page', offset: 160 * 9525 }, wrap: { type }, behindDocument: type === 0,
      margins: { left: 3 * 9525, right: 4 * 9525, top: 5 * 9525, bottom: 6 * 9525 } };
    return `<p>BEFORE${i}<img data-image-floating="${JSON.stringify(g).replaceAll('"', '&quot;')}" src="data:image/png;base64,${png}" width="80" height="40" alt="그림 &amp; ${i}">AFTER${i}</p>`;
  }).join('');
  for (let round = 0; round < 2; round++) {
    const out = serializeHwp(HwpDocument, html, { ...paper, header: '', footer: '' }, format), doc = new HwpDocument(out.bytes);
    try {
      const controls = JSON.parse(doc.getPageControlLayout(0)).controls.filter(c => c.type === 'image');
      assert.equal(controls.length, 3);
      for (const [i, c] of controls.entries()) {
        assert.ok(Math.abs(c.x - (120 + 100 * i)) < .1, JSON.stringify(c)); assert.ok(Math.abs(c.y - 160) < .1, JSON.stringify(c));
        assert.equal(c.w, 80); assert.equal(c.h, 40);
        const props = JSON.parse(doc.getPictureProperties(c.secIdx, c.paraIdx, c.controlIdx));
        assert.equal(props.treatAsChar, false); assert.equal(props.description, `그림 & ${i}`);
        assert.equal(props.outerMarginLeft, 225); assert.equal(props.outerMarginBottom, 450);
        assert.equal(props.textWrap, ['BehindText', 'Square', 'TopAndBottom'][i]);
      }
      const imported = editableHwp(doc); html = imported.html;
      assert.equal((html.match(/data-image-floating=/g) || []).length, 3); assert.doesNotMatch(html, /WORDSTUDIO-PICTURE/);
      for (let i = 0; i < 3; i++) { assert.ok(html.includes(`BEFORE${i}`)); assert.ok(html.includes(`AFTER${i}`)); }
    } finally { doc.free(); }
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: paper-relative cell pictures use page coordinates in native layout and retain their metadata`, async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  for (const nested of [false, true]) for (const type of [0, 1, 3]) {
    const geometry = { horizontalPosition: { relative: 'page', offset: 200 * 9525 }, verticalPosition: { relative: 'page', offset: 160 * 9525 }, wrap: { type }, behindDocument: type === 0,
      margins: { left: 3 * 9525, right: 4 * 9525, top: 5 * 9525, bottom: 6 * 9525 } };
    const picture = `<p>BEFORE<img data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}" src="data:image/png;base64,${png}" width="80" height="40" alt="Cell picture">AFTER</p>`;
    let html = `<table style="width:500px"><tr style="height:180px"><td>${nested ? `<p>OUTER</p><table style="width:350px"><tr><td>${picture}</td></tr></table>` : picture}</td></tr></table>`;
    for (let round = 0; round < 2; round++) {
      const out = serializeHwp(HwpDocument, html, { ...paper, header: '', footer: '' }, format), doc = new HwpDocument(out.bytes);
      try {
        const pictures = JSON.parse(doc.getPageControlLayout(0)).controls.filter(c => c.type === 'image');
        assert.equal(pictures.length, 1);
        const c = pictures[0];
        assert.equal(c.cellPath.length, nested ? 2 : 1, JSON.stringify({nested,type,round,c}));
        assert.ok(Math.abs(c.x - 200) < .1, JSON.stringify(c)); assert.ok(Math.abs(c.y - 160) < .1, JSON.stringify(c));
        const props = JSON.parse(doc.getCellPicturePropertiesByPath(c.secIdx, c.paraIdx, JSON.stringify(c.cellPath), c.controlIdx));
        assert.equal(props.treatAsChar, false); assert.equal(props.description, 'Cell picture');
        assert.equal(props.textWrap, ['BehindText', 'Square', '', 'TopAndBottom'][type]);
        assert.equal(props.outerMarginLeft, 225); assert.equal(props.outerMarginBottom, 450);
        const converted = editableHwp(doc); html = converted.html;
        assert.equal((html.match(/data-image-floating=/g) || []).length, 1);
        assert.match(html, /alt="Cell picture"/); assert.ok(html.includes('BEFORE')); assert.ok(html.includes('AFTER'));
      } finally { doc.free(); }
    }
  }
});

test('a standalone centered footer table stays centered in native layout after two saves', async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  let page = { ...paper, size: 'A4', orientation: 'portrait', margins: { top: 25, right: 20, bottom: 30, left: 20 }, footerDistance: 10,
    header: '', footer: '', footerHtml: '<table style="width:90px;margin-left:auto;margin-right:auto"><tr><td style="width:90px;padding:0"><p style="font-size:9pt;line-height:100%;margin:0"><span data-page-number="PAGE">9</span></p></td></tr></table>' };
  for (const format of ['hwp', 'hwpx']) for (let round = 0; round < 2; round++) {
    const out = serializeHwp(HwpDocument, '<p>BODY</p>', page, format), doc = new HwpDocument(out.bytes);
    try {
      const def = JSON.parse(doc.getPageDef(0)), table = JSON.parse(doc.getPageControlLayout(0)).controls.find(c => c.type === 'table' && c.headerFooter?.kind === 'footer');
      assert.ok(table); assert.ok(Math.abs(table.x + table.w / 2 - def.width / 150) < .2);
      page = editableHwp(doc).page; assert.match(page.footerHtml, /margin-left:auto;margin-right:auto/);
    } finally { doc.free(); }
  }
});

test('Hancom furniture uses paper-edge distances and bottom-aligns the footer in its reserved area', async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  const page = { ...paper, size: 'A4', orientation: 'portrait', margin: 30, margins: { top: 30, right: 20, bottom: 35, left: 20 }, headerDistance: 5, footerDistance: 7,
    headerHtml: '<p style="font-size:9pt;line-height:100%;margin:0">HEADER</p>', footerHtml: '<p style="font-size:9pt;line-height:100%;margin:0">FOOTER</p>' };


for (const format of ['hwp', 'hwpx']) {
    const out = serializeHwp(HwpDocument, '<p>BODY</p>', page, format), doc = new HwpDocument(out.bytes);
    try {
      const def = JSON.parse(doc.getPageDef(0)), converted = editableHwp(doc);
      assert.ok(Math.abs(converted.page.headerDistance - 5) < .02); assert.ok(Math.abs(converted.page.footerDistance - 7) < .02);
      assert.ok(Math.abs(def.marginTop * 25.4 / 7200 - 5) < .02); assert.ok(Math.abs(def.marginBottom * 25.4 / 7200 - 7) < .02);
      const tree = JSON.parse(doc.getPageLayerTreeWithProfile(0, 'print', true, true));
      const find = (node, kind) => node.groupKind?.kind === kind ? node : (node.children || []).map(n => find(n, kind)).find(Boolean) || (node.child && find(node.child, kind));
      const header = find(tree.root, 'header'), footer = find(tree.root, 'footer');
      assert.ok(header && footer);
      assert.ok(Math.abs(header.children[0].bounds.y - 5 * 96 / 25.4) < .03);
      const bottom = Math.max(...footer.children.map(n => n.bounds.y + n.bounds.height));
      assert.ok(Math.abs(bottom - (def.height / 75 - 7 * 96 / 25.4)) < .03, `Footer bottom ${bottom}`);
    } finally { doc.free(); }
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: rich odd/even furniture retains native formatted stories, images, tables and live fields`, async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  let page = { ...paper, differentOddEven: true, pageNumbers: true,
    headerHtml: `<p><b>ODD HEADER</b> <span data-page-number="PAGE">99</span>/<span data-page-number="NUMPAGES">99</span></p><table><tr><td><p>HEADER CELL</p></td></tr></table><p><img src="data:image/png;base64,${png}" width="12" height="12"></p>`,
    evenHeaderHtml: '<p><i>EVEN HEADER</i></p>', footerHtml: '<p>ODD FOOTER</p>', evenFooterHtml: '<p>EVEN FOOTER</p>', evenPageNumbers: true };
  let html = '<p>BODY ONE</p><div data-page-break="true"></div><p>BODY TWO</p>';
  for (let round = 0; round < 2; round++) {
  const out = serializeHwp(HwpDocument, html, page, format), doc = new HwpDocument(out.bytes);
  try {
    assert.equal(doc.getParagraphCount(0), 2); assert.equal(doc.pageCount(), 2);
    const zip = await JSZip.loadAsync(doc.exportHwpx()), xml = await zip.file('Contents/section0.xml').async('string');
    const headers = [...xml.matchAll(/<hp:header\b[^>]*>[\s\S]*?<\/hp:header>/g)].map(m => m[0]);
    assert.equal(headers.length, 2); assert.match(headers.find(h => h.includes('applyPageType="ODD"')), /HEADER CELL/);
    assert.match(headers.join(''), /<hp:pic/); assert.match(headers.join(''), /<hp:tbl/);
    assert.match(headers.join(''), /numType="PAGE"/); assert.match(headers.join(''), /numType="TOTAL_PAGE"/);
    assert.match(xml, /applyPageType="EVEN"[\s\S]*?EVEN FOOTER[\s\S]*?numType="PAGE"/);
    assert.doesNotMatch(xml, /WORDSTUDIO-FURNITURE|WORDSTUDIO-FIELD|<hp:t>99<\/hp:t>/);
    const info = await zip.file('Contents/header.xml').async('string'); assert.match(info, /<hh:bold/); assert.match(info, /<hh:italic/);
    for (const text of ['ODD HEADER', 'EVEN HEADER', 'ODD FOOTER', 'EVEN FOOTER', 'BODY ONE', 'BODY TWO']) assert.equal(xml.split(`<hp:t>${text}</hp:t>`).length - 1, 1);
    const converted = editableHwp(doc); html = converted.html; page = converted.page;
    assert.match(page.headerHtml, /font-weight:bold/); assert.match(page.evenHeaderHtml, /font-style:italic/);
    assert.match(page.headerHtml, /data-page-number="PAGE"/); assert.match(page.headerHtml, /data-page-number="NUMPAGES"/);
    assert.match(page.headerHtml, /<table/); assert.match(page.headerHtml, /<img/);
  } finally { doc.free(); }
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: sections keep paper, margins, headers and shared notes/images through two saves`, async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  const second = { ...paper, size: 'Custom', width: 240, height: 320, orientation: 'portrait', margins: { top: 30, right: 17, bottom: 25, left: 19 }, header: 'SECOND HEADER', footer: 'SECOND FOOTER', columns: { count: 2, space: 24, balance: false } };
  const marker = `<div data-section-page="${JSON.stringify(second).replaceAll('"', '&quot;')}"></div>`;
  let html = `<p><b>FIRST 한글</b></p>${marker}<div data-column-count="2" data-column-space="24" data-column-balance="false"><p>SECOND <span data-note-reference="n1" data-note-kind="footnote" data-note-label="1">1</span></p><p><img src="data:image/png;base64,${png}" width="24" height="24"></p><p>PAGE <span data-page-number="PAGE">1</span></p></div><div data-document-note="n1" data-note-kind="footnote" data-note-label="1"><p>NOTE BODY</p></div>`;
  let first = paper;
  for (let round = 0; round < 2; round++) {
    const out = serializeHwp(HwpDocument, html, first, format), doc = new HwpDocument(out.bytes);
    try {
      assert.equal(doc.getSectionCount(), 2);
      assert.equal(JSON.parse(doc.getColumnDef(1)).columnCount, 2);
      assert.equal(JSON.parse(doc.getParaPropertiesAt(1, 0)).pageBreakBefore, false);
      const a = JSON.parse(doc.getPageDef(0)), b = JSON.parse(doc.getPageDef(1));
      assert.equal(a.width, 61200); assert.equal(a.landscape, true);
      assert.ok(Math.abs(b.width * 25.4 / 7200 - 240) < .02); assert.equal(b.landscape, false);
      assert.ok(Math.abs((b.marginTop + b.marginHeader) * 25.4 / 7200 - 30) < .02);
      assert.equal(JSON.parse(doc.getHeaderFooter(0, true, 0)).text, paper.header);
      assert.equal(JSON.parse(doc.getHeaderFooter(1, true, 0)).text, second.header);
      assert.equal(JSON.parse(doc.getHeaderFooter(1, false, 0)).text, second.footer);
      const zip = await JSZip.loadAsync(doc.exportHwpx()), xml = await zip.file('Contents/section1.xml').async('string');
      assert.match(xml, /<hp:footNote/); assert.match(xml, /NOTE BODY/); assert.match(xml, /<hp:pic/); assert.match(xml, /numType="PAGE"/);
      assert.doesNotMatch(xml, /WORDSTUDIO-SECTION|WORDSTUDIO-COLUMNS|WORDSTUDIO-NOTE/);
      const converted = editableHwp(doc); html = converted.html; first = converted.page;
      assert.match(html, /FIRST 한글/); assert.match(html, /SECOND/); assert.match(html, /data-section-page=/);
    } finally { doc.free(); }
  }
});

test('empty HWP sections remain intentional pages and nested section markers fail explicitly', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const blank = { ...paper, header: '', footer: '' };
  const section = `<div data-section-page="${JSON.stringify(blank).replaceAll('"', '&quot;')}"></div>`;
  for (const format of ['hwp', 'hwpx']) {
    const out = serializeHwp(HwpDocument, section + section + '<p>THIRD</p>', blank, format), doc = new HwpDocument(out.bytes);
    try { assert.equal(doc.getSectionCount(), 3); assert.equal(doc.pageCount(), 3); assert.equal(doc.getTextRange(2, 0, 0, 100), 'THIRD'); }
    finally { doc.free(); }
    assert.throws(() => serializeHwp(HwpDocument, `<table><tr><td><p>BEFORE</p>${section}<p>AFTER</p></td></tr></table>`, blank, format), /구역 나누기/);
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: native column groups and breaks retain first-page settings and text without marker paragraphs`, async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const html = '<div data-column-count="1" data-column-space="0" data-column-balance="false"><p>FULL WIDTH</p></div>' +
    '<div data-column-count="2" data-column-space="24" data-column-balance="true"><p>LEFT</p><div data-column-break="true"></div><p>RIGHT</p></div><div data-page-break="true"></div><div data-column-count="2" data-column-space="24" data-column-balance="true"><p>NEW PAGE</p></div><p>AFTER</p>';
  const output = serializeHwp(HwpDocument, html, paper, format), doc = new HwpDocument(output.bytes);
  try {
    const files = await JSZip.loadAsync(doc.exportHwpx()), xml = await files.file('Contents/section0.xml').async('string');
    assert.doesNotMatch(xml, /WORDSTUDIO-COLUMNS|WORDSTUDIO-PAGE-BREAK/);
    assert.equal(doc.getParagraphCount(0), 5);
    assert.match(xml, /type="BalancedNewspaper"[^>]*colCount="2"[^>]*sameGap="1800"/);
    assert.match(xml, /<hp:p[^>]*columnBreak="1"[^>]*>[\s\S]*?RIGHT/);
    assert.ok(JSON.parse(doc.getParaPropertiesAt(0, 3)).pageBreakBefore);
    const def = JSON.parse(doc.getPageDef(0)); assert.equal(def.landscape, true); assert.equal(def.width, 61200); assert.equal(def.height, 79200);
    assert.equal(JSON.parse(doc.getHeaderFooter(0, true, 0)).text, paper.header);
    for (const text of ['FULL WIDTH', 'LEFT', 'RIGHT', 'NEW PAGE', 'AFTER']) assert.equal(xml.split(`<hp:t>${text}</hp:t>`).length - 1, 1);
  } finally { doc.free(); }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: native footnote/endnote imports retain their story and reference through two saves`, async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  const native = HwpDocument.createEmpty();
  let bytes;
  try {
    native.createBlankDocument(); native.insertText(0, 0, 0, 'BEFORE AFTER');
    const foot = JSON.parse(native.insertFootnote(0, 0, 6));
    assert.ok(foot.ok);
    assert.ok(JSON.parse(native.insertTextInFootnote(0, 0, foot.controlIdx, 0, 1, 'Native 각주 😀')).ok);
    const end = JSON.parse(native.insertEndnote(0, 0, 0));
    assert.ok(end.ok);
    assert.ok(JSON.parse(native.insertTextInFootnote(0, 0, end.controlIdx, 0, 1, 'Native 미주 설명')).ok);
    bytes = format === 'hwp' ? native.exportHwp() : native.exportHwpx();
  } finally { native.free(); }
  for (let round = 0; round < 3; round++) {
    const source = new HwpDocument(bytes);
    let converted;
    try {
      const zip = await JSZip.loadAsync(source.exportHwpx()), xml = await zip.file('Contents/section0.xml').async('string');
      assert.equal((xml.match(/<hp:footNote\b/g) || []).length, 1);
      assert.equal((xml.match(/<hp:endNote\b/g) || []).length, 1);
      assert.equal((xml.match(/Native 각주/g) || []).length, 1);
      assert.equal((xml.match(/Native 미주/g) || []).length, 1);
      converted = editableHwp(source);
      assert.equal((converted.html.match(/data-note-reference=/g) || []).length, 2);
      assert.equal((converted.html.match(/data-document-note=/g) || []).length, 2);
      assert.match(converted.html, /Native 각주 😀/); assert.match(converted.html, /Native 미주 설명/);
      assert.doesNotMatch(converted.html, /WORDSTUDIO-NOTE|원본 보기에서 확인|변환하지 못한/);
    } finally { source.free(); }
    if (round < 2) bytes = serializeHwp(HwpDocument, converted.html, converted.page, format).bytes;
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: live PAGE/NUMPAGES survive body, cell, nested cell and two saves after astral Unicode`, async () => {
  const { HwpDocument, serializeHwp, editableHwp } = await setup;
  const field = kind => `<span data-page-number="${kind}">37</span>`;
  let html = `<p>START 😀 ${field('PAGE')} / ${field('NUMPAGES')} END</p><table><tr><td><p>CELL ${field('PAGE')} AFTER</p><table><tr><td><p>NESTED 😀 ${field('NUMPAGES')} END</p></td></tr></table></td></tr></table>`;
  for (let save = 0; save < 2; save++) {
    const out = serializeHwp(HwpDocument, html, { ...paper, header: '', footer: '' }, format);
    const doc = new HwpDocument(out.bytes);
    try {
      const parts = await JSZip.loadAsync(doc.exportHwpx()), xml = await parts.file('Contents/section0.xml').async('string');
      assert.equal((xml.match(/<hp:autoNum\s/g) || []).length, 4);
      assert.equal((xml.match(/numType="PAGE"/g) || []).length, 2);
      assert.equal((xml.match(/numType="TOTAL_PAGE"/g) || []).length, 2);
      assert.ok(!xml.includes('WORDSTUDIO-FIELD-')); assert.doesNotMatch(xml, /<hp:t>37<\/hp:t>/);
      html = editableHwp(doc).html;
      assert.equal((html.match(/data-page-number="PAGE"/g) || []).length, 2);
      assert.equal((html.match(/data-page-number="NUMPAGES"/g) || []).length, 2);
      for (const text of ['START 😀', 'END', 'CELL', 'AFTER', 'NESTED 😀']) assert.ok(html.includes(text), text);
    } finally { doc.free(); }
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: supplementary Unicode before empty lines keeps native body and cell pitch`, async () => {
  const { HwpDocument, editableHwp, sourceLinePitches, cellParagraphKey } = await setup;
  const native = HwpDocument.createEmpty(); let bytes;
  const paragraph = kind => `<p>😀${kind}_ALPHA<br><br>${kind}_BETA</p>`;
  try {
    native.createBlankDocument();
    native.pasteHtml(0, 0, 0, paragraph('BODY') + `<table style="width:500px"><tr><td>${paragraph('CELL')}</td></tr></table>`);
    bytes = format === 'hwp' ? native.exportHwp() : native.exportHwpx();
  } finally { native.free(); }
  const doc = new HwpDocument(bytes);
  try {
    const runs = Array.from({ length: doc.pageCount() }, (_, page) => JSON.parse(doc.getPageTextLayout(page)).runs).flat();
    const html = editableHwp(doc).html;
    const pitches = sourceLinePitches(doc);
    for (const kind of ['BODY', 'CELL']) {
      const first = runs.find(r => r.text.includes(kind + '_ALPHA')), last = runs.find(r => r.text.includes(kind + '_BETA'));
      assert.ok(first && last, kind + ': source lines missing');
      const expected = (last.y - first.y) / 2;
      const paragraph = [...html.matchAll(/<p\b([^>]*)>([\s\S]*?)<\/p>/g)].find(p => p[2].includes(kind + '_BETA'));
      assert.ok(paragraph, kind + ': converted paragraph missing');
      const key = first.parentParaIdx === undefined ? `${first.secIdx}:${first.paraIdx}`
        : cellParagraphKey(first.secIdx, first.parentParaIdx, first.cellPath);
      const actual = pitches.get(key);
      assert.ok(Math.abs(actual - expected) < .1, `${kind}: source pitch ${expected}, converted ${actual}`);
    }
  } finally { doc.free(); }
});

test('before-spacing requires matching stored and painted lines across a cell continuation', async () => {
  const { sourceSplitCellBefore } = await setup;
  const address = [0, 0, 0, 0];
  const run = (verticalPosition, textStart) =>
    ({ verticalPosition, textStart, height: 16, textHeight: 16, spacing: 4 });
  const raw = [[run(12, 0), run(32, 6), run(0, 12)], [run(28, 0), run(48, 6)]];
  let terminalText = '';
  const model = {
    getCellParagraphCount: () => raw.length,
    getCellParaPropertiesAt: (section, paragraph, control, cell, p) => JSON.stringify({
      spacingBefore: p ? 16 : 24, spacingAfter: 0, headType: 'None', lineSpacingType: 'Percent', lineSpacing: 125,
    }),
    getCellCharPropertiesAt: () => JSON.stringify({ fontSize: 1200 }),
    getTextInCellByPath: () => terminalText,
    getParagraphLineMetrics: (section, paragraph, path) => JSON.stringify({
      stored: true, controlCount: 0, lines: raw[JSON.parse(path)[0].cellParaIndex],
    }),
  };
  const boxes = [{ pageIndex: 9, y: 60, h: 52 }, { pageIndex: 10, y: 60, h: 68 }];
  const visual = (page, y) => ({ page, y, h: 16 });
  const observed = new Map([
    ['0:0:0:0:0', [visual(9, 74), visual(9, 94), visual(10, 62)]],
    ['0:0:0:0:1', [visual(10, 90), visual(10, 110)]],
  ]);
  assert.deepEqual(sourceSplitCellBefore(model, address, observed, boxes, 2, 2),
    { before: [12, 8], fillsFragments: true });
  assert.deepEqual(sourceSplitCellBefore(model, address, observed, [{ ...boxes[0] }, { ...boxes[1], h: 70 }], 2, 2),
    { before: [12, 8], fillsFragments: false });
  // A source may continue stored offsets across the cut. Its visual
  // starts must still match the stored text offsets and fill the fragments.
  terminalText = 'abcdefghijklmnopqrstuvwx';
  raw[0][2].verticalPosition = 52;
  for (const runs of observed.values()) runs.forEach((visual, i) => { visual.charStart = i * 6; });
  assert.equal(sourceSplitCellBefore(model, address, observed, boxes, 2, 2), undefined);
  assert.deepEqual(sourceSplitCellBefore(model, address, observed, boxes, 2, 2, true),
    { before: [12, 8], fillsFragments: true });
  observed.get('0:0:0:0:0')[2].charStart++;
  assert.equal(sourceSplitCellBefore(model, address, observed, boxes, 2, 2, true), undefined);
  observed.get('0:0:0:0:0')[2].charStart--;
  raw[0][2].verticalPosition = 0;
  terminalText = '';

  observed.get('0:0:0:0:0')[2].y += 4;
  assert.equal(sourceSplitCellBefore(model, address, observed, boxes, 2, 2), undefined);

  // Restore the same authored before-spacing when the page break lands
  // just before a paragraph whose first line cannot fit the old fragment.
  raw[0] = [run(12, 0), run(32, 6)];
  raw[1] = [run(8, 0), run(28, 6)];
  observed.set('0:0:0:0:0', [visual(9, 74), visual(9, 94)]);
  observed.set('0:0:0:0:1', [visual(10, 70), visual(10, 90)]);
  const boundaries = [{ ...boxes[0], h: 66 }, { ...boxes[1], h: 48 }];
  assert.deepEqual(sourceSplitCellBefore(model, address, observed, boundaries, 2, 2),
    { before: [12, 8], fillsFragments: true });
  boundaries[0].h = 82;
  assert.deepEqual(sourceSplitCellBefore(model, address, observed, boundaries, 2, 2),
    { before: [12, 8], fillsFragments: false });

  raw[0] = [run(12, 0), run(32, 6), run(0, 12), run(20, 18)];
  raw[1] = [run(48, 0), run(0, 6), run(20, 12)];
  raw[2] = [run(48, 0)];
  observed.set('0:0:0:0:0', [visual(9, 74), visual(9, 94), visual(10, 62), visual(10, 82)]);
  observed.set('0:0:0:0:1', [visual(10, 110), visual(11, 62), visual(11, 82)]);
  const three = [boxes[0], boxes[1], { pageIndex: 11, y: 60, h: 68 }];
  assert.deepEqual(sourceSplitCellBefore(model, address, observed, three, 2, 2),
    { before: [12, 8, 8], fillsFragments: true });

  terminalText = 'nonempty';
  assert.equal(sourceSplitCellBefore(model, address, observed, three, 2, 2), undefined);
  terminalText = '';
  assert.deepEqual(sourceSplitCellBefore(model, address, observed, [three[0], { ...three[1], h: 75 }, three[2]], 2, 2),
    { before: [12, 8, 8], fillsFragments: false });
});

test('painted HWP cell gaps require both the mixed line pitches and the terminal blank box', async () => {
  const { sourcePaintedCellBefore } = await setup;
  const address = [0, 5, 0, 3];
  const line = (height, spacing, textStart = 0) => ({ height, textHeight: height, spacing, textStart, verticalPosition: 0 });
  const stories = [
    { text: '한글문단 표현확인', lines: [line(16, 6), line(16, 6, 5)] },
    { text: '작은 문자', lines: [line(12, 5)] },
    { text: '', lines: [line(12, 5)] },
  ];
  const story = path => stories[JSON.parse(path)[0].cellParaIndex];
  const model = {
    getCellParagraphCount: () => stories.length,
    getCellParaPropertiesAt: () => JSON.stringify({ spacingBefore: 10, spacingAfter: 0, headType: 'None', lineSpacingType: 'Percent' }),
    getTextInCellByPath: (_section, _paragraph, path) => story(path).text,
    getParagraphLineMetrics: (_section, _paragraph, path) => JSON.stringify({ stored: true, controlCount: 0, lines: story(path).lines }),
  };
  const painted = new Map([
    ['0:5:0:3:0', [{ page: 0, y: 107, h: 16, charStart: 0 }, { page: 1, y: 202, h: 16, charStart: 5 }]],
    ['0:5:0:3:1', [{ page: 1, y: 229, h: 12, charStart: 0 }]],
  ]);
  const pitches = new Map([['page:0:5:0:3:2', 1], ['top:0:5:0:3:2', 251], ['height:0:5:0:3:2', 12]]);
  const boxes = [{ pageIndex: 0, y: 100, h: 35 }, { pageIndex: 1, y: 200, h: 65 }];
  assert.deepEqual(sourcePaintedCellBefore(model, address, painted, pitches, boxes, 2, 2),
    { before: [5, 5, 5], cuts: [{ paragraph: 0, continuedLine: true }], terminalBlank: true, fillsRow: true });
  painted.get('0:5:0:3:1')[0].y = 229.2;
  assert.equal(sourcePaintedCellBefore(model, address, painted, pitches, boxes, 2, 2), undefined);
  painted.get('0:5:0:3:1')[0].y = 229;
  pitches.delete('top:0:5:0:3:2');
  assert.equal(sourcePaintedCellBefore(model, address, painted, pitches, boxes, 2, 2), undefined);
});

test('a confirmed peer can explain a whole-paragraph cut, but cannot move its painted start', async () => {
  const { sourcePaintedCellBefore } = await setup;
  const address = [0, 5, 0, 3];
  const line = { height: 16, textHeight: 16, spacing: 6, textStart: 0 };
  const model = {
    getCellParagraphCount: () => 2,
    getCellParaPropertiesAt: () => JSON.stringify({ spacingBefore: 10, spacingAfter: 0, headType: 'None', lineSpacingType: 'Percent' }),
    getTextInCellByPath: () => '본문',
    getParagraphLineMetrics: () => JSON.stringify({ stored: true, controlCount: 0, lines: [line] }),
  };
  const first = { page: 0, y: 107, h: 16, charStart: 0 }, second = { page: 1, y: 207, h: 16, charStart: 0 };
  const observed = new Map([['0:5:0:3:0', [first]], ['0:5:0:3:1', [second]]]);
  const fragments = [{ pageIndex: 0, y: 100, h: 65 }, { pageIndex: 1, y: 200, h: 25 }];
  assert.equal(sourcePaintedCellBefore(model, address, observed, new Map(), fragments, 2, 2), undefined);
  assert.deepEqual(sourcePaintedCellBefore(model, address, observed, new Map(), fragments, 2, 2, false, true),
    { before: [5, 5], cuts: [{ paragraph: 1, continuedLine: false }], terminalBlank: false, fillsRow: true });
  second.y++;
  assert.equal(sourcePaintedCellBefore(model, address, observed, new Map(), fragments, 2, 2, false, true), undefined);
});

test('authored cell breaks and an intact nested peer require their original measured positions', async () => {
  const { sourcePaintedCellBefore, sourceInlinePeerInMiddleFragment } = await setup;
  let secondStart = 4;
  const line = textStart => ({ height: 16, textHeight: 16, spacing: 6, textStart });
  const model = {
    getCellParagraphCount: () => 1,
    getCellParaPropertiesAt: () => JSON.stringify({ spacingBefore: 10, spacingAfter: 0, headType: 'None', lineSpacingType: 'Percent' }),
    getTextInCellByPath: () => 'abc\nabc',
    getParagraphLineMetrics: () => JSON.stringify({ stored: true, controlCount: 0, lines: [line(0), line(secondStart)] }),
  };
  const address = [0, 5, 0, 3];
  const paint = [{ page: 0, y: 107, h: 16, charStart: 0 }, { page: 0, y: 129, h: 16, charStart: 4 }];
  const map = new Map([['0:5:0:3:0', paint]]);
  const fragments = [{ pageIndex: 0, y: 100, h: 47 }];
  assert.equal(sourcePaintedCellBefore(model, address, map, new Map(), fragments, 2, 2), undefined);
  assert.deepEqual(sourcePaintedCellBefore(model, address, map, new Map(), fragments, 2, 2, true),
    { before: [5], cuts: [], terminalBlank: false, fillsRow: true });
  paint[1].charStart = ++secondStart;
  assert.equal(sourcePaintedCellBefore(model, address, map, new Map(), fragments, 2, 2, true), undefined);

  const first = { cellIdx: 0, pageIndex: 3, x: 102, y: 204, w: 30, h: 20 };
  const second = { cellIdx: 1, pageIndex: 3, x: 132, y: 204, w: 24, h: 20 };
  const peer = {
    getCellParagraphCount: () => 3,
    getTextInCellByPath: () => '',
    getParagraphLineMetrics: (_section, _paragraph, path) => JSON.stringify({
      stored: true, lines: [line(0)],
      ...(JSON.parse(path)[0].cellParaIndex === 1 ? { controlCount: 1, tables: [{ inline: true, index: 0 }] } : { controlCount: 0 }),
    }),
    getTableDimensionsByPath: () => JSON.stringify({ cellCount: 2 }),
    getTableCellBboxesByPath: () => JSON.stringify([first, second]),
  };
  const middle = { pageIndex: 3, x: 100, y: 200, w: 60, h: 30 };
  assert.equal(sourceInlinePeerInMiddleFragment(peer, address, middle), true);
  second.pageIndex++;
  assert.equal(sourceInlinePeerInMiddleFragment(peer, address, middle), false);
  second.pageIndex--;
  second.w += 10;
  assert.equal(sourceInlinePeerInMiddleFragment(peer, address, middle), false);
});
+test('natural HWP page endings use the stored text box only at a proven adjacent boundary', async () => {
  const { sourceTerminalTextAtLineTop } = await setup;
  const positions = new Map([
    ['page:0:3', 0], ['page:0:4', 0], ['page:0:5', 1],
    ['top:0:4', 583.7], ['height:0:4', 16], ['lines:0:4', 1], ['top:0:5', 100],
  ]);
  const paper = { marginTop: 80, marginHeader: 20, footerArea: { y: 600 }, columns: [{}] };
  const line = { verticalPosition: 483.7, height: 16, textHeight: 16, spacing: 9.6, baseline: 13.6, textStart: 0 };
  const model = {
    getPageInfo: () => JSON.stringify(paper),
    getParaPropertiesAt: () => JSON.stringify({ pageBreakBefore: false }),
    getParagraphLineMetrics: (_section, paragraph) => JSON.stringify({
      stored: true, controlCount: 0, lines: paragraph === 4 ? [line] :
        [{ verticalPosition: 456.7, height: 17, spacing: 10 }],
    }),
  };
  const top = () => sourceTerminalTextAtLineTop(model, 0, 4, positions, 16, 25.6);
  assert.equal(top(), true);
  positions.set('top:0:5', 104);
  assert.equal(top(), false);
  positions.set('top:0:5', 100);
  paper.footerArea.y = 605;
  assert.equal(top(), false);
  paper.footerArea.y = 600;
  line.verticalPosition += 2;
  assert.equal(top(), false);
  line.verticalPosition -= 2;
  model.getParaPropertiesAt = () => JSON.stringify({ pageBreakBefore: true });
  assert.equal(top(), false);
});

test('before-spacing in a split cell with a nested table requires both native fragments', async () => {
  const { sourceBeforeNestedCellTail } = await setup;
  const address = [0, 2, 0, 1];
  const first = { pageIndex: 4, y: 60, h: 80 };
  const tail = { pageIndex: 5, y: 60, h: 188, w: 200 };
  const positions = [18, 58.4, 18];
  const bounds = { pageIndex: 5, y: 65, h: 160, w: 170, rowSpan: 1, colSpan: 1 };
  let terminalText = '';
  let inline = true;
  const model = {
    getCellParagraphCount: () => positions.length,
    getCellParaPropertiesAt: () => JSON.stringify({
      spacingBefore: 36, spacingAfter: 0, headType: 'None', lineSpacingType: 'Percent', lineSpacing: 140,
    }),
    getCellCharPropertiesAt: () => JSON.stringify({ fontSize: 1200 }),
    getTextInCellByPath: () => terminalText,
    getParagraphLineMetrics: (_section, _paragraph, path) => {
      const index = JSON.parse(path)[0].cellParaIndex;
      const isTail = index === 2;
      const height = isTail ? 166 : 16;
      return JSON.stringify({
        stored: true, controlCount: +isTail, tables: isTail ? [{ index: 0, inline, outerTop: 3, outerBottom: 3 }] : [],
        lines: [{ verticalPosition: positions[index], textStart: 0, height, textHeight: height, spacing: 6.4 }],
      });
    },
    getTableDimensionsByPath: () => JSON.stringify({ rowCount: 1, colCount: 1, cellCount: 1 }),
    getTableCellBboxesByPath: () => JSON.stringify([bounds]),
  };
  const pitches = new Map([
    ['page:0:2:0:1:0', 4], ['top:0:2:0:1:0', 80], ['height:0:2:0:1:0', 16], ['lines:0:2:0:1:0', 1],
    ['page:0:2:0:1:1', 4], ['top:0:2:0:1:1', 120.4], ['height:0:2:0:1:1', 16], ['lines:0:2:0:1:1', 1],
  ]);
  const recover = (last = tail) => sourceBeforeNestedCellTail(model, address, pitches, first, last, 2, 2);
  assert.deepEqual(recover(), [18, 18]);
  assert.equal(recover({ ...tail, h: 184 }), undefined);
  inline = false;
  assert.equal(recover(), undefined);
  inline = true;
  terminalText = 'Separate title';
  assert.equal(recover(), undefined);
  terminalText = '';
  bounds.pageIndex = 6;
  assert.equal(recover(), undefined);
  bounds.pageIndex = 5;
  pitches.set('top:0:2:0:1:0', 81);
  assert.equal(recover(), undefined);
});

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
          const heights = [...converted.html.matchAll(/<tr\b[^>]*\sstyle="height:([\d.]+)px;/g)].map(match => Number(match[1]));
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
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-hwp-cell-content-'));
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

for (const format of ['hwp', 'hwpx']) test(`${format}: hard line breaks survive in ordinary body and plain nested cells`, async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty();
  try {
    source.createBlankDocument();
    source.pasteHtml(0, 0, 0, '<p>BODY_ALPHA<br>BODY_BETA</p><table><tr><td><p>CELL_ALPHA<br>CELL_BETA</p><table><tr><td><p>NESTED_ALPHA<br>NESTED_BETA</p></td></tr></table></td></tr></table>');
    assert.equal(source.getTextRange(0, 0, 0, 0xffffffff), 'BODY_ALPHA\nBODY_BETA');
    const doc = new HwpDocument(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
    try {
      const html = editableHwp(doc).html;
      const text = html.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>/g, '');
      for (const prefix of ['BODY', 'CELL', 'NESTED']) assert.ok(text.includes(`${prefix}_ALPHA\n${prefix}_BETA`), prefix + ': ' + html);
      assert.equal((html.match(/<br\s*\/?>/g) || []).length, 3);
    } finally { doc.free(); }
  } finally { source.free(); }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: inactive stored cell margins use table padding, including nested cells and explicit zero`, async () => {
  const { HwpDocument, editableHwp, serializeHwp } = await setup;
  const source = HwpDocument.createEmpty(); source.createBlankDocument();
  source.pasteHtml(0, 0, 0, '<p></p><table><tr><td><p>INHERITED</p><table><tr><td><p>NESTED</p></td></tr></table></td><td><p>ZERO</p></td></tr></table><p></p>');
  const zip = await JSZip.loadAsync(source.exportHwpx()); source.free();
  let xml = await zip.file('Contents/section0.xml').async('string');
  xml = xml.replace(/<hp:inMargin\b[^>]*\/>/g, '<hp:inMargin left="75" right="150" top="225" bottom="300"/>')
    .replace(/<hp:cellMargin\b[^>]*\/>/g, '<hp:cellMargin left="1800" right="1800" top="1800" bottom="1800"/>')
    .replace(/(<hp:tc\b[^>]*\bhasMargin=)"[^"]*"/g, '$1"0"');
  // The final cell explicitly opts into zero padding. Zero is a real override.
  const last = xml.lastIndexOf('<hp:tc ');
  xml = xml.slice(0, last) + xml.slice(last).replace('hasMargin="0"', 'hasMargin="1"').replace(/<hp:cellMargin\b[^>]*\/>/, '<hp:cellMargin left="0" right="0" top="0" bottom="0"/>');
  zip.file('Contents/section0.xml', xml);
  const fixture = new HwpDocument(await zip.generateAsync({ type: 'uint8array' }));
  let bytes = format === 'hwp' ? fixture.exportHwp() : fixture.exportHwpx(); fixture.free();
  for (let round = 0; round < 3; round++) {
    const doc = new HwpDocument(bytes);
    try {
      const imported = editableHwp(doc), cells = [...imported.html.matchAll(/<td\b[^>]*style="([^"]*)"[^>]*>/g)].map(m => m[1]);
      assert.equal(cells.length, 3);
      for (const css of cells.slice(0, 2)) {
        const padding = [...css.matchAll(/(?:^|;)padding:([^;]+);/g)].at(-1)[1].split(/\s+/).map(parseFloat);
        assert.deepEqual(padding, [3, 2, 4, 1]);
      }
      const padding = [...cells[2].matchAll(/(?:^|;)padding:([^;]+);/g)].at(-1)[1].split(/\s+/).map(parseFloat);
      assert.ok(padding.every(p => p === 0), cells[2]);
      if (round < 2) bytes = serializeHwp(HwpDocument, imported.html, imported.page, format).bytes;
    } finally { doc.free(); }
  }
});

test('single-line final cell paragraphs discard trailing leading and later exact CSS wins over adjacent percentages', async () => {
  const { HwpDocument, editableHwp, serializeHwp } = await setup;
  for (const format of ['hwp', 'hwpx']) {
    let html = '<table style="width:400px"><tr><td style="padding:0"><p style="font-size:12pt;line-height:200%"><span style="font-size:12pt">FIRST LINE</span></p><p style="font-size:12pt;line-height:150%"><span style="font-size:12pt">FINAL LINE</span></p></td></tr></table>';
    for (let round = 0; round < 2; round++) {
      const output = serializeHwp(HwpDocument, html, paper, format), doc = new HwpDocument(output.bytes);
      try {
        const imported = editableHwp(doc); html = imported.html;
        const paragraphs = [...html.matchAll(/<p\b[^>]*style="([^"]*)"[^>]*>([\s\S]*?)<\/p>/g)];
        const css = paragraphs.find(p => p[2].includes('FINAL LINE'))[1];
        const rules = Object.fromEntries(css.split(';').filter(p => p.includes(':')).map(p => { const i = p.indexOf(':'); return [p.slice(0, i).trim(), p.slice(i + 1).trim()]; }));
        assert.equal(rules['line-height'], '16px', `${format} round ${round}: ${css}`);
        assert.ok(paragraphs.some(p => p[2].includes('FIRST LINE')));
      } finally { doc.free(); }
    }
  }
});

for (const format of ['hwp', 'hwpx']) test(`${format}: empty cell paragraphs retain their slots and the following paragraph format through two saves`, async () => {
  const { HwpDocument, editableHwp, serializeHwp } = await setup;
  let html = '<table><tr><td><p>BEFORE</p><p style="font-size:18pt;line-height:40px"></p><p style="text-align:right">AFTER</p></td><td><p></p></td></tr></table>';
  for (let round = 0; round < 3; round++) {
    const output = serializeHwp(HwpDocument, html, paper, format), doc = new HwpDocument(output.bytes);
    try {
      const table = JSON.parse(doc.getControls()).find(c => c.ctrlId === 'tbl' && c.list === 0);
      assert.equal(doc.getCellParagraphCount(0, table.para, table.controlIndex, 0), 3);
      const imported = editableHwp(doc); html = imported.html;
      const cell = html.match(/<td\b[^>]*>([\s\S]*?)<\/td>/)[1];
      const paragraphs = [...cell.matchAll(/<p\b([^>]*)>([\s\S]*?)<\/p>/g)];
      assert.equal(paragraphs.length, 3);
      assert.equal(paragraphs[1][2].replace(/<[^>]*>/g, '').trim(), '');
      assert.match(paragraphs[1][1], /line-height:40(?:\.0)?px/);
      assert.match(paragraphs[2][1], /text-align:right/);
      assert.match(paragraphs[2][2], /AFTER/);
    } finally { doc.free(); }
  }
});

test('native fixed leading keeps physical px/pt units and the last CSS declaration through repeated saves', async () => {
  const { HwpDocument, editableHwp, serializeHwp } = await setup;
  for (const format of ['hwp', 'hwpx']) for (const line of ['40px', '30pt', '160%;line-height:40px']) {
    let html = `<p style="line-height:${line}">FIXED LEADING</p>`;
    for (let round = 0; round < 3; round++) {
      const output = serializeHwp(HwpDocument, html, paper, format), doc = new HwpDocument(output.bytes);
      try {
        const properties = JSON.parse(doc.getParaPropertiesAt(0, 0));
        assert.equal(properties.lineSpacingType, 'Fixed', `${format} ${line} round ${round}`);
        assert.equal(properties.lineSpacing, 40, `${format} ${line} round ${round}`);
        html = editableHwp(doc).html;
      } finally { doc.free(); }
    }
  }
});

test('automatic zero-height rows share the remaining merged-cell height without inflating the final row', async () => {
  const { HwpDocument, editableHwp } = await setup;
  const source = HwpDocument.createEmpty(); source.createBlankDocument();
  source.pasteHtml(0, 0, 0, '<table style="width:400px"><tr><td><p>HEAD</p></td><td><p>HEAD2</p></td></tr><tr><td><p>A</p></td><td rowspan="3"><p>MERGED</p></td></tr><tr><td><p>B</p></td></tr><tr><td><p>C</p></td></tr></table>');
  const zip = await JSZip.loadAsync(source.exportHwpx()); source.free();
  const heights = [1500, 1500, 1500, 6000, 0, 0]; let i = 0;
  let xml = await zip.file('Contents/section0.xml').async('string');
  xml = xml.replace(/<hp:cellSz\b[^>]*\/>/g, tag => tag.replace(/height="\d+"/, `height="${heights[i++]}"`));
  assert.equal(i, heights.length);
  zip.file('Contents/section0.xml', xml);
  const fixture = new HwpDocument(await zip.generateAsync({ type: 'uint8array' }));
  try {
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(format === 'hwp' ? fixture.exportHwp() : fixture.exportHwpx());
      try {
        const html = editableHwp(doc).html;
        const rows = [...html.matchAll(/<tr\b[^>]*style="[^"]*height:([\d.]+)px/g)].map(m => Number(m[1]));
        assert.deepEqual(rows, [20, 20, 30, 30], format);
        for (const text of ['HEAD', 'MERGED', 'A', 'B', 'C']) assert.ok(html.includes(text));
      } finally { doc.free(); }
    }
  } finally { fixture.free(); }
});

test('native CSS cascade applies inner formatting after inherited declarations', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  for (const format of ['hwp', 'hwpx']) {
    const html = '<p><span style="font-weight:normal;font-size:10pt;font-style:normal"><strong style="font-size:12pt;font-size:18pt">B</strong><em>I</em>N<span style="font-style:italic"><span style="font-style:normal">R</span></span></span></p>';
    const doc = new HwpDocument(serializeHwp(HwpDocument, html, paper, format).bytes);
    try {
      const chars = [0, 1, 2, 3].map(i => JSON.parse(doc.getCharPropertiesAt(0, 0, i)));
      assert.equal(chars[0].bold, true); assert.equal(chars[0].fontSize, 1800);
      assert.equal(chars[1].italic, true); assert.equal(chars[1].fontSize, 1000);
      assert.equal(chars[2].bold, false); assert.equal(chars[2].italic, false);
      assert.equal(chars[3].italic, false);
    } finally { doc.free(); }
  }
});

test('large native exports retain tables and exact paragraph boundaries', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const marker = '<div data-page-break="true"></div>';
  const first = '<p style="text-align:right">첫 쪽</p><table><tr><td rowspan="2"><p>공통 셀</p></td><td><p>상단</p></td></tr><tr><td><p>하단</p></td></tr></table><p></p>';
  const second = '<p style="text-align:center">둘째 쪽</p><p>마지막 문장</p>';
  const padding = '한'.repeat(350000);
  const padded = html => html.replace('<p style=', `<p title="${padding}" style=`);
  const small = first + marker + second, large = padded(first) + marker + padded(second);
  assert.ok(large.length < 2000000 && Buffer.byteLength(large) > 2000000);

  function inventory(bytes) {
    const doc = new HwpDocument(bytes);
    try {
      const paragraphs = Array.from({ length:doc.getParagraphCount(0) }, (_, index) => {
        const properties = JSON.parse(doc.getParaPropertiesAt(0, index));
        return [doc.getTextRange(0, index, 0, 0xffffffff), properties.alignment, properties.pageBreakBefore];
      });
      const tables = JSON.parse(doc.getControls()).filter(control => control.list === 0 && control.ctrlId === 'tbl');
      assert.equal(tables.length, 1);
      const [{ para, controlIndex }] = tables;
      const dimensions = JSON.parse(doc.getTableDimensions(0, para, controlIndex));
      assert.equal(dimensions.cellCount, 3);
      const cells = Array.from({ length:dimensions.cellCount }, (_, index) =>
        [JSON.parse(doc.getCellInfo(0, para, controlIndex, index)).rowSpan,
          doc.getTextInCell(0, para, controlIndex, index, 0, 0, 0xffffffff)]);
      return { pages:doc.pageCount(), paragraphs, cells };
    } finally { doc.free(); }
  }

  const withoutFurniture = { ...paper, header:'', footer:'' };
  for (const format of ['hwp','hwpx']) {
    const expected = inventory(serializeHwp(HwpDocument, small, withoutFurniture, format).bytes);
    assert.deepEqual(expected.cells, [[2,'공통 셀'],[1,'상단'],[1,'하단']]);
    const actual = inventory(serializeHwp(HwpDocument, large, withoutFurniture, format).bytes);
    assert.deepEqual(actual, expected);
    const sharedPage = inventory(serializeHwp(HwpDocument, padded(first) + padded(second), withoutFurniture, format).bytes);
    assert.deepEqual(sharedPage, inventory(serializeHwp(HwpDocument, first + second, withoutFurniture, format).bytes));
    const indivisible = `<p title="${padding.repeat(2)}">한 문단</p>`;
    assert.throws(() => serializeHwp(HwpDocument, indivisible, withoutFurniture, format), /문서 구조 보존 한도/);
  }
});

test('measured flowing table stays editable beside its one-row label in both native formats', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const label = '<table><tr><td><p>표 제목</p></td></tr></table>';
  const blank = '<p><span></span></p>';
  const flow = '<table data-hwp-centered-borders="continuation" data-table-anchor="{&quot;x&quot;:0,&quot;y&quot;:0,&quot;bottom&quot;:0,&quot;reference&quot;:&quot;paragraph&quot;,&quot;horizontal&quot;:&quot;margin&quot;}"><tr><th><p>표 머리</p></th></tr><tr><td><p>표 본문</p></td></tr></table>';
  const page = { ...paper, header: '', footer: '' };

  for (const format of ['hwp', 'hwpx']) {
    for (const grouped of [true, false]) {
      const before = grouped ? blank : '<p>보존할 문단</p>';
      const { bytes } = serializeHwp(HwpDocument, label + before + flow, page, format);
      const native = new HwpDocument(bytes);
      try {
        const tables = JSON.parse(native.getControls()).filter(control => control.list === 0 && control.ctrlId === 'tbl');
        assert.equal(tables.length, 2);
        const [heading, body] = tables;
        const info = control => JSON.parse(native.getTableProperties(0, control.para, control.controlIndex));
        assert.ok(info(heading).treatAsChar);
        assert.equal(info(body).treatAsChar, !grouped);
        assert.equal(info(body).pageBreak, grouped ? 2 : 0);
        if (grouped) assert.deepEqual([info(heading).outerTop, info(heading).outerBottom], [0, 0]);
        assert.equal(native.getTextRange(0, heading.para + 1, 0, 0xffffffff), grouped ? '' : '보존할 문단');
        assert.equal(native.getTextInCell(0, body.para, body.controlIndex, 1, 0, 0, 0xffffffff), '표 본문');
      } finally { native.free(); }
    }
  }
});

test('native body and cell runs retain authored widths, text color and spacing', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const html = '<p><span style="--word-character-scale:95;color:rgb(12, 34, 56);letter-spacing:-.5pt">한글 &amp; latin</span><span style="color:#123456">그대로</span></p>' +
    '<table><tr><td><p><span style="--word-character-scale:120;color:#c39"><b></b>지정 아님</span><span style="--word-character-scale:88;color:#1a2b3c">표 안쪽</span></p></td></tr></table>';
  const settings = { ...paper, header:'', footer:'' };
  for (const format of ['hwp','hwpx']) {
    const { bytes, warnings } = serializeHwp(HwpDocument, html, settings, format);
    assert.deepEqual(warnings, []);
    const doc = new HwpDocument(bytes);
    try {
      const controls = JSON.parse(doc.getControls());
      const table = controls.find(c=>c.list===0&&c.ctrlId==='tbl');
      assert.ok(table);
      const row = doc.getTextRange(0, 0, 0, 0xffffffff);
      assert.equal(row, '한글 & latin그대로');
      const width = props => [props.ratios, props.textColor.toLowerCase()];
      const first = JSON.parse(doc.getCharPropertiesAt(0,0,0));
      assert.deepEqual(width(first), [Array(7).fill(95), '#0c2238']);
      const control = new HwpDocument(serializeHwp(HwpDocument, html.replace('--word-character-scale:95;', ''), settings, format).bytes);
      try {
        const withoutWidth = JSON.parse(control.getCharPropertiesAt(0,0,0));
        assert.ok(withoutWidth.spacings.every(spacing => spacing < 0));
        assert.deepEqual(first.spacings, withoutWidth.spacings);
      } finally { control.free(); }
      assert.deepEqual(width(JSON.parse(doc.getCharPropertiesAt(0,0,'한글 & latin'.length))), [Array(7).fill(100), '#123456']);
      const literal = doc.getTextInCell(0,table.para,table.controlIndex,0,0,0,0xffffffff);
      assert.equal(literal, '지정 아님표 안쪽');
      assert.deepEqual(width(JSON.parse(doc.getCellCharPropertiesAt(0,table.para,table.controlIndex,0,0,0))), [Array(7).fill(100), '#cc3399']);
      assert.deepEqual(width(JSON.parse(doc.getCellCharPropertiesAt(0,table.para,table.controlIndex,0,0,'지정 아님'.length))), [Array(7).fill(88), '#1a2b3c']);
    } finally { doc.free(); }
  }
});

test('native width survives an unstyled wrapper around the final spaces of the run', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const html = '<p><span style="--word-character-scale:95;color:#123456;letter-spacing:-0.6pt">본문 한 줄<span> </span></span></p>';
  const settings = { ...paper, header:'', footer:'' };
  for (const format of ['hwp', 'hwpx']) {
    const { bytes, warnings } = serializeHwp(HwpDocument, html, settings, format);
    assert.deepEqual(warnings, []);
    const doc = new HwpDocument(bytes);
    const unscaled = new HwpDocument(serializeHwp(HwpDocument, html.replace('--word-character-scale:95;', ''), settings, format).bytes);
    try {
      const text = '본문 한 줄 ';
      assert.equal(doc.getTextRange(0, 0, 0, 0xffffffff), text);
      for (let index = 0; index < text.length; index++) {
        const { ratios, spacings, textColor } = JSON.parse(doc.getCharPropertiesAt(0, 0, index));
        assert.deepEqual(ratios, Array(7).fill(95));
        const original = JSON.parse(unscaled.getCharPropertiesAt(0, 0, index));
        assert.ok(original.spacings.every(spacing => spacing < 0));
        assert.deepEqual(spacings, original.spacings);
        assert.equal(textColor.toLowerCase(), '#123456');
      }
    } finally { unscaled.free(); doc.free(); }
  }
});

test('unattributed native width colors fall back to the original editable document', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  class PasteWithoutMarker extends HwpDocument {
    static createEmpty() {
      const doc = HwpDocument.createEmpty();
      const paste = doc.pasteHtml.bind(doc);
      doc.pasteHtml = (s, p, at, html) => paste(s, p, at, html.replace(
        /(--word-character-scale:95;color:)#[0-9a-f]{6}/, '$1#123456'));
      return doc;
    }
  }
  const html = '<p><span style="--word-character-scale:95;color:#123456">원래 색</span></p><p>다음 문단</p>';
  const settings = { ...paper, header:'', footer:'' };
  for (const format of ['hwp','hwpx']) {
    const { bytes, warnings } = serializeHwp(PasteWithoutMarker, html, settings, format);
    assert.match(warnings[0], /장평/);
    const doc = new HwpDocument(bytes);
    try {
      assert.equal(doc.getTextRange(0, 0, 0, 0xffffffff), '원래 색');
      assert.equal(doc.getTextRange(0, 1, 0, 0xffffffff), '다음 문단');
      const properties = JSON.parse(doc.getCharPropertiesAt(0, 0, 0));
      assert.equal(properties.textColor.toLowerCase(), '#123456');
      assert.deepEqual(properties.ratios, Array(7).fill(100));
    } finally { doc.free(); }
  }
});

test('native spread circles for decimal labels preserve editable text in body and nested cells', async () => {
  const { HwpDocument, editableHwp, serializeHwp } = await setup;
  const html = '<p>BODYEND</p><table><tr><td><p>OUTER</p><table><tr><td><p>NESTEDEND</p></td></tr></table></td></tr></table><p>GUARDEND</p>';
  const settings = { ...paper, header: '', footer: '' };
  const staged = await JSZip.loadAsync(serializeHwp(HwpDocument, html, settings, 'hwpx').bytes);
  const section = 'Contents/section0.xml';
  const markup = await staged.file(section).async('string');
  const circle = (a, b) => '</hp:t><hp:compose circleType="SHAPE_CIRCLE" charSz="-3" composeType="SPREAD" charPrCnt="2" composeText="' +
    String.fromCodePoint(a, b) + '"><hp:charPr prIDRef="0"/><hp:charPr prIDRef="0"/></hp:compose><hp:t>';
  staged.file(section, markup
    .replace('BODYEND', 'BODY' + circle(0xf0289, 0xf0293) + 'END')
    .replace('NESTEDEND', 'NESTED' + circle(0xf0289, 0xf0298) + 'END')
    .replace('GUARDEND', 'GUARD' + circle(0xf0288, 0xf0293) + 'END'));
  const hwpx = await staged.generateAsync({type:'uint8array'});
  const output = new HwpDocument(hwpx);
  let hwp;
  try { hwp = output.exportHwp(); } finally { output.free(); }
  for (const bytes of [hwp, hwpx]) {
    const native = new HwpDocument(bytes);
    try {
      const data = editableHwp(native);
      const text = data.html.replace(/<[^>]+>/g, '');
      assert.match(text, /BODY⑪END/);
      assert.match(text, /NESTED⑯END/);
      assert.match(text, /GUARD.*(?:변환하지 못한 개체|원본 보기에서 확인).*END/);
      assert.doesNotMatch(data.html, /[\u{F0288}-\u{F029B}]/u);
    } finally { native.free(); }
  }
});

test('native save restores a short merged grid after the native parser inflated its cell', async () => {
  const { HwpDocument, serializeHwp, planHwpGridRows } = await setup;
  const settings = { ...paper, header:'', footer:'' };
  const rows = '<table style="width:220px"><tr style="height:24px"><td rowspan="3" style="width:100px;height:52px"><p>MERGED</p></td><td style="width:120px;height:24px"><p>TOP</p></td></tr>' +
    '<tr data-hwp-grid-row="true" style="height:4px"><td style="width:120px;height:4px"><p></p></td></tr>' +
    '<tr style="height:24px"><td style="width:120px;height:24px"><p>BOTTOM</p></td></tr></table>';
  const state = doc => {
    const c=JSON.parse(doc.getControls()).find(control=>control.list===0&&control.ctrlId==='tbl');
    const pos=[0,c.para,c.controlIndex];
    const props=JSON.parse(doc.getCellProperties(...pos,2));
    const text=Array.from({length:4},(_,cell)=>doc.getTextInCell(...pos,cell,0,0,0xffffffff));
    return [props.height,text];
  };
  const input = serializeHwp(HwpDocument,rows.replace(' data-hwp-grid-row="true"', ''),settings,'hwpx');
  const intermediate = await JSZip.loadAsync(input.bytes), section = 'Contents/section0.xml';
  const xml = await intermediate.file(section).async('string');
  const adjusted = xml.replace(/<hp:tc\b[\s\S]*?<\/hp:tc>/g, cell =>
    /<hp:cellAddr\b[^>]*\browAddr="1"/.test(cell) ? cell.replace(/(<hp:cellSz\b[^>]*\bheight=")300"/, (tag, prefix) => prefix + '1000"') : cell);
  assert.notEqual(xml,adjusted);
  intermediate.file(section,adjusted);
  const source = new HwpDocument(await intermediate.generateAsync({type:'uint8array'}));
  try {
    assert.deepEqual(state(source), [1000,['MERGED','TOP','','BOTTOM']]);
    // A cell size explicitly different from its row size is a separate user
    // setting. The reverse layout proof cannot revise it on the row's behalf.
    const mismatched = rows.replace('width:120px;height:4px','width:120px;height:5px');
    assert.equal(planHwpGridRows(mismatched)(HwpDocument,source), source);
    const restored=planHwpGridRows(rows)(HwpDocument,source);
    assert.notEqual(restored,source);
    try {
      assert.deepEqual(state(restored), [300,['MERGED','TOP','','BOTTOM']]);
      for (const bytes of [restored.exportHwpx(),restored.exportHwp()]) {
        const reopened = new HwpDocument(bytes);
        try { assert.deepEqual(state(reopened),state(restored)); } finally { reopened.free(); }
      }
    } finally { restored.free(); }
  } finally { source.free(); }
});

test('a native cell page reference retains its original numeric display and unique bookmark target', async () => {
  const { HwpDocument, editableHwp, serializeHwp } = await setup;
  const settings = { ...paper, header: '', footer: '' };
  const html = '<p><span data-docx-bookmark="대상 &amp; 절"></span>DESTINATION</p>' +
    '<table><tr><td><p>LEFT12</p></td><td><p>GUARD34</p></td></tr></table>';
  const staging = await JSZip.loadAsync(serializeHwp(HwpDocument,html,settings,'hwpx').bytes);
  const name='Contents/section0.xml', xml=await staging.file(name).async('string');
  const field = (id, target, digits) =>
    '</hp:t><hp:ctrl><hp:fieldBegin id="' + id + '" type="CROSSREF" name="" editable="0" dirty="0">' +
    '<hp:parameters cnt="1" name=""><hp:stringParam name="Command">?' + target +
    ';6;0;1;0;</hp:stringParam></hp:parameters></hp:fieldBegin></hp:ctrl><hp:t>' + digits +
    '</hp:t><hp:ctrl><hp:fieldEnd beginIDRef="' + id + '"/></hp:ctrl><hp:t>';
  staging.file(name,xml.replace('LEFT12','LEFT' + field(301,'대상 &amp; 절','12'))
    .replace('GUARD34','GUARD' + field(302,'없는 책갈피','34')));
  const zipped = await staging.generateAsync({type:'uint8array'}), origin=new HwpDocument(zipped);
  let binary; try { binary=origin.exportHwp(); } finally { origin.free(); }
  for (const native of [zipped,binary]) {
    const d=new HwpDocument(native);
    try {
      const result=editableHwp(d);
      assert.match(result.html,/LEFT<a href="#대상 &amp; 절">12<\/a>/);
      assert.match(result.html,/data-docx-bookmark="대상 &amp; 절"/);
      assert.match(result.html,/GUARD34/);
      assert.match(result.html,/\[변환하지 못한 개체/);
      assert.doesNotMatch(result.html,/href="#없는 책갈피"/);
      assert.ok(result.warnings.some(w=>w.includes('쪽 번호를 표시')));
    }finally {d.free();}
  }
});

test('binary save preserves identical intentional body paragraphs on explicit separate pages', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const settings = { ...paper, header: '', footer: '' };
  const body = '<p>Keep this repeated line.</p>', breakTag = '<div data-page-break="true"></div>';
  const { bytes, pages, warnings } = serializeHwp(HwpDocument,[body,body,body].join(breakTag),settings,'hwp');
  assert.equal(pages,3);
  assert.deepEqual(warnings,[]);
  const reopened = new HwpDocument(bytes);
  try {
    for (let page = 0; page < 3; page++) {
      const runs = JSON.parse(reopened.getPageTextLayout(page)).runs.filter(run=>!run.headerFooter);
      assert.equal(runs.map(run=>run.text).join(''),'Keep this repeated line.');
    }
  } finally { reopened.free(); }
});

test('native table keeps the measured one-line lead and the complete following table together', async () => {
  const { HwpDocument, serializeHwp } = await setup;
  const settings = { size: 'A4', orientation: 'portrait', margin: 20,
    margins: { top: 30, right: 20, bottom: 20, left: 20 }, header: '', footer: '', pageNumbers: false };
  const lead = '조판 검증을 위한 한 줄 예시이며 원본의 글자와 표 속 내용을 그대로 유지할 수 있습니다'.slice(0,44);
  const line = text => '<p style="text-align:justify;text-indent:-40px;line-height:130%;margin:0">' +
    '<span style="font-family:serif;font-size:14pt">' + text + '</span></p>';
  const table = '<table style="width:600px"><tr style="height:854px"><td style="width:600px;height:854px"><p>TABLE CELL</p></td></tr></table>';
  const html = line('STATIONARY') + line(lead) + table + '<div data-page-break="true"></div>' + line('FOLLOWING');
  const printed = { pages: 2, tableLeads: [{ sheet: 0, literal: lead, width: 600, height: 854, rows: 1 }] };
  const normal = serializeHwp(HwpDocument, html, settings, 'hwpx');
  const corrected = serializeHwp(HwpDocument, html, settings, 'hwpx', undefined, [], printed);
  assert.equal(normal.pages, 3);
  assert.equal(corrected.pages, 2);
  assert.deepEqual(corrected.warnings, normal.warnings);
  const doc = new HwpDocument(corrected.bytes);
  try {
    const runs = JSON.parse(doc.getPageTextLayout(0)).runs;
    const leadRuns = runs.filter(run => run.paraIdx === 1 && !run.cellPath?.length);
    assert.equal(leadRuns.length, 1);
    assert.equal(leadRuns[0].text, lead);
    assert.equal(doc.getTextInCell(0, 2, 0, 0, 0, 0, 0xffffffff), 'TABLE CELL');
    assert.equal(JSON.parse(doc.getPageTextLayout(1)).runs.map(run => run.text).join(''), 'FOLLOWING');
  } finally { doc.free(); }
  printed.tableLeads[0].literal = lead + 'X';
  const unproven = serializeHwp(HwpDocument, html, settings, 'hwpx', undefined, [], printed);
  assert.deepEqual(unproven.bytes, normal.bytes);
});

test('a mixed column pair requires two physically painted lines at the stored cut', async () => {
  const { parseHwpSourceWrap, preserveHwpColumnLineEnds } = await setup;
  const text = '첫째 줄 둘째 줄';
  const html = '<p><span>' + text + '</span></p>';
  const stored = { stored: true, controlCount: 0, lines: [
    { textStart: 0, verticalPosition: 40, width: 320, height: 18, textHeight: 18, spacing: 7.2 },
    { textStart: 5, verticalPosition: 65.2, width: 320, height: 15.3, textHeight: 15.3, spacing: 6.1 },
  ] };
  const first = { page: 0, y: 140, h: 18, charStart: 0 };
  const second = { page: 0, y: 165.2, h: 15.3, charStart: 5 };
  const proof = (painted, record = stored) => {
    const marked = preserveHwpColumnLineEnds(html, text, () => JSON.stringify(record), painted);
    assert.equal(marked.replace(/ data-hwp-source-wrap="[^"]+"/, ''), html);
    return parseHwpSourceWrap(marked.match(/data-hwp-source-wrap="([^"]+)"/)?.[1].replaceAll('&quot;', '"'));
  };
  const matched = proof([first, second]);
  assert.deepEqual(matched.ends, [3, 6]);
  assert.equal(matched.columnWidth, 320);
  assert.deepEqual(matched.columnPitches, [25.2, 21.4]);
  for (const invalid of [{ ...second, page: 1 }, { ...second, y: 166 }, { ...second, h: 18 }, { ...second, charStart: 6 }]) {
    const { columnPitches, ...oldProof } = matched;
    assert.deepEqual(proof([first, invalid]), oldProof);
  }
  for (const invalid of [[25.2], [25.2, 25.2], [25.2, -1], [25.2, 10000]]) {
    assert.equal(parseHwpSourceWrap(JSON.stringify({ ...matched, columnPitches: invalid })), undefined);
  }
});

test('whole-paragraph peers must supply the measured end of every source fragment', async () => {
  const { sourceWholeParagraphPeerRow } = await setup;
  const profile = (count, cuts, fillsRow) => ({
    before: Array(count).fill(12),
    cuts: cuts.map(paragraph => ({ paragraph, continuedLine: false })),
    terminalBlank: false,
    fillsRow,
  });
  const compact = profile(2, [1], false), longer = profile(5, [1], true);
  const two = [[true, true], [false, true]];
  assert.equal(sourceWholeParagraphPeerRow([compact, longer], two), true);
  assert.equal(sourceWholeParagraphPeerRow([compact, longer], [[true, false], [false, true]]), false);
  const long = profile(11, [3, 8], true), middle = profile(5, [1], false);
  const three = [[true, false], [false, true], [true, false]];
  assert.equal(sourceWholeParagraphPeerRow([long, middle], three), true);
  assert.equal(sourceWholeParagraphPeerRow([long, middle], [[true, false], [false, false], [true, false]]), false);
  assert.equal(sourceWholeParagraphPeerRow([long, { ...middle, terminalBlank: true }], three), false);
  assert.equal(sourceWholeParagraphPeerRow([long, { ...middle, cuts: [{ paragraph: 1, continuedLine: true }] }], three), false);
});
