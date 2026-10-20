'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const JSZip = require('jszip');

const fixture = path.join(__dirname, 'fixtures/hwp/manifest-relative.hwpx');
const coreDirectory = path.resolve(process.env.RHWP_CORE_DIRECTORY || path.dirname(require.resolve('@rhwp/core')));
const core = (async () => {
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(coreDirectory, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(coreDirectory, 'rhwp_bg.wasm')) });
  return HwpDocument;
})();
const archive = async () => JSZip.loadAsync(await fs.readFile(fixture));
const bytes = zip => zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
async function render(input) {
  const HwpDocument = await core;
  const doc = new HwpDocument(input);
  try {
    const pages = [];
    for (let i = 0; i < doc.pageCount(); i++) {
      const svg = doc.renderPageSvg(i);
      pages.push({ svg: createHash('sha256').update(svg).digest('hex'), layout: JSON.parse(doc.getPageTextLayout(i)),
        width: Number(svg.match(/<svg[^>]*\bwidth="([\d.]+)"/)[1]), height: Number(svg.match(/<svg[^>]*\bheight="([\d.]+)"/)[1]) });
    }
    return pages;
  } finally { doc.free(); }
}
const textOf = pages => pages.flatMap(page => page.layout.runs.map(run => run.text)).join('');

test('OPF-relative section opens without a ZIP-root copy and matches the root-path control', async () => {
  const zip = await archive();
  assert.equal(zip.file('section0.xml'), null);
  assert.match(await zip.file('META-INF/container.xml').async('string'), /full-path="Contents\/content.hpf"/);
  const pages = await render(await bytes(zip));
  assert.equal(pages.length, 1);
  assert.ok(Math.abs(pages[0].width - 793.7) < 1 && Math.abs(pages[0].height - 1122.5) < 1, 'A4 paper');
  assert.match(textOf(pages), /첫 번째 문단: HWPX 상대 경로 검증/);
  assert.match(textOf(pages), /두 번째 문단: Word Studio/);
  zip.file('Contents/content.hpf', (await zip.file('Contents/content.hpf').async('string')).replace('href="section0.xml"', 'href="Contents/section0.xml"'));
  assert.deepEqual(pages, await render(await bytes(zip)), 'Changing only path notation must not change rendering');
});

test('OPF-relative section takes precedence over an unrelated ZIP-root section', async () => {
  const zip = await archive();
  const section = await zip.file('Contents/section0.xml').async('string');
  zip.file('section0.xml', section.replace('두 번째 문단: Word Studio', 'ROOT DECOY'));
  const text = textOf(await render(await bytes(zip)));
  assert.match(text, /두 번째 문단: Word Studio/);
  assert.doesNotMatch(text, /ROOT DECOY/);
});

test('relative subdirectories and dot segments preserve OPF spine order', async () => {
  const zip = await archive();
  zip.file('Contents/sub/section1.xml', (await zip.file('Contents/section0.xml').async('string')).replace('첫 번째 문단:', 'SPINE FIRST:'));
  const manifest = (await zip.file('Contents/content.hpf').async('string'))
    .replace('href="section0.xml"', 'href="./parts/../section0.xml"')
    .replace('</opf:manifest>', '<opf:item id="section1" href="sub/section1.xml" media-type="application/xml"/></opf:manifest>')
    .replace('<opf:spine>', '<opf:spine><opf:itemref idref="section1"/>');
  zip.file('Contents/content.hpf', manifest);
  const text = textOf(await render(await bytes(zip)));
  assert.ok(text.indexOf('SPINE FIRST:') >= 0 && text.indexOf('SPINE FIRST:') < text.indexOf('첫 번째 문단:'), 'Use spine order, not file names');
});

test('relative master pages and parent-directory binary resources are actually loaded', async () => {
  const zip = await archive();
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==', 'base64');
  zip.file('BinData/image1.png', png);
  zip.file('Contents/masterpage0.xml', '<masterPage xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" id="masterpage0" type="BOTH" pageNumber="0"><hp:subList textWidth="40000" textHeight="70000"><hp:p paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>MASTER RESOURCE</hp:t></hp:run></hp:p></hp:subList></masterPage>');
  zip.file('Contents/content.hpf', (await zip.file('Contents/content.hpf').async('string'))
    // Keep section lookup compatible with 1.3.0 to isolate the resource bug.
    .replace('href="section0.xml"', 'href="Contents/section0.xml"')
    .replace('<opf:manifest>', '<opf:manifest><opf:item id="masterpage0" href="masterpage0.xml" media-type="application/xml"/><opf:item id="image1" href="../BinData/image1.png" media-type="image/png" isEmbeded="1"/>'));
  const input = await bytes(zip);
  assert.match(textOf(await render(input)), /MASTER RESOURCE/);
  const HwpDocument = await core;
  const doc = new HwpDocument(input);
  try {
    // Export forces the lazy ZIP resource to be read; a manifest entry alone
    // would not prove its bytes survived resolution.
    const exported = await JSZip.loadAsync(doc.exportHwpx());
    assert.deepEqual(await exported.file('BinData/image1.png').async('nodebuffer'), png);
  } finally { doc.free(); }
});

test('missing and unsafe package paths remain errors', async () => {
  const HwpDocument = await core;
  const missing = await archive(); missing.remove('Contents/section0.xml');
  const missingBytes = await bytes(missing);
  assert.throws(() => new HwpDocument(missingBytes), /필수 파일 누락: Contents\/section0\.xml/);
  for (const href of ['../../section0.xml', '/section0.xml', 'https://example.com/section0.xml', 'C:\\section0.xml']) {
    const zip = await archive();
    zip.file('Contents/content.hpf', (await zip.file('Contents/content.hpf').async('string')).replace('href="section0.xml"', `href="${href}"`));
    const input = await bytes(zip);
    assert.throws(() => new HwpDocument(input), /invalid package path/);
  }
});
