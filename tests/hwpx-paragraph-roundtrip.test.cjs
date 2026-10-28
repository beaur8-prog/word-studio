'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
let HwpDocument;
test.before(async () => {
  const root = process.env.RHWP_CORE_DIRECTORY || path.dirname(require.resolve('@rhwp/core'));
  const core = await import(pathToFileURL(path.join(root, 'rhwp.js')));
  await core.default({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) }); HwpDocument = core.HwpDocument;
});

for (const origin of ['HTML in modern package', 'legacy plain HWPX', 'binary HWP']) test(`${origin}: paragraph margins and hanging indent remain stable through two HWPX saves`, () => {
  let doc = HwpDocument.createEmpty();
  const values = (d, paragraph) => {
    const p = JSON.parse(d.getParaPropertiesAt(0, paragraph));
    return Object.fromEntries(['marginLeft', 'marginRight', 'indent', 'spacingBefore', 'spacingAfter', 'lineSpacing', 'lineSpacingType'].map(k => [k, p[k]]));
  };
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>LEAD</p><p style="margin-left:12px;margin-right:8px;text-indent:-5px;margin-top:6px;margin-bottom:4px;line-height:37px">HANGING 한글</p>');
    if (origin !== 'HTML in modern package') {
      let bytes = origin === 'binary HWP' ? doc.exportHwp() : doc.exportHwpx();
      if (origin === 'legacy plain HWPX') {
        const zip = unzipSync(bytes);
        zip['version.xml'] = strToU8(strFromU8(zip['version.xml']).replace(/xmlVersion="[^"]*"/, 'xmlVersion="1.2"'));
        bytes = zipSync(zip);
      }
      const seed = new HwpDocument(bytes); doc.free(); doc = seed;
    }
    const paragraph = Array.from({ length: doc.getParagraphCount(0) }, (_, i) => i).find(i => doc.getTextRange(0, i, 0, doc.getParagraphLength(0, i)).includes('HANGING'));
    assert.notEqual(paragraph, undefined);
    const expected = values(doc, paragraph);
    assert(expected.indent < 0); assert(expected.spacingBefore > 0); assert(expected.spacingAfter > 0);
    for (let save = 0; save < 2; save++) {
      const again = new HwpDocument(doc.exportHwpx()); doc.free(); doc = again;
      assert.deepEqual(values(doc, paragraph), expected, `save ${save + 1}`);
      assert.equal(doc.getTextRange(0, paragraph, 0, doc.getParagraphLength(0, paragraph)), 'HANGING 한글');
    }
  } finally { doc.free(); }
});
