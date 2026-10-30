const fs = require('node:fs/promises'), path = require('node:path'), assert = require('node:assert/strict');
module.exports = async function createFixture(directory) {
  const { default: init, HwpDocument } = await import('@rhwp/core');
  await init({ module_or_path: await fs.readFile(path.join(path.dirname(require.resolve('@rhwp/core')), 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    doc.pasteHtml(0, 0, 0, '<p>BODY &amp; 😀_ALPHA<br>BODY_BETA</p><p>😀😀😀😀MULTI_ALPHA<br><br>MULTI_BETA<br></p>' +
      '<table><tr><td><p>CELL_ALPHA<br>CELL_BETA</p><table><tr><td><p>NESTED_ALPHA<br>NESTED_BETA</p></td></tr></table></td></tr></table>' +
      `<p>PICTURE_ALPHA<br>😀&amp;<img src="data:image/png;base64,${png}" width="24" height="24">PICTURE_BETA<br>PICTURE_END</p><p>TAIL</p>`);
    assert.equal(doc.getTextRange(0, 0, 0, 0xffffffff), 'BODY & 😀_ALPHA\nBODY_BETA');
    assert.equal(doc.getTextRange(0, 1, 0, 0xffffffff), '😀😀😀😀MULTI_ALPHA\n\nMULTI_BETA\n');
    const files = {};
    for (const format of ['hwp', 'hwpx']) { files[format] = path.join(directory, 'hard-breaks.' + format); await fs.writeFile(files[format], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx()); }
    const runs = Array.from({ length: doc.pageCount() }, (_, page) => JSON.parse(doc.getPageTextLayout(page)).runs).flat();
    const first = runs.find(r => r.text === 'MULTI_ALPHA'), last = runs.find(r => r.text === 'MULTI_BETA');
    assert.ok(first && last);
    return { ...files, multiLineGap: last.y - first.y };
  } finally { doc.free(); }
};
