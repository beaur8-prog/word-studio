const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
module.exports = async function createHwpLayoutFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument(); doc.insertText(0, 0, 0, 'FIRST 한글');
    doc.insertColumnBreak(0, 0, doc.getParagraphLength(0, 0)); doc.insertText(0, 1, 0, 'SECOND 둘째 단');
    doc.insertPageBreak(0, 1, doc.getParagraphLength(0, 1)); doc.insertText(0, 2, 0, 'THIRD 다음 쪽');
    doc.insertEquation(0, 2, doc.getParagraphLength(0, 2), 'a over b', 1200, 0);
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==', 'base64');
    doc.insertPicture(0, 0, 0, '[]', png, 3600, 3600, 1, 1, 'png', 'BACKDROP');
    const c = JSON.parse(doc.getControls()).find(c => c.userDesc === '그림');
    doc.setPictureProperties(0, c.para, c.controlIndex, JSON.stringify({ ...JSON.parse(doc.getPictureProperties(0, c.para, c.controlIndex)),
      treatAsChar: false, textWrap: 'BehindText', horzRelTo: 'Paper', vertRelTo: 'Paper', horzOffset: 9000, vertOffset: 12000 }));
    doc.setPageDef(0, JSON.stringify({ width: 84188, height: 119056, landscape: false, marginTop: 5669, marginBottom: 5669, marginLeft: 5669, marginRight: 5669, marginHeader: 0, marginFooter: 0, marginGutter: 0 }));
    doc.setColumnDef(0, 2, 0, 1, 1800);
    await fs.mkdir(directory, { recursive: true });
    const files = {};
    for (const extension of ['hwp', 'hwpx']) {
      const bytes = extension === 'hwp' ? doc.exportHwp() : doc.exportHwpx();
      files[extension] = path.join(directory, 'layout.' + extension); await fs.writeFile(files[extension], bytes);
    }
    return files;
  } finally { doc.free(); }
};
