const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
function populateShapes(doc) {
  doc.createBlankDocument(); doc.insertText(0, 0, 0, 'BODY TEXT');
  const title = JSON.parse(doc.createShapeControl(JSON.stringify({ sectionIdx: 0, paraIdx: 0, charOffset: 0,
    shapeType: 'textbox', width: 14400, height: 3600, treatAsChar: false, textWrap: 'InFrontOfText' })));
  doc.insertTextInCell(0, title.paraIdx, title.controlIdx, 0, 0, 0, 'TITLE 한글');
  doc.setShapeProperties(0, title.paraIdx, title.controlIdx, JSON.stringify({ horzRelTo: 'Paper', vertRelTo: 'Paper', horzOffset: 9000, vertOffset: 7200,
    borderColor: 0, borderWidth: 75, fillType: 'solid', fillBgColor: 0xffffff }));
  const ellipse = JSON.parse(doc.createShapeControl(JSON.stringify({ sectionIdx: 0, paraIdx: 0, charOffset: 0,
    shapeType: 'ellipse', width: 6000, height: 4500, treatAsChar: false, textWrap: 'InFrontOfText' })));
  doc.setShapeProperties(0, ellipse.paraIdx, ellipse.controlIdx, JSON.stringify({ horzRelTo: 'Paper', vertRelTo: 'Paper', horzOffset: 30000, vertOffset: 45000,
    borderColor: 255, borderWidth: 75, fillType: 'none' }));
}
module.exports = async function createShapeFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    populateShapes(doc); await fs.mkdir(directory, { recursive: true }); const files = {};
    for (const format of ['hwp', 'hwpx']) {
      files[format] = path.join(directory, 'shapes.' + format);
      await fs.writeFile(files[format], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
    }
    return files;
  } finally { doc.free(); }
};
module.exports.populateShapes = populateShapes;
