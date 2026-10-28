const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
function populate(doc) {
  doc.createBlankDocument();
  doc.setPageDef(0, JSON.stringify({ width: 59528, height: 84188, landscape: false, marginTop: 5669, marginBottom: 5669,
    marginLeft: 5669, marginRight: 5669, marginHeader: 0, marginFooter: 0, marginGutter: 0 }));
  const c = JSON.parse(doc.createTable(0, 0, 0, 2, 2));
  doc.setTableProperties(0, c.paraIdx, c.controlIdx, JSON.stringify({ pageBreak: 1, repeatHeader: true }));
  for (let cell = 0; cell < 4; cell++) {
    const texts = cell < 2 ? ['HEADER ' + cell] : Array.from({ length: cell === 2 ? 110 : 75 }, (_, i) => `${cell === 2 ? 'LEFT' : 'RIGHT'} ${i} 한글 paragraph`);
    for (let p = 0; p < texts.length; p++) {
      if (p) doc.splitParagraphInCell(0, c.paraIdx, c.controlIdx, cell, p - 1, texts[p - 1].length);
      doc.insertTextInCell(0, c.paraIdx, c.controlIdx, cell, p, 0, texts[p]);
      doc.applyCharFormatInCell(0, c.paraIdx, c.controlIdx, cell, p, 0, 0xffffffff, JSON.stringify({ fontSize: 1000 }));
      doc.applyParaFormatInCell(0, c.paraIdx, c.controlIdx, cell, p, JSON.stringify({ lineSpacingType: 'Fixed', lineSpacing: 1500 }));
    }
    if (cell < 2) doc.setCellProperties(0, c.paraIdx, c.controlIdx, cell, JSON.stringify({ isHeader: true }));
  }
  return c;
}
module.exports = async function createRowFlowFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    populate(doc); await fs.mkdir(directory, { recursive: true }); const files = {};
    for (const format of ['hwp', 'hwpx']) {
      files[format] = path.join(directory, 'row-flow.' + format);
      await fs.writeFile(files[format], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
    }
    return files;
  } finally { doc.free(); }
};
module.exports.populate = populate;
