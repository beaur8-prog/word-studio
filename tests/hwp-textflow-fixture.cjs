const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

function populate(doc, fontFamily) {
  doc.createBlankDocument();
  const c = JSON.parse(doc.createTable(0, 0, 0, 1, 1));
  const texts = ['UPPER TITLE', 'BIG TITLE', '     '];
  for (let i = 0; i < texts.length; i++) {
    if (i) doc.splitParagraphInCell(0, c.paraIdx, c.controlIdx, 0, i - 1, texts[i - 1].length);
    doc.insertTextInCell(0, c.paraIdx, c.controlIdx, 0, i, 0, texts[i]);
    doc.applyCharFormatInCell(0, c.paraIdx, c.controlIdx, 0, i, 0, 0xffffffff, JSON.stringify({ fontSize: i === 0 ? 2100 : 4000, ...(fontFamily ? { fontFamily } : {}) }));
    doc.applyParaFormatInCell(0, c.paraIdx, c.controlIdx, 0, i, JSON.stringify({ lineSpacingType: 'Percent', lineSpacing: [180, 0, 150][i] }));
  }
  doc.splitParagraph(0, 0, doc.getParagraphLength(0, 0));
  doc.insertText(0, 1, 0, String.fromCodePoint(0xF081C).repeat(80) + ' NORMAL ┈');
  doc.applyCharFormat(0, 1, 0, 0xffffffff, JSON.stringify({ fontSize: 1000 }));
}
module.exports = async function createTextflowFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  await fs.mkdir(directory, { recursive: true }); const files = {};
  for (const font of ['', 'Studio Text']) {
    const doc = HwpDocument.createEmpty();
    try {
      populate(doc, font);
      for (const format of ['hwp', 'hwpx']) {
        const key = format + (font ? '-tall' : '');
        files[key] = path.join(directory, 'textflow-' + (font ? 'tall' : 'default') + '.' + format);
        await fs.writeFile(files[key], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
      }
    } finally { doc.free(); }
  }
  return files;
};
module.exports.populate = populate;
