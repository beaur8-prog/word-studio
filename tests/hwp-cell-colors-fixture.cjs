const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  await fs.mkdir(directory, { recursive: true }); const files = {};
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.setPageDef(0, JSON.stringify({ width: 59528, height: 84188, landscape: false, marginTop: 5669, marginBottom: 5669, marginLeft: 5669, marginRight: 5669, marginHeader: 0, marginFooter: 0, marginGutter: 0 }));
    const c = JSON.parse(doc.createTable(0, 0, 0, 2, 3));
    doc.setTableProperties(0, c.paraIdx, c.controlIdx, JSON.stringify({ pageBreak: 1, repeatHeader: true }));
    for (let cell = 0; cell < 6; cell++) {
      const texts = cell < 3 ? [['BLUE HEADER', 'YELLOW HEADER', 'CLEAR HEADER'][cell]] : Array.from({ length: 85 }, (_, i) => `BODY ${cell} ${i}`);
      for (let p = 0; p < texts.length; p++) {
        if (p) doc.splitParagraphInCell(0, c.paraIdx, c.controlIdx, cell, p - 1, texts[p - 1].length);
        doc.insertTextInCell(0, c.paraIdx, c.controlIdx, cell, p, 0, texts[p]);
        doc.applyCharFormatInCell(0, c.paraIdx, c.controlIdx, cell, p, 0, 0xffffffff, JSON.stringify({ fontSize: 1000, ...(cell === 0 ? { textColor: '#ffffff', bold: true } : cell === 1 ? { textColor: '#c00000' } : {}) }));
        doc.applyParaFormatInCell(0, c.paraIdx, c.controlIdx, cell, p, JSON.stringify({ lineSpacingType: 'Fixed', lineSpacing: 2000 }));
      }
      if (cell < 3) doc.setCellProperties(0, c.paraIdx, c.controlIdx, cell, JSON.stringify({ isHeader: true, ...(cell < 2 ? { fillType: 'solid', fillColor: cell === 0 ? '#1f4e78' : '#fff2cc' } : {}) }));
    }
    for (const format of ['hwp', 'hwpx']) {
      const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), loaded = new HwpDocument(bytes);
      try {
        const headers = Array.from({ length: 3 }, (_, cell) => JSON.parse(loaded.getCellProperties(0, c.paraIdx, c.controlIdx, cell)));
        if (headers[0].fillColor !== '#1f4e78' || headers[1].fillColor !== '#fff2cc' || !headers.every(h => h.isHeader)) throw Error('Fixture did not retain native header colors');
      } finally { loaded.free(); }
      files[format] = path.join(directory, 'header-colors.' + format); await fs.writeFile(files[format], bytes);
    }
    return files;
  } finally { doc.free(); }
};
