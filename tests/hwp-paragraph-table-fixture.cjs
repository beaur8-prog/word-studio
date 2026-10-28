const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  await fs.mkdir(directory, { recursive: true }); const files = {};
  for (const [name, offset, rows] of [['zero', 0, 1], ['offset', 1500, 1], ['long', 1500, 14]]) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      doc.setPageDef(0, JSON.stringify({ width: 42520, height: 51024, landscape: false, marginTop: 2835, marginBottom: 2835, marginLeft: 2835, marginRight: 2835, marginHeader: 0, marginFooter: 0, marginGutter: 0 }));
      doc.pasteHtml(0, 0, 0, '<p style="font-size:10pt;line-height:20px">BEFORE</p><table style="width:450px">' + Array.from({ length: rows }, (_, i) => `<tr><td style="width:450px;height:75px;padding:0"><p>FLOAT ROW ${i}</p></td></tr>`).join('') + '</table><p style="font-size:10pt;line-height:20px">AFTER</p><p style="line-height:17px"> </p>');
      const c = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
      doc.setTableProperties(0, c.para, c.controlIndex, JSON.stringify({ pageBreak: 2, repeatHeader: false, treatAsChar: false, textWrap: 'TopAndBottom', vertRelTo: 'Para', vertAlign: 'Top', vertOffset: offset, horzRelTo: 'Column', horzAlign: 'Left', horzOffset: 600, outerTop: 225, outerBottom: 300 }));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), loaded = new HwpDocument(bytes);
        try {
          const position = JSON.parse(loaded.getPageControlLayout(0)).controls.find(c => c.type === 'table');
          const properties = JSON.parse(loaded.getTableProperties(0, c.para, c.controlIndex));
          files[`${format}-${name}`] = { path: path.join(directory, `${name}.${format}`), x: position.x, y: position.y, pages: loaded.pageCount(), offset: offset / 75, bottom: properties.outerBottom / 75, rows };
          await fs.writeFile(files[`${format}-${name}`].path, bytes);
        } finally { loaded.free(); }
      }
    } finally { doc.free(); }
  }
  return files;
};
