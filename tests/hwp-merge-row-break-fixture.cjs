const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  await fs.mkdir(directory, { recursive: true }); const files = {};
  for (const policy of [0, 2]) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      doc.pasteHtml(0, 0, 0, '<p>LEAD</p><table style="width:300px"><tr><td rowspan="2" style="width:100px;height:320px"><p>MERGED ONCE</p></td><td style="width:200px;height:160px"><p>ROW ONE</p></td></tr><tr><td style="width:200px;height:160px"><p>ROW TWO</p></td></tr></table>');
      const c = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
      doc.setTableProperties(0, c.para, c.controlIndex, JSON.stringify({ pageBreak: policy, repeatHeader: false, treatAsChar: false, textWrap: 'TopAndBottom', vertRelTo: 'Para', vertOffset: 0 }));
      for (let cell = 0; cell < 3; cell++) doc.setCellProperties(0, c.para, c.controlIndex, cell, JSON.stringify({ verticalAlign: 0 }));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), original = new HwpDocument(bytes);
        try {
          const boxes = JSON.parse(original.getTableCellBboxes(0, c.para, c.controlIndex));
          if (original.pageCount() !== 1 || new Set(boxes.filter(b => b.cellIdx === 0).map(b => b.pageIndex)).size !== 1)
            throw Error('Fixture must start with its merged cell entirely on one source page');
        } finally { original.free(); }
        files[`${format}-${policy}`] = path.join(directory, `merge-${policy}.${format}`);
        await fs.writeFile(files[`${format}-${policy}`], bytes);
      }
    } finally { doc.free(); }
  }
  return files;
};
