const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty(), result = {};
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<table style="width:500px"><tr><td style="padding:4px;border:0.2mm solid #000000"><p>OUTER</p><table style="width:460px"><tr><td style="width:220px;border-top:0.2mm solid #c00000;border-right:0.5mm dashed #0070c0;border-bottom:0.3mm double #008000;border-left:none;padding:4px"><p>NESTED LABEL</p></td><td style="width:240px;border:0.2mm solid #000000;padding:4px"><p>NESTED VALUE</p></td></tr></table><p>TAIL</p></td></tr></table>');
    const outer = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
    const tablePath = JSON.stringify([{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: 1 }, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }]);
    for (const format of ['hwp', 'hwpx']) {
      const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), loaded = new HwpDocument(bytes);
      try {
        const dimensions = JSON.parse(loaded.getTableDimensionsByPath(0, 0, tablePath));
        if (dimensions.cellCount !== 2) throw Error('Invalid nested-border fixture');
        const cells = [0, 1].map(i => JSON.parse(loaded.getCellPropertiesByPath(0, 0, tablePath, i)));
        const edges = ['Top', 'Right', 'Bottom', 'Left'].map(side => cells[0]['border' + side]);
        if (JSON.stringify(edges.map(b => [b.type, b.width, b.color])) !== JSON.stringify([[1, 3, '#c00000'], [2, 7, '#0070c0'], [8, 5, '#008000'], [0, 0, '#000000']])) throw Error('Fixture did not retain its physical border styles');
        const file = path.join(directory, 'nested-borders.' + format); await fs.writeFile(file, bytes); result[format] = { file, cells };
      } finally { loaded.free(); }
    }
    return result;
  } finally { doc.free(); }
};
