const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

module.exports = async function createFixture() {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')));
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const source = HwpDocument.createEmpty();
  try {
    source.createBlankDocument();
    source.pasteHtml(0, 0, 0, '<p>BEFORE</p><table style="width:320px">' +
      '<tr><td rowspan="2" style="width:160px"><p>UPPER OWNER</p></td><td style="width:160px"><p>UPPER RIGHT</p></td></tr>' +
      '<tr><td rowspan="2" style="width:160px"><p>LOWER OWNER</p></td></tr>' +
      '<tr><td style="width:160px"><p>LOWER LEFT</p></td></tr></table><p>AFTER</p>');
    const table = JSON.parse(source.getControls()).find(control => control.list === 0 && control.ctrlId === 'tbl');
    const update = [
      { row: 0, col: 0, rowSpan: 2, height: 2400 },
      { row: 0, col: 1, rowSpan: 1, height: 2250 },
      { row: 1, col: 1, rowSpan: 2, height: 2400 },
      { row: 2, col: 0, rowSpan: 1, height: 2250 },
    ];
    const { cellCount } = JSON.parse(source.getTableDimensions(0, table.para, table.controlIndex));
    const cells = Array.from({ length: cellCount }, (_, index) =>
      JSON.parse(source.getCellInfo(0, table.para, table.controlIndex, index)));
    for (const cell of update) {
      const { row, col, rowSpan, height } = cell;
      const index = cells.findIndex(cell => cell.row === row && cell.col === col);
      if (cells[index]?.rowSpan !== rowSpan) {
        throw Error('Fixture must have the intended pair of crossing vertical merges');
      }
      source.setCellProperties(0, table.para, table.controlIndex, index,
        JSON.stringify({ height, paddingTop: 225, paddingBottom: 225, applyInnerMargin: true, verticalAlign: 0 }));
    }
    source.reflowLinesegs();
    return { hwp: source.exportHwp(), hwpx: source.exportHwpx() };
  } finally { source.free(); }
};
