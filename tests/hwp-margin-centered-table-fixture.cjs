const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
let fixture;

module.exports = () => fixture ??= (async () => {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(core + '/rhwp.js'));
  await init({ module_or_path: await fs.readFile(core + '/rhwp_bg.wasm') });
  const result = {};
  for (const alignment of ['Center', 'Left']) {
    const original = HwpDocument.createEmpty();
    original.createBlankDocument();
    try {
      const def = JSON.parse(original.getPageDef(0));
      original.setPageDef(0, JSON.stringify({ ...def, width: 600 * 75, height: 840 * 75,
        marginLeft: 75 * 75, marginRight: 45 * 75, marginTop: 62 * 75, marginBottom: 62 * 75,
        marginHeader: 10 * 75, marginFooter: 10 * 75 }));
      original.pasteHtml(0, 0, 0, '<table style="width:410px"><tr><td style="width:410px;height:54px;border:1px solid black;padding:0">' +
        '<p style="margin:0;line-height:20px">Math table title</p></td></tr></table><p>Body follows this.</p>');
      const table = JSON.parse(original.getControls()).find(control => control.list === 0 && control.ctrlId === 'tbl');
      if (table?.para !== 0) throw Error('Expected first paragraph title table');
      original.setTableProperties(0, 0, table.controlIndex, JSON.stringify({
        treatAsChar: false, textWrap: 'TopAndBottom', pageBreak: 0, repeatHeader: false,
        vertRelTo: 'Paper', vertAlign: 'Top', vertOffset: 101 * 75, horzRelTo: 'Page', horzAlign: alignment,
        horzOffset: 0, outerTop: 0, outerBottom: 16 * 75, allowOverlap: false, restrictInPage: true,
      }));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? original.exportHwp() : original.exportHwpx();
        const reader = new HwpDocument(bytes);
        try {
          const props = JSON.parse(reader.getTableProperties(0, 0, table.controlIndex));
          if (props.horzAlign !== alignment || props.horzRelTo !== 'Page' || props.vertRelTo !== 'Paper' || props.vertAlign !== 'Top')
            throw Error('Fixture lost its authored mixed-axis anchor');
          const [place] = JSON.parse(reader.getTableCellBboxes(0, 0, table.controlIndex));
          if (alignment === 'Center' && (Math.abs(place.x - 110) > .1 ||
            Math.abs(place.y - (props.vertOffset + props.outerTop) / 75) > .1)) throw Error('Native engine did not center the fixture');
          result[format + ':' + alignment] = { bytes, place };
        } finally { reader.free(); }
      }
    } finally { original.free(); }
  }
  return result;
})();
