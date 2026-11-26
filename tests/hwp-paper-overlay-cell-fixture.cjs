const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');

let fixture;
module.exports = () => fixture ??= (async () => {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(core + '/rhwp.js'));
  await init({ module_or_path: await fs.readFile(core + '/rhwp_bg.wasm') });
  const original = HwpDocument.createEmpty();
  let authored;
  try {
    original.createBlankDocument();
    const page = JSON.parse(original.getPageDef(0));
    original.setPageDef(0, JSON.stringify({ ...page, width: 600 * 75, height: 850 * 75,
      marginLeft: 75 * 75, marginRight: 45 * 75, marginTop: 62 * 75, marginBottom: 62 * 75,
      marginHeader: 10 * 75, marginFooter: 10 * 75 }));
    original.pasteHtml(0, 0, 0, '<table style="width:428px"><tr><td style="width:428px;height:44px;border:1px solid black"><p>TABLE CELL</p></td></tr></table><p>AFTER TABLE</p>');
    const table = JSON.parse(original.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
    original.setTableProperties(0, table.para, table.controlIndex, JSON.stringify({ treatAsChar: true }));
    const shape = JSON.parse(original.createShapeControl(JSON.stringify({ sectionIdx: 0, paraIdx: 1, charOffset: 0,
      shapeType: 'ellipse', width: 7200, height: 6000, treatAsChar: false, textWrap: 'InFrontOfText' })));
    original.setShapeProperties(0, shape.paraIdx, shape.controlIdx, JSON.stringify({ treatAsChar: false, textWrap: 'InFrontOfText',
      horzRelTo: 'Paper', horzAlign: 'Left', horzOffset: 330 * 75,
      vertRelTo: 'Paper', vertAlign: 'Top', vertOffset: 430 * 75, borderColor: 255, borderWidth: 35, fillType: 'none' }));
    const archive = unzipSync(original.exportHwpx());
    const xml = strFromU8(archive['Contents/section0.xml']);
    const ellipse = xml.match(/<hp:ellipse\b[\s\S]*?<\/hp:ellipse>/)?.[0];
    if (!ellipse) throw Error('Missing authored overlay');
    // Keep a visible tail after the inline table and attach the page overlay
    // to the same source line. Studio lays the tail out in a narrow cell.
    archive['Contents/section0.xml'] = strToU8(xml.replace(ellipse, '').replace('</hp:tbl>', '</hp:tbl><hp:t>TAIL</hp:t>' + ellipse));
    authored = new HwpDocument(zipSync(archive));
    authored.reflowLinesegs();
    const result = {};
    for (const format of ['hwp', 'hwpx']) {
      const bytes = format === 'hwp' ? authored.exportHwp() : authored.exportHwpx();
      const source = new HwpDocument(bytes);
      try {
        const control = JSON.parse(source.getControls()).find(c => c.list === 0 && c.userDesc === '타원');
        const props = JSON.parse(source.getShapeProperties(0, control.para, control.controlIndex));
        const bounds = JSON.parse(source.getShapeBBox(0, control.para, control.controlIndex));
        if (control.para !== table.para || props.treatAsChar || props.textWrap !== 'InFrontOfText' || props.horzRelTo !== 'Paper' || props.vertRelTo !== 'Paper')
          throw Error('Invalid authored paper anchor in ' + format);
        result[format] = { bytes, bounds };
      } finally { source.free(); }
    }
    return result;
  } finally { authored?.free(); original.free(); }
})();
