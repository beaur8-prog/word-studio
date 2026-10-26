const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createOverlayFixture(directory, verticalAlign = 0) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')));
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<table><tr><td style="height:300px;width:300px;padding:8px"><p>EDITABLE OVERLAY 한글</p><p>SECOND LINE</p></td></tr></table>');
    let table = JSON.parse(doc.getControls()).find(c => c.ctrlId === 'tbl' && c.list === 0);
    doc.setCellProperties(0, table.para, table.controlIndex, 0, JSON.stringify({ verticalAlign }));
    doc.setTableProperties(0, table.para, table.controlIndex, JSON.stringify({ treatAsChar: false, textWrap: 'InFrontOfText', horzRelTo: 'Para', vertRelTo: 'Para', horzOffset: 0, vertOffset: 0 }));
    const png = await require('sharp')({ create: { width: 120, height: 120, channels: 4, background: { r: 180, g: 220, b: 250, alpha: 1 } } }).png().toBuffer();
    doc.insertPicture(0, table.para, 0, '[]', png, 22500, 22500, 120, 120, 'png', 'FORM BACKDROP');
    const picture = JSON.parse(doc.getControls()).find(c => c.userDesc === '그림' && c.list === 0);
    doc.setPictureProperties(0, picture.para, picture.controlIndex, JSON.stringify({ treatAsChar: true, textWrap: 'TopAndBottom' }));
    await fs.mkdir(directory, { recursive: true }); const files = {};
    for (const format of ['hwp', 'hwpx']) { files[format] = path.join(directory, 'overlay.' + format); await fs.writeFile(files[format], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx()); }
    return files;
  } finally { doc.free(); }
};
