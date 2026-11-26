const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
module.exports = async function createFixture(directory) {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  let doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>BEFORE</p><table style="width:400px"><tr><td style="width:400px;height:120px;padding:0"><p>TABLE BODY</p></td></tr></table><p>AFTER</p>');
    const zip = unzipSync(doc.exportHwpx());
    zip['Contents/section0.xml'] = strToU8(strFromU8(zip['Contents/section0.xml']).replace('</hp:tbl>', '</hp:tbl><hp:t>SUFFIX</hp:t>'));
    const loaded = new HwpDocument(zipSync(zip)); doc.free(); doc = loaded;
    const table = JSON.parse(doc.getControls()).find(c => c.ctrlId === 'tbl' && c.list === 0);
    doc.setTableProperties(0, table.para, table.controlIndex, JSON.stringify({ treatAsChar: true }));
    const files = {};
    for (const format of ['hwp', 'hwpx']) {
      files[format] = path.join(directory, 'inline-tail.' + format);
      await fs.writeFile(files[format], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
    }
    return files;
  } finally { doc.free(); }
};
