const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  let doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p>LEAD</p><table style="width:200px"><tr><td style="height:80px"><p>TABLE TEXT</p></td></tr></table><p>TAIL</p><p style="line-height:30px"> </p><table><tr><td><p>SECOND</p></td></tr></table>');
    const archive = unzipSync(doc.exportHwpx());
    // Put indentation and suffix whitespace on the table's own source line;
    // the native insertion API otherwise creates a separate table paragraph.
    archive['Contents/section0.xml'] = strToU8(strFromU8(archive['Contents/section0.xml'])
      .replace('<hp:tbl ', '<hp:t>    </hp:t><hp:tbl ').replace('</hp:tbl>', '</hp:tbl><hp:t>  </hp:t>'));
    const reopened = new HwpDocument(zipSync(archive)); doc.free(); doc = reopened;
    const tables = JSON.parse(doc.getControls()).filter(c => c.list === 0 && c.ctrlId === 'tbl');
    const c = tables[0]; doc.setTableProperties(0, c.para, c.controlIndex, JSON.stringify({ treatAsChar: true }));
    doc.applyCharFormat(0, c.para, 0, doc.getParagraphLength(0, c.para), JSON.stringify({ fontSize: 3000 }));
    doc.applyParaFormat(0, c.para, JSON.stringify({ lineSpacingType: 'Percent', lineSpacing: 150 }));
    for (let p = 0; p < doc.getParagraphCount(0); p++) if (doc.getTextRange(0, p, 0, doc.getParagraphLength(0, p)) === ' ')
      doc.applyParaFormat(0, p, JSON.stringify({ lineSpacingType: 'Fixed', lineSpacing: 4500 }));
    await fs.mkdir(directory, { recursive: true }); const files = {};
    for (const format of ['hwp', 'hwpx']) {
      files[format] = path.join(directory, 'inline-table.' + format);
      await fs.writeFile(files[format], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
    }
    files.x = JSON.parse(doc.getPageControlLayout(0)).controls.find(c => c.type === 'table').x;
    return files;
  } finally { doc.free(); }
};
