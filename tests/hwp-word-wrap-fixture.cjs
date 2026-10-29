const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const JSZip = require('jszip');

const text = '가나다라마바 사아자차카타 파하거너더러 머버서어저처 가나다라마바 사아자차카타';
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const result = {};
  for (const context of ['body', 'ordinary', 'nested']) {
    const source = HwpDocument.createEmpty();
    try {
      source.createBlankDocument();
      const paragraph = `<p style="font-family:'Noto Sans KR';font-size:12pt;line-height:160%">${text}</p>`;
      const table = `<table style="width:140px"><tr><td style="width:140px;padding:0;border:none">${paragraph}</td></tr></table>`;
      const html = context === 'body' ? paragraph : context === 'ordinary' ? table :
        `<table style="width:180px"><tr><td style="width:180px;padding:0;border:none"><p>OUTER</p>${table}</td></tr></table>`;
      source.pasteHtml(0, 0, 0, '<p></p>' + html);
      if (!source.getParagraphLength(0, 0)) source.deleteParagraph(0, 0);
      source.setPageDef(0, JSON.stringify({ width: 18500, height: 50000, marginLeft: 1000, marginRight: 1000,
        marginTop: 1000, marginBottom: 1000, marginHeader: 0, marginFooter: 0 }));
      if (context === 'body') source.applyCharFormat(0, 0, 0, 0xffffffff, JSON.stringify({ fontSize: 1200, fontFamily: 'Noto Sans KR' }));
      for (const unit of [0, 1]) {
        const zip = await JSZip.loadAsync(source.exportHwpx());
        const header = await zip.file('Contents/header.xml').async('string');
        // Non-Latin names differ from the Latin field: verified native unit 0
        // keeps Korean words together and serializes as BREAK_WORD.
        zip.file('Contents/header.xml', header.replace(/breakNonLatinWord="[^"]*"/g, `breakNonLatinWord="${unit ? 'KEEP_WORD' : 'BREAK_WORD'}"`));
        const native = new HwpDocument(await zip.generateAsync({ type: 'uint8array' }));
        try {
          const outer = context === 'body' ? undefined : JSON.parse(native.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
          const cells = outer ? [{ controlIndex: outer.controlIndex, cellIndex: 0, cellParaIndex: context === 'nested' ? 1 : 0 },
            ...(context === 'nested' ? [{ controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }] : [])] : [];
          const properties = context === 'body' ? JSON.parse(native.getParaPropertiesAt(0, 0)) :
            JSON.parse(native.getCellParaPropertiesAtByPath(0, 0, JSON.stringify(cells)));
          if (properties.koreanBreakUnit !== unit) throw Error(`Native ${context} fixture requested ${unit}, got ${properties.koreanBreakUnit}`);
          for (const format of ['hwp', 'hwpx']) {
            const key = `${context}-${unit}-${format}`, file = path.join(directory, key + '.' + format);
            await fs.writeFile(file, format === 'hwp' ? native.exportHwp() : native.exportHwpx());
            result[key] = { file, text, unit };
          }
        } finally { native.free(); }
      }
    } finally { source.free(); }
  }
  return result;
};
