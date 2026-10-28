const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  await fs.mkdir(directory, { recursive: true }); const files = {};
  for (const enabled of [false, true]) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      doc.setPageDef(0, JSON.stringify({ width: 42520, height: 51024, landscape: false, marginTop: 2835, marginBottom: 2835, marginLeft: 2835, marginRight: 2835, marginHeader: 0, marginFooter: 0, marginGutter: 0 }));
      doc.setSectionDef(0, JSON.stringify({ hideEmptyLine: enabled }));
      const lead = Array.from({ length: 30 }, (_, i) => `<p style="font-size:10pt;line-height:20px;margin:0">LEAD ${i}</p>`).join('');
      doc.pasteHtml(0, 0, 0, lead + '<p style="font-size:10pt;line-height:20px;margin:0"></p>'.repeat(3) + '<p style="font-size:10pt;line-height:20px;margin:0">AFTER</p>');
      // The mutation API keeps ParaShape's doubled scale; the read API
      // reports physical pixels (3000 / 150 = 20px).
      for (let p = 0; p < 34; p++) doc.applyParaFormat(0, p, JSON.stringify({ lineSpacingType: 'Fixed', lineSpacing: 3000 }));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), original = new HwpDocument(bytes);
        try {
          if (JSON.parse(original.getSectionDef(0)).hideEmptyLine !== enabled) throw Error('Fixture section policy was lost');
          if (JSON.parse(original.getParaPropertiesAt(0, 0)).lineSpacing !== 20) throw Error('Fixture must have a 20px physical pitch');
          if (original.pageCount() !== 2 || original.getParagraphCount(0) !== 34) throw Error(`Fixture must have two pages and 34 source paragraphs; got ${original.pageCount()}/${original.getParagraphCount(0)}`);
          const file = path.join(directory, `${enabled}.${format}`); await fs.writeFile(file, bytes);
          files[`${format}-${enabled}`] = { path: file, pages: original.pageCount(), paragraphs: original.getParagraphCount(0) };
        } finally { original.free(); }
      }
    } finally { doc.free(); }
  }
  return files;
};
