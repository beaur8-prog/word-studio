const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const JSZip = require('jszip');
module.exports = async function fixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty(); let bytes;
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p></p><p style="font-family:Arial;font-size:12pt;line-height:150%"><span style="font-family:Arial;font-size:12pt">한글 Latin 123, <b>Bold words</b><br>다음 Next line.</span></p><p style="font-family:Arial;font-size:12pt">AFTER</p><p></p>');
    doc.deleteParagraph(0, doc.getParagraphCount(0) - 1); doc.deleteParagraph(0, 0);
    const zip = await JSZip.loadAsync(doc.exportHwpx());
    let header = await zip.file('Contents/header.xml').async('string'), fontId;
    header = header.replace(/<hh:fontface\b([^>]*)>([\s\S]*?)<\/hh:fontface>/g, (whole, attributes, children) => {
      if (!/\blang="LATIN"/.test(attributes)) return whole;
      fontId = (children.match(/<hh:font\b/g) || []).length;
      return `<hh:fontface${attributes.replace(/fontCnt="\d+"/, `fontCnt="${fontId + 1}"`)}>${children}<hh:font id="${fontId}" face="Courier New" type="TTF" isEmbedded="0"/></hh:fontface>`;
    });
    if (fontId === undefined) throw Error('Missing source Latin font table');
    header = header.replace(/(<hh:fontRef\b[^>]*\blatin=")\d+/g, (_, prefix) => prefix + fontId);
    zip.file('Contents/header.xml', header); bytes = await zip.generateAsync({ type: 'uint8array' });
  } finally { doc.free(); }
  const loaded = new HwpDocument(bytes);
  try {
    const char = JSON.parse(loaded.getCharPropertiesAt(0, 0, 3));
    if (char.fontFamilies[0] !== 'Arial' || char.fontFamilies[1] !== 'Courier New') throw Error('Fixture must have distinct native script faces: ' + JSON.stringify({ fonts: char.fontFamilies, text: loaded.getTextRange(0, 0, 0, 0xffffffff) }));
    const files = {};
    for (const format of ['hwp', 'hwpx']) {
      const file = path.join(directory, 'script-fonts.' + format);
      await fs.writeFile(file, format === 'hwp' ? loaded.exportHwp() : loaded.exportHwpx()); files[format] = file;
    }
    return files;
  } finally { loaded.free(); }
};
