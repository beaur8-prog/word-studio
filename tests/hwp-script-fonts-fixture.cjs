const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const JSZip = require('jszip');
module.exports = async function fixture(directory, { hangulFace = 'Arial', latinFace = 'Courier New', latinSize = 100, hangulSpacing = 0, hangulWidth = 100, latinWidth = 100, columns = 1, body = '한글 Latin 123, <b>Bold words</b><br>다음 Next line.' } = {}) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty(); let bytes;
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, `<p></p><p style="font-family:${hangulFace};font-size:12pt;line-height:150%"><span style="font-family:${hangulFace};font-size:12pt">${body}</span></p><p style="font-family:${hangulFace};font-size:12pt">AFTER</p><p></p>`);
    doc.deleteParagraph(0, doc.getParagraphCount(0) - 1); doc.deleteParagraph(0, 0);
    if (columns > 1) doc.setColumnDef(0, columns, 0, 1, 1800);
    const zip = await JSZip.loadAsync(doc.exportHwpx());
    let header = await zip.file('Contents/header.xml').async('string'), fontId;
    header = header.replace(/<hh:fontface\b([^>]*)>([\s\S]*?)<\/hh:fontface>/g, (whole, attributes, children) => {
      if (!/\blang="LATIN"/.test(attributes)) return whole;
      fontId = (children.match(/<hh:font\b/g) || []).length;
      return `<hh:fontface${attributes.replace(/fontCnt="\d+"/, `fontCnt="${fontId + 1}"`)}>${children}<hh:font id="${fontId}" face="${latinFace}" type="TTF" isEmbedded="0"/></hh:fontface>`;
    });
    if (fontId === undefined) throw Error('Missing source Latin font table');
    header = header.replace(/(<hh:fontRef\b[^>]*\blatin=")\d+/g, (_, prefix) => prefix + fontId);
    if (latinSize !== 100) header = header.replace(/(<hh:relSz\b[^>]*\blatin=")100/g, (_, prefix) => prefix + latinSize);
    if (hangulSpacing !== 0) header = header.replace(/(<hh:spacing\b[^>]*\bhangul=")0/g, (_, prefix) => prefix + hangulSpacing);
    header = header.replace(/<hh:ratio\b[^>]*>/g, ratio => ratio.replace(/(hangul|latin)="100"/g,
      (whole, script) => script + '="' + (script === 'hangul' ? hangulWidth : latinWidth) + '"'));
    zip.file('Contents/header.xml', header); bytes = await zip.generateAsync({ type: 'uint8array' });
  } finally { doc.free(); }
  const loaded = new HwpDocument(bytes);
  try {
    const char = JSON.parse(loaded.getCharPropertiesAt(0, 0, 3));
    if (char.fontFamilies[0] !== hangulFace || char.fontFamilies[1] !== latinFace) throw Error('Fixture must have distinct native script faces: ' + JSON.stringify({ fonts: char.fontFamilies, text: loaded.getTextRange(0, 0, 0, 0xffffffff) }));
    if (char.relativeSizes[0] !== 100 || char.relativeSizes[1] !== latinSize) throw Error('Fixture must preserve native script sizes');
    if (char.spacings[0] !== hangulSpacing || char.spacings[1] !== 0) throw Error('Fixture must preserve native script tracking');
    if (char.ratios[0] !== hangulWidth || char.ratios[1] !== latinWidth) throw Error('Fixture must preserve native script widths');
    const files = {};
    for (const format of ['hwp', 'hwpx']) {
      const file = path.join(directory, 'script-fonts.' + format);
      await fs.writeFile(file, format === 'hwp' ? loaded.exportHwp() : loaded.exportHwpx()); files[format] = file;
    }
    return files;
  } finally { loaded.free(); }
};
