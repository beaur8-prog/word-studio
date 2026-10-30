const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const JSZip = require('jszip');
const createBodyFixture = require('./hwp-script-fonts-fixture.cjs');
module.exports = async function fixture(directory) {
  const source = await createBodyFixture(directory), core = path.dirname(require.resolve('@rhwp/core'));
  const { HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  const doc = new HwpDocument(await fs.readFile(source.hwpx)); let bytes;
  try {
    const p = doc.getParagraphCount(0) - 1;
    const span = text => `<span style="font-family:Arial;font-size:12pt">${text}</span>`;
    doc.pasteHtml(0, p, doc.getParagraphLength(0, p), `<p>😀 ${span('After astral body')}</p><table style="width:500px"><tr><td><p>${span('셀 😀 Table Latin 456')}</p><p>${span('추가 Second cell line')}</p></td><td><p>${span('중첩 Outer text')}</p><table style="width:200px"><tr><td><p>${span('안쪽 😀 Nested Latin 789')}</p></td></tr></table><p>${span('표뒤 After nested table')}</p></td></tr></table><p></p>`);
    const zip = await JSZip.loadAsync(doc.exportHwpx());
    let header = await zip.file('Contents/header.xml').async('string');
    const latin = header.match(/<hh:fontface\b[^>]*\blang="LATIN"[^>]*>([\s\S]*?)<\/hh:fontface>/)?.[1];
    const fontId = latin?.match(/<hh:font\b[^>]*\bid="(\d+)"[^>]*\bface="Courier New"/)?.[1];
    if (fontId === undefined) throw Error('Missing source Courier New face');
    header = header.replace(/(<hh:fontRef\b[^>]*\blatin=")\d+/g, (_, prefix) => prefix + fontId);
    zip.file('Contents/header.xml', header); bytes = await zip.generateAsync({ type: 'uint8array' });
  } finally { doc.free(); }
  const loaded = new HwpDocument(bytes);
  try {
    const files = {};
    for (const format of ['hwp', 'hwpx']) {
      const file = path.join(directory, 'script-cells.' + format);
      await fs.writeFile(file, format === 'hwp' ? loaded.exportHwp() : loaded.exportHwpx()); files[format] = file;
    }
    return files;
  } finally { loaded.free(); }
};
