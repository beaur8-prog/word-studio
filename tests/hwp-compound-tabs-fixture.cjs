const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');

module.exports = async function fixture(directory) {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const doc = HwpDocument.createEmpty(); let bytes;
  try {
    doc.createBlankDocument();
    doc.pasteHtml(0, 0, 0, '<p style="font-family:Arial;font-size:12pt">BODY\tONE\tTWO</p>' +
      '<table style="width:500px"><tr><td style="padding:4px"><p style="font-family:Arial;font-size:12pt">CELL\tONE\tTWO</p>' +
      '<table style="width:350px"><tr><td style="padding:4px"><p style="font-family:Arial;font-size:12pt">NEST\tONE\tTWO</p></td></tr></table>' +
      '</td></tr></table><p>AFTER</p>');
    doc.applyParaFormat(0, 0, JSON.stringify({ tabStops: [
      { position: 15000, type: 0, fill: 0 }, { position: 27000, type: 1, fill: 2 },
    ] }));
    const properties = JSON.parse(doc.getParaPropertiesAt(0, 0));
    if (properties.tabStops.length !== 2) throw Error('Native fixture must retain both source stops');
    const files = unzipSync(doc.exportHwpx());
    let xml = strFromU8(files['Contents/section0.xml']);
    const stack = [], edits = [];
    for (const m of xml.matchAll(/<hp:p\b[^>]*>|<\/hp:p>|\b(?:CELL|NEST)\b/g)) {
      if (m[0].startsWith('<hp:p')) stack.push(m);
      else if (m[0] === '</hp:p>') stack.pop();
      else {
        const p = stack.at(-1); if (!p) throw Error('Missing source cell paragraph');
        edits.push({ from: p.index, to: p.index + p[0].length,
          value: p[0].replace(/paraPrIDRef="[^"]*"/, `paraPrIDRef="${properties.paraShapeId}"`) });
      }
    }
    if (edits.length !== 2) throw Error('Fixture needs ordinary and nested cell text');
    for (const e of edits.reverse()) xml = xml.slice(0, e.from) + e.value + xml.slice(e.to);
    files['Contents/section0.xml'] = strToU8(xml.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, ''));
    bytes = zipSync(files);
  } finally { doc.free(); }
  const loaded = new HwpDocument(bytes), files = {};
  try {
    for (const format of ['hwp', 'hwpx']) {
      const file = path.join(directory, 'compound-tabs.' + format);
      await fs.writeFile(file, format === 'hwp' ? loaded.exportHwp() : loaded.exportHwpx()); files[format] = file;
    }
  } finally { loaded.free(); }
  return files;
};
