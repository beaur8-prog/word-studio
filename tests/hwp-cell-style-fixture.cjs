const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const result = [];
  for (const mode of ['bold', 'smaller', 'family']) {
    const source = HwpDocument.createEmpty();
    try {
      source.createBlankDocument();
      const text = Array.from({ length: 12 }, (_, i) => `<span style="font-family:Arial;font-size:10pt">ALPHA${i} </span>` +
        (mode === 'bold' ? '<span style="font-family:Arial;font-size:10pt;font-weight:bold">BOLD </span>' :
          mode === 'smaller' ? '<span style="font-family:Arial;font-size:8pt">small </span>' :
            '<span style="font-family:Times New Roman;font-size:10pt">SERIF </span>') +
        '<span style="font-family:Arial;font-size:10pt">OMEGA</span>').join('<br>');
      source.pasteHtml(0, 0, 0, `<table style="width:480px"><tr><td style="width:480px;height:64px;padding:0;border:none"><p style="font-family:Arial;font-size:10pt;line-height:160%">${text}</p></td></tr></table>`);
      for (const format of ['hwp', 'hwpx']) {
        const file = path.join(directory, `${mode}.${format}`), bytes = format === 'hwp' ? source.exportHwp() : source.exportHwpx();
        await fs.writeFile(file, bytes);
        const native = new HwpDocument(bytes);
        try {
          const runs = JSON.parse(native.getPageTextLayout(0)).runs.filter(r => r.parentParaIdx !== undefined && /ALPHA|OMEGA|BOLD|SERIF|small/.test(r.text));
          const lines = new Set(runs.map(r => r.y));
          if (lines.size !== 12) throw Error('Fixture requires twelve native lines with the same base size on every line');
          result.push({ mode, format, file, nativeLines: lines.size, logicalText: text.replace(/<[^>]*>/g, '').replace(/\s/g, ''),
            nativeHeight: Math.max(...runs.map(r => r.y + r.h)) - Math.min(...runs.map(r => r.y)) });
        } finally { native.free(); }
      }
    } finally { source.free(); }
  }
  return result;
};
