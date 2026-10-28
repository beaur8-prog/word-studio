const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const files = {};
  for (const nested of [false, true]) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      const inner = nested ? '<table style="width:200px"><tr><td style="padding:0"><p>NESTED</p></td></tr></table>' : '';
      doc.pasteHtml(0, 0, 0, `<table style="width:400px"><tr><td style="padding:0"><p style="line-height:24px">BEFORE</p>${inner}<p style="line-height:40px"></p><p style="text-align:right;line-height:24px">AFTER</p></td><td style="padding:0"><p></p></td></tr></table>`);
      for (const format of ['hwp', 'hwpx']) {
        const key = format + (nested ? '-nested' : '');
        files[key] = path.join(directory, key + '.' + format);
        await fs.writeFile(files[key], format === 'hwp' ? doc.exportHwp() : doc.exportHwpx());
      }
    } finally { doc.free(); }
  }
  return files;
};
