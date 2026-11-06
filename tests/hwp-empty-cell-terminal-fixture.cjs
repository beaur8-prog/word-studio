const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url'), { createRequire } = require('node:module'), ts = require('typescript'), { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
function load(file, cache = new Map()) {
  if (cache.has(file)) return cache.get(file).exports;
  const module = { exports: {} }; cache.set(file, module);
  const native = createRequire(file), code = ts.transpileModule(require('node:fs').readFileSync(file, 'utf8'), { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS, esModuleInterop: true } }).outputText;
  new Function('require', 'module', 'exports', code)(name => name.startsWith('.') && !path.extname(name) ? load(path.resolve(path.dirname(file), name + '.ts'), cache) : native(name), module, module.exports);
  return module.exports;
}
module.exports = async function fixture() {
  const root = path.dirname(require.resolve('@rhwp/core')), { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')));
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const { editableHwp } = load(path.join(__dirname, '../src/hwp/conversion.ts')), result = {};
  for (const mode of ['empty', 'space', 'two', 'break']) for (const nested of [false, true]) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      const blank = `<p style="font-family:Arial;font-size:10pt;line-height:300%;margin:0">EMPTY-TERMINAL${mode === 'break' ? '<br>' : ''}</p>`;
      const table = `<table style="width:240px"><tr><td style="width:120px;height:1px;padding:0;border:none"><p>MARKER</p></td><td style="width:120px;height:1px;padding:0;border:none">${blank}${mode === 'two' ? blank : ''}</td></tr></table>`;
      doc.pasteHtml(0, 0, 0, '<p>BEFORE</p>' + (nested ? `<table style="width:300px"><tr><td style="padding:0;border:none">${table}</td></tr></table>` : table) + '<p>AFTER</p>');
      const files = unzipSync(doc.exportHwpx());
      files['Contents/section0.xml'] = strToU8(strFromU8(files['Contents/section0.xml']).replace(/EMPTY-TERMINAL/g, mode === 'space' ? ' ' : ''));
      const hwpx = zipSync(files, { level: 1 }), source = new HwpDocument(hwpx);
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? source.exportHwp() : hwpx, reopened = new HwpDocument(bytes);
        try { result[`${mode}-${nested ? 'nested' : 'ordinary'}-${format}`] = { bytes, editable: editableHwp(reopened), nativePages: reopened.pageCount() }; }
        finally { reopened.free(); }
      }
      source.free();
    } finally { doc.free(); }
  }
  return result;
};
