const { test } = require('node:test'), assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), { pathToFileURL } = require('node:url');
const ts = require('typescript'), { unzipSync } = require('fflate');
const file = path.resolve(__dirname, '../src/hwp/text-navigation.ts');
const compiled = ts.transpileModule(fs.readFileSync(file, 'utf8'), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, esModuleInterop: true },
}).outputText;
const loaded = { exports: {} };
new Function('require', 'module', 'exports', compiled)(name => require(name.startsWith('.') ? path.resolve(path.dirname(file), name) : name), loaded, loaded.exports);
const { hwpTabDefinitions, stageHwpTextNavigation } = loaded.exports;
const core = (async () => {
  const root = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: fs.readFileSync(path.join(root, 'rhwp_bg.wasm')) });
  return HwpDocument;
})();

// Native binary codes are independent of our normalized HWPX name mapping:
// 2 is a dashed leader; 3 is dotted. The publisher's language-plan HWP/PDF
// pair has code 3 in its dotted table of contents. Do not derive this oracle
// from whether our own importer/exporter agrees with itself.
test('binary dotted/dashed tab leaders keep their semantics through normalized HWPX', async () => {
  const Core = await core, source = Core.createEmpty();
  try {
    source.createBlankDocument(); source.insertText(0, 0, 0, 'ITEM\tONE\tTWO');
    source.applyParaFormat(0, 0, JSON.stringify({ tabStops: [
      { position: 15000, type: 0, fill: 3 }, { position: 30000, type: 1, fill: 2 },
    ] }));
    for (const format of ['hwp', 'hwpx']) {
      const doc = new Core(format === 'hwp' ? source.exportHwp() : source.exportHwpx());
      try {
        const properties = JSON.parse(doc.getParaPropertiesAt(0, 0));
        assert.deepEqual(properties.tabStops.map(s => s.fill), [3, 2], format);
        const stops = hwpTabDefinitions(unzipSync(doc.exportHwpx()))(String(properties.paraShapeId), properties.defaultTabSpacing, 400);
        assert.deepEqual(stops.slice(0, 2), [
          { position: 100, type: 'left', leader: 'dot' }, { position: 200, type: 'right', leader: 'hyphen' },
        ], format);
      } finally { doc.free(); }
    }
  } finally { source.free(); }
});

test('saving dotted/dashed editor tabs writes native codes, not just a symmetric private mapping', async () => {
  const Core = await core, source = Core.createEmpty();
  try {
    source.createBlankDocument();
    const stops = [{ position: 100, type: 'left', leader: 'dot' }, { position: 200, type: 'right', leader: 'hyphen' }];
    const staged = stageHwpTextNavigation(`<p data-tab-stops='${JSON.stringify(stops)}'>ITEM\tONE\tTWO</p>`);
    source.pasteHtml(0, 0, 0, staged.html);
    const restored = staged.restore(Core, source);
    try {
      let bytes = restored.exportHwpx();
      for (const format of ['hwp', 'hwpx']) for (let save = 0; save < 2; save++) {
        const doc = new Core(bytes);
        try {
          assert.deepEqual(JSON.parse(doc.getParaPropertiesAt(0, 0)).tabStops.slice(0, 2).map(s => s.fill), [3, 2], `${format} save ${save}`);
          bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx();
        } finally { doc.free(); }
      }
    } finally { if (restored !== source) restored.free(); }
  } finally { source.free(); }
});
