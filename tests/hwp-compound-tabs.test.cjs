const { test } = require('node:test'), assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const ts = require('typescript'), { strToU8, unzipSync } = require('fflate');
const file = path.resolve(__dirname, '../src/hwp/text-navigation.ts');
const source = ts.transpileModule(fs.readFileSync(file, 'utf8'), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, esModuleInterop: true },
}).outputText;
const loaded = { exports: {} };
new Function('require', 'module', 'exports', source)(name => require(name.startsWith('.') ? path.resolve(path.dirname(file), name) : name), loaded, loaded.exports);
const { hwpTabDefinitions } = loaded.exports;

test('each normalized switch contributes one tab without mixing HWPUNIT cases', () => {
  const xml = '<hh:head><hh:tabPr id="7" autoTabRight="0">' +
    '<hp:switch><hp:case><hh:tabItem pos="7500" type="LEFT" unit="HWPUNIT"/></hp:case><hp:default><hh:tabItem pos="15000" type="LEFT"/></hp:default></hp:switch>' +
    '<hp:switch><hp:case><hh:tabItem pos="13500" type="RIGHT" unit="HWPUNIT"/></hp:case><hp:default><hh:tabItem pos="27000" type="RIGHT" leader="DASH"/></hp:default></hp:switch>' +
    '</hh:tabPr><hh:paraPr id="3" tabPrIDRef="7"/></hh:head>';
  const stops = hwpTabDefinitions({ 'Contents/header.xml': strToU8(xml) })('3', 8000, 400);
  assert.deepEqual(stops.slice(0, 2), [{ position: 100, type: 'left' }, { position: 180, type: 'right', leader: 'dot' }]);
  assert.ok(stops.slice(2).every(s => s.position > 180));
});

test('unswitched stops and each independent default branch stay in source order', () => {
  const xml = '<hh:head><hh:tabPr id="7"><hh:tabItem pos="9000" type="CENTER"/>' +
    '<hp:switch><hp:default><hh:tabItem pos="18000" type="DECIMAL" leader="DOT"/></hp:default></hp:switch>' +
    '<hp:switch><hp:case><hh:tabItem pos="14000" type="RIGHT" unit="HWPUNIT"/></hp:case></hp:switch>' +
    '<hh:tabItem pos="36000" type="RIGHT" leader="SOLID"/></hh:tabPr><hh:paraPr id="3" tabPrIDRef="7"/></hh:head>';
  assert.deepEqual(hwpTabDefinitions({ 'Contents/header.xml': strToU8(xml) })('3', 8000, 400).slice(0, 3), [
    { position: 60, type: 'center' }, { position: 120, type: 'decimal', leader: 'hyphen' },
    { position: 240, type: 'right', leader: 'underscore' },
  ]);
});

test('a default grid stop beyond the text width does not erase usable explicit stops', () => {
  const xml = '<hh:head><hh:tabPr id="7"><hh:tabItem pos="15000" type="LEFT"/>' +
    '<hh:tabItem pos="27000" type="RIGHT" leader="DASH"/><hh:tabItem pos="64000" type="LEFT"/>' +
    '</hh:tabPr><hh:paraPr id="3" tabPrIDRef="7"/></hh:head>';
  const stops = hwpTabDefinitions({ 'Contents/header.xml': strToU8(xml) })('3', 8000, 400);
  assert.deepEqual(stops.slice(0, 2), [{ position: 100, type: 'left' }, { position: 180, type: 'right', leader: 'dot' }]);
  assert.ok(stops.slice(2).every(s => s.position > 180));
});

test('actual native normalization keeps two compound stops for HWP and HWPX', async () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'word-compound-tabs-unit-'));
  try {
    const files = await require('./hwp-compound-tabs-fixture.cjs')(directory);
    const root = path.dirname(require.resolve('@rhwp/core'));
    const { HwpDocument } = await import(require('node:url').pathToFileURL(path.join(root, 'rhwp.js')).href);
    for (const format of ['hwp', 'hwpx']) {
      const doc = new HwpDocument(fs.readFileSync(files[format]));
      try {
        const properties = JSON.parse(doc.getParaPropertiesAt(0, 0));
        assert.equal(properties.tabStops.length, 2);
        const stops = hwpTabDefinitions(unzipSync(doc.exportHwpx()))(String(properties.paraShapeId), properties.defaultTabSpacing, 400);
        assert.deepEqual(stops.slice(0, 2), [{ position: 100, type: 'left' }, { position: 180, type: 'right', leader: 'dot' }], format);
      } finally { doc.free(); }
    }
  } finally { fs.rmSync(directory, { recursive: true, force: true }); }
});
