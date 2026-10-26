const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const { createRequire } = require('node:module');
const ts = require('typescript');
const createFixture = require('./hwp-layout-fixture.cjs');
const setup = (async () => {
  const source = path.resolve(__dirname, '../src/hwp/layout-metadata.ts'), module = { exports: {} };
  new Function('require', 'module', 'exports', ts.transpileModule(await fs.readFile(source, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, esModuleInterop: true },
  }).outputText)(createRequire(source), module, module.exports);
  return module.exports;
})();
test('original HWP and HWPX paragraph headers retain explicit column/page breaks and physical column gaps', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-layout-'));
  try {
    const files = await createFixture(dir), { readHwpLayoutMetadata } = await setup;
    for (const format of ['hwp', 'hwpx']) {
      const result = await readHwpLayoutMetadata(await fs.readFile(files[format]), format === 'hwpx');
      assert.deepEqual(result.warnings, [], format);
      assert.equal(Object.keys(result.paragraphs).length, 3, format);
      assert.deepEqual(result.paragraphs['0:0'].columns, { count: 2, space: 24 }, format);
      assert.equal(result.paragraphs['0:1'].columnBreak, true, format);
      assert.equal(result.paragraphs['0:2'].pageBreak, true, format);
    }
  } finally { await fs.rm(dir, { recursive: true, force: true }); }
});
test('nested paragraphs and fake XML inside comments do not change the main paragraph index', async () => {
  const { parseHwpxSection } = await setup;
  const result = parseHwpxSection('<s:sec><p:p pageBreak="0"><p:colPr colCount="2" sameGap="1500"/><p:p pageBreak="1"/></p:p><!-- <p:p pageBreak="1"/> --><p:p columnBreak="1"/></s:sec>', 0);
  assert.deepEqual(Object.keys(result), ['0:0', '0:1']);
  assert.equal(result['0:0'].pageBreak, false); assert.equal(result['0:1'].columnBreak, true);
});
test('truncated binary metadata fails safely, without guessing paragraph indices', async () => {
  const { parseHwpSection, readHwpLayoutMetadata } = await setup;
  assert.throws(() => parseHwpSection(new Uint8Array([66, 0, 16, 0]), 0), /Truncated/);
  const fallback = await readHwpLayoutMetadata(new Uint8Array([1, 2, 3]));
  assert.deepEqual(fallback.paragraphs, {}); assert.equal(fallback.warnings.length, 1);
});
