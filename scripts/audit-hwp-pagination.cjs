'use strict';
// Run after test:hwp-samples has downloaded the public corpus. Optional argument:
// path to an unpacked original @rhwp/core 0.8.7 package for before/after evidence.
const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { createHash } = require('node:crypto');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
async function inspect(modulePath, bytes) {
  const { default: init, HwpDocument } = await import(pathToFileURL(modulePath).href);
  await init({ module_or_path: await fs.readFile(path.join(path.dirname(modulePath), 'rhwp_bg.wasm')) });
  const document = new HwpDocument(bytes), characters = new Set(), overflow = [];
  try {
    for (let page = 0; page < document.pageCount(); page++) {
      const svg = document.renderPageSvg(page);
      const height = Number(svg.match(/<svg[^>]*\bheight="([\d.]+)"/)[1]);
      for (const run of JSON.parse(document.getPageTextLayout(page)).runs) {
        if (run.text.trim() && (run.y < -0.5 || run.y + run.h > height + 0.5)) overflow.push(page + 1);
        // Ignore generated page numbers and compare real source positions.
        // Repeated table headers count once, independently of page breaks.
        if (run.paraIdx > 0xfffffff0) continue;
        const location = JSON.stringify([run.secIdx, run.parentParaIdx, run.controlIdx, run.cellPath, run.paraIdx]);
        for (let i = 0; i < run.text.length; i++) if (!/\s/.test(run.text[i])) characters.add(location + ':' + (run.charStart + i) + ':' + run.text[i]);
      }
    }
    return { pages: document.pageCount(), physicalOverflowPages: [...new Set(overflow)], sourceCharacters: characters.size,
      sourceCharacterDigest: createHash('sha256').update([...characters].sort().join('\n')).digest('hex') };
  } finally { document.free(); }
}
async function main() {
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'tests/hwp-samples.json')));
  const sample = manifest.find(row => row.file === '02-work-rules.hwp');
  const bytes = await fs.readFile(path.join(root, 'analysis/hwp-validation/samples', sample.file));
  assert.equal(createHash('sha256').update(bytes).digest('hex'), sample.sha256, 'Use the unchanged public source');
  const after = await inspect(require.resolve('@rhwp/core'), bytes);
  assert.equal(after.pages, 163); assert.deepEqual(after.physicalOverflowPages, []);
  const before = process.argv[2] ? await inspect(path.resolve(process.argv[2], 'rhwp.js'), bytes) : undefined;
  if (before) {
    assert.equal(before.pages, 154);
    assert.equal(after.sourceCharacterDigest, before.sourceCharacterDigest, 'Source characters must not change when repaginating');
  }
  const report = { input: sample.file, sha256: sample.sha256, before, after };
  const output = path.join(root, 'analysis/hwp-pagination'); await fs.mkdir(output, { recursive: true });
  await fs.writeFile(path.join(output, 'engine-audit.json'), JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report, null, 2));
}
main().catch(error => { console.error(error); process.exitCode = 1; });
