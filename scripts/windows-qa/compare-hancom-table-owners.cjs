'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const audit = require('./hancom-table-owner-audit.cjs');

async function main() {
  const [control, candidate] = process.argv.slice(2);
  if (!control || !candidate) throw Error('Two actual Hancom output directories are required');
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')));
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  for (const format of ['hwp','hwpx']) {
    const before = new HwpDocument(await fs.readFile(path.join(control, 'hancom-saved.' + format)));
    const after = new HwpDocument(await fs.readFile(path.join(candidate, 'hancom-saved.' + format)));
    try {
      const result = audit(before, after);
      console.log('WINDOWS_NATIVE_TABLE_OWNER_AUDIT ' + JSON.stringify({ format, ...result }));
      if (!result.bodySequenceExact || !result.sameCellMultiset || !result.tableGlyphSequenceExact) process.exitCode = 1;
    } finally {
      before.free();
      after.free();
    }
  }
}
main().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
