'use strict';
// Export the original source stories independently of the editable converter.
// Usage: node scripts/export-hwp-audit-source.cjs [corpus-dir] [output-dir]
const fs = require('node:fs/promises');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const JSZip = require('jszip');
(async () => {
  const corpus = path.resolve(process.argv[2] || 'analysis/hwp-row-flow-2026-10-26');
  const output = path.resolve(process.argv[3] || 'analysis/hwp-completion-fixes-2026-10-26');
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  const manifest = JSON.parse(await fs.readFile(path.join(corpus, 'manifest.json'), 'utf8'));
  const sections = {}, hashes = {};
  for (const sample of manifest.samples) {
    const bytes = await fs.readFile(path.join(corpus, sample.hwp.file));
    const hash = createHash('sha256').update(bytes).digest('hex');
    if (hash !== sample.hwp.sha256) throw Error(`Source hash mismatch: ${sample.id}`);
    const doc = new HwpDocument(bytes);
    try {
      const zip = await JSZip.loadAsync(doc.exportHwpx()); sections[sample.id] = {}; hashes[sample.id] = hash;
      for (const name of Object.keys(zip.files).filter(n => /^Contents\/section\d+\.xml$/.test(n)).sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0])))
        sections[sample.id][name] = await zip.file(name).async('string');
    } finally { doc.free(); }
  }
  await fs.mkdir(output, { recursive: true });
  await fs.writeFile(path.join(output, 'source-sections.json'), JSON.stringify(sections));
  await fs.writeFile(path.join(output, 'source-audit-hashes.json'), JSON.stringify(hashes, null, 2) + '\n');
  console.log(`Exported ${Object.keys(sections).length} original story archives with verified SHA-256.`);
})().catch(error => { console.error(error); process.exitCode = 1; });
