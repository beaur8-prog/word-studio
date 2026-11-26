'use strict';

const { createHash } = require('node:crypto');
const { readFile } = require('node:fs/promises');
const { dirname, join } = require('node:path');
const { pathToFileURL } = require('node:url');

const normalized = text => text.replace(/\s/gu, '');
const fingerprint = text => createHash('sha256').update(text).digest('hex');
const coordinate = position => position.map(value => Math.round(value * 1000) / 1000);

async function main() {
  const inventory = JSON.parse(await readFile(process.env.WINDOWS_PUBLIC_QA_INVENTORY, 'utf8'));
  const sourcePath = inventory.find(item => item.verified)?.path;
  if (!sourcePath) throw Error('Fixed source unavailable');
  const engineRoot = dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(join(engineRoot, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(join(engineRoot, 'rhwp_bg.wasm')) });
  const src = new HwpDocument(await readFile(sourcePath));
  const observe = async path => {
    const sheets = JSON.parse(await readFile(path, 'utf8'));
    const paragraphs = new Map();
    for (const {index, paragraphs: entries} of sheets) for (const p of entries) {
      paragraphs.set(p.key, paragraphs.has(p.key) ? null : {page:index, ...p});
    }
    return { sheets, paragraphs };
  };
  try {
    const [control, scaled] = await Promise.all(process.argv.slice(2).map(observe));
    const diffs = [];
    for (let p=0; p<=124; p++) {
      const key = fingerprint(normalized(src.getTextRange(0,p,0,0xffffffff)));
      const a = control.paragraphs.get(key), b = scaled.paragraphs.get(key);
      if (!a || !b) continue;
      const samePage = a.page === b.page;
      const lane = point => Number(point.p[0] > 560);
      const delta = [b.p[1]-a.p[1],b.p[3]-a.p[3]];
      diffs.push({ p, samePage, sameLane:lane(a)===lane(b), before:[a.page,lane(a),...coordinate(a.p)],
        after:[b.page,lane(b),...coordinate(b.p)], delta:coordinate(delta), lines:[a.lines,b.lines] });
    }
    let last;
    const changes = diffs.filter(v => {
      const changed = !last || v.before[0]!==last.before[0] || v.before[1]!==last.before[1] ||
        !v.samePage || !v.sameLane || Math.abs(v.delta[0]-last.delta[0])>.18 || Math.abs(v.delta[1])>.22;
      last=v;
      return changed;
    });
    console.log('WINDOWS_HWP_DENSITY_REFLOW ' + JSON.stringify({ all:diffs.length, changes }));
  } finally { src.free(); }
}

main().catch(error => {
  console.error('WINDOWS_HWP_DENSITY_REFLOW_FAILED ' + error.name);
  process.exitCode = 1;
});
