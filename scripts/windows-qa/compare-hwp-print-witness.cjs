'use strict';

const { readFile, writeFile } = require('node:fs/promises');
const { createHash } = require('node:crypto');
const { pathToFileURL } = require('node:url');
const { dirname, join } = require('node:path');

const noSpace = value => value.replace(/\s/gu, '');
const keyFor = value => createHash('sha256').update(value).digest('hex');

function bodyOwners(doc) {
  const owners = new Map();
  for (let index = 0; index < doc.pageCount(); index++) {
    const { runs } = JSON.parse(doc.getPageTextLayout(index));
    const { columns } = JSON.parse(doc.getPageInfo(index));
    for (const run of runs) {
      if (run.headerFooter || run.cellPath?.length || !Number.isInteger(run.paraIdx) || run.paraIdx === 0xffffffff) continue;
      const text = noSpace(run.text);
      if (!text) continue;
      const key = [run.secIdx, run.paraIdx].join(':');
      const lane = columns.findIndex(column => run.x >= column.x - 1 && run.x < column.x + column.width + 1);
      const owner = owners.get(key) ?? { section:run.secIdx, paragraph:run.paraIdx, page:index+1, x:run.x, y:run.y,
        lane, columns, text:'', lines:new Set(), lineBands:[] };
      const active = owner.lineBands.at(-1);
      const overlap = active && Math.min(active.bottom, run.y + run.h) - Math.max(active.top, run.y);
      if (!active || index !== active.page || lane !== active.lane || overlap < Math.min(active.bottom-active.top, run.h)*.5)
        owner.lineBands.push({page:index,lane,top:run.y,bottom:run.y+run.h,start:owner.text.length});
      owner.text += text;
      owner.lines.add(index + ':' + lane + ':' + run.y);
      owners.set(key, owner);
    }
  }
  return [...owners.values()];
}

async function main() {
  const inventory = JSON.parse(await readFile(process.env.WINDOWS_PUBLIC_QA_INVENTORY, 'utf8'));
  const inputs = inventory.filter(item => item.verified);
  if (inputs.length !== 1) throw Error('Expected one pinned HWP source');
  const engineRoot = dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(join(engineRoot, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(join(engineRoot, 'rhwp_bg.wasm')) });
  const source = new HwpDocument(await readFile(inputs[0].path));
  try {
    const sheets = JSON.parse(await readFile(process.env.WINDOWS_NATIVE_PRINT_GEOMETRY, 'utf8'));
    const printed = new Map();
    for (const sheet of sheets) for (const paragraph of sheet.paragraphs) {
      printed.set(paragraph.key, printed.has(paragraph.key) ? null : { ...paragraph, page:sheet.index });
    }
    const records = bodyOwners(source).filter(value => value.text.length >= 10);
    const totals = new Map();
    for (const value of records) totals.set(value.text, (totals.get(value.text) ?? 0) + 1);
    const fixed = records.filter(value => totals.get(value.text) === 1);
    const differences = [];
    let matched = 0, pageStarts = 0, laneStarts = 0;
    const firstComparison = [];
    const scored = [];
    const exactLines = { eligible:0, sameCount:0, sameBreaks:0, sameStartAndBreaks:0 };
    for (const before of fixed) {
      const after = printed.get(keyFor(before.text));
      if (!after) continue;
      matched++;
      const lane = before.columns.findIndex(column => after.p[0] >= column.x - 2 && after.p[0] < column.x + column.width + 2);
      const samePage = before.page === after.page;
      const sameLane = before.lane === lane;
      pageStarts += Number(samePage);
      laneStarts += Number(samePage && sameLane);
      if (after.lineStarts && before.lineBands.length) {
        exactLines.eligible++;
        const nativeOffsets = before.lineBands.map(band => band.start);
        exactLines.sameCount += Number(nativeOffsets.length === after.lineStarts.length);
        const sameBreaks = JSON.stringify(nativeOffsets) === JSON.stringify(after.lineStarts);
        exactLines.sameBreaks += Number(sameBreaks);
        exactLines.sameStartAndBreaks += Number(sameBreaks && samePage && sameLane);
        scored.push({ id: before.section + ':' + before.paragraph, sameCount: nativeOffsets.length === after.lineStarts.length,
          sameBreaks, sameStartAndBreaks: sameBreaks && samePage && sameLane });
      }
      const record = { p:before.paragraph, length:before.text.length,
        source:[before.page,before.lane,before.x,before.y,before.lines.size],
        studio:[after.page,lane,after.p[0],after.y[0],after.lines], font:after.font,
        size:after.size, lineHeight:after.lineHeight };
      if ((!samePage || !sameLane) && differences.length < 24) differences.push(record);
      if (before.section === 0 && before.paragraph >= 25 && before.paragraph <= 37) firstComparison.push(record);
    }
    console.log('WINDOWS_HWP_PRINT_WITNESS ' + JSON.stringify({ originalPages:source.pageCount(), studioPages:sheets.length,
      fixed:fixed.length, matched,pageStarts,laneStarts,differences,firstComparison }));
    if (process.env.WINDOWS_CAPTURE_HWP_BODY_LINES === '1')
      console.log('WINDOWS_HWP_EXACT_PRINT_LINES ' + JSON.stringify({ fixed:fixed.length, matched,pageStarts,laneStarts, ...exactLines }));
    if (process.env.WINDOWS_HWP_LINE_VERDICT)
      await writeFile(process.env.WINDOWS_HWP_LINE_VERDICT, JSON.stringify(scored));
  } finally { source.free(); }
}

main().catch(error => {
  console.error('WINDOWS_HWP_PRINT_WITNESS_FAILED ' + error.name);
  process.exitCode = 1;
});
