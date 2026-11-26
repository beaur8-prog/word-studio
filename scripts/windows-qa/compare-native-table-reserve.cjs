'use strict';

const { readFile } = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const compact = value => value.replace(/\s/gu,'');
const same = (a,b) => JSON.stringify(a) === JSON.stringify(b);
function inspect(doc) {
  const logical = Array.from({length:doc.getSectionCount()},(_,sec)=>Array.from(
    {length:doc.getParagraphCount(sec)},(_,para)=>doc.getTextRange(sec,para,0,0xffffffff)));
  const owners = new Map(), objects = new Map();
  for (let page=0;page<doc.pageCount();page++) {
    for (const run of JSON.parse(doc.getPageTextLayout(page)).runs) {
      if (run.headerFooter || run.paraIdx===0xffffffff || !compact(run.text)) continue;
      const owner = JSON.stringify([run.secIdx,run.parentParaIdx,run.paraIdx,run.cellPath]);
      owners.set(owner,(owners.get(owner) || '')+compact(run.text));
    }
    for (const box of JSON.parse(doc.getPageControlLayout(page)).controls) {
      if (box.headerFooter) continue;
      const id = JSON.stringify([box.type,box.secIdx,box.paraIdx,box.controlIdx,box.cellPath]);
      const list = objects.get(id) ?? [];
      list.push([box.w,box.h]);
      objects.set(id,list);
    }
  }
  return {logical,owners,objects};
}
async function main() {
  const root=path.dirname(require.resolve('@rhwp/core'));
  const {default:init,HwpDocument}=await import(pathToFileURL(path.join(root,'rhwp.js')).href);
  await init({module_or_path:await readFile(path.join(root,'rhwp_bg.wasm'))});
  const [before,after]=await Promise.all(process.argv.slice(2).map(async file=>new HwpDocument(await readFile(file))));
  try {
    const a=inspect(before), b=inspect(after);
    const onlyDigits=[];
    const changed = [];
    for (const key of new Set([...a.owners.keys(),...b.owners.keys()])) {
      const x=a.owners.get(key) ?? '', y=b.owners.get(key) ?? '';
      if (x===y) continue;
      if (x.replace(/\p{Nd}/gu,'')===y.replace(/\p{Nd}/gu,'') && x.length===y.length) onlyDigits.push([key,x.length]);
      else changed.push([JSON.parse(key),x.length,y.length,[...x].sort().join('')===[...y].sort().join('')]);
    }
    const exact = same(a.logical,b.logical);
    const sameObjectSizes = same([...a.objects].sort(),[...b.objects].sort());
    console.log('WINDOWS_NATIVE_RESERVE_COMPARE ' + JSON.stringify({ pages:[before.pageCount(),after.pageCount()],
      logicalExact:exact, sameObjectSizes, owners:[a.owners.size,b.owners.size], numericOwners:onlyDigits.length, changed }));
    if (!exact || !sameObjectSizes || changed.length > 0) process.exitCode = 1;
  } finally {before.free();after.free();}
}
main().catch(()=>{console.log('WINDOWS_NATIVE_RESERVE_COMPARE_FAILED');process.exitCode=1;});
