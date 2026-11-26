'use strict';
const { readFile } = require('node:fs/promises');
const { pathToFileURL } = require('node:url');
const path = require('node:path');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');

const shape = (doc,page) => JSON.parse(doc.getPageControlLayout(page)).controls.filter(c=>!c.headerFooter);
const find = doc => {
  for (let page=0; page<doc.pageCount(); page++) {
    const table = shape(doc,page).find(c=>c.type==='table' && c.rowCount===1 && c.colCount===2 && !c.cellPath?.length);
    if (table) return {page,table};
  }
};
function summary(doc) {
  const {page,table} = find(doc);
  const props = JSON.parse(doc.getTableProperties(table.secIdx,table.paraIdx,table.controlIdx));
  const para = JSON.parse(doc.getParaPropertiesAt(table.secIdx,table.paraIdx));
  const {lines,stored} = JSON.parse(doc.getParagraphLineMetrics(table.secIdx,table.paraIdx,'[]'));
  const chars = Array.from({length:doc.pageCount()},(_,page)=>JSON.parse(doc.getPageTextLayout(page)).runs)
    .flat().filter(run=>!run.headerFooter && run.paraIdx!==0xffffffff)
    .map(run=>run.text.replace(/\s/gu,'')).join('');
  return {pages:doc.pageCount(),page:page+1,props:[props.outerTop,props.outerBottom,props.outerLeft,props.outerRight],
    break:para.pageBreakBefore,box:[table.x,table.y,table.w,table.h],charCount:chars.length,lines,stored};
}
function nth(xml,name,ordinal) {
  const stack=[];
  let index=0;
  for (const match of xml.matchAll(new RegExp('<\\/?'+name+'\\b[^>]*>','g'))) {
    if (match[0].startsWith('</')) {
      const open=stack.pop();
      if (open&&!stack.length&&index++===ordinal) return {from:open.index+open[0].length,end:match.index};
    } else if(!match[0].endsWith('/>')) stack.push(match);
  }
}
async function main() {
  const dir=path.dirname(require.resolve('@rhwp/core'));
  const {default:init,HwpDocument}=await import(pathToFileURL(path.join(dir,'rhwp.js')));
  await init({module_or_path:await readFile(path.join(dir,'rhwp_bg.wasm'))});
  const [hwp,hwpx]=await Promise.all(process.argv.slice(2).map(async file=>new HwpDocument(await readFile(file))));
  try {
    const effects = {};
    const repack = new HwpDocument(hwp.exportHwpx());
    try {effects.binaryAsPackage=summary(repack);} finally{repack.free();}
    const {table}=find(hwpx), file='Contents/section'+table.secIdx+'.xml';
    const parts=unzipSync(await readFile(process.argv[3])), xml=strFromU8(parts[file]);
    const owner=nth(xml,'hp:p',table.paraIdx);
    const inner=xml.slice(owner.from,owner.end);
    const tag=/<hp:outMargin\b[^>]*\/>/.exec(inner);
    if(!tag)throw Error('Missing wrapper margin');
    for(const [trial,values] of [
      ['authoredTop',{top:140}],['authoredOuter',{top:140,bottom:140,left:140,right:140}],
      ['zeroTop',{top:0}],['zeroVertical',{top:0,bottom:0}],
    ]){
      const at=owner.from+tag.index;
      const margin=Object.entries(values).reduce((tag,[key,val])=>tag.replace(new RegExp('\\b'+key+'="[^"]*"'),key+'="'+val+'"'),tag[0]);
      const output=new HwpDocument(zipSync({...parts,[file]:strToU8(xml.slice(0,at)+margin+xml.slice(at+tag[0].length))},{level:1}));
      try{effects[trial]=summary(output);}finally{output.free();}
    }
    console.log('WINDOWS_FORM_NATIVE_WRITER_FIT '+JSON.stringify({originalBinary:summary(hwp),originalPackage:summary(hwpx),effects}));
  } finally {hwp.free();hwpx.free();}
}
main().catch(()=>{console.log('WINDOWS_FORM_NATIVE_WRITER_FAILED');process.exitCode=1;});
