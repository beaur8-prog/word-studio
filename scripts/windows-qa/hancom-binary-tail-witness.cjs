'use strict';
const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url');

const body=(doc,page)=>JSON.parse(doc.getPageTextLayout(page)).runs.filter(run=>!run.headerFooter&&
  run.paraIdx!==0xffffffff&&!run.cellPath?.length&&run.text.trim());
const controls=(doc,page)=>JSON.parse(doc.getPageControlLayout(page)).controls.filter(c=>!c.headerFooter);
const sameOwner=(a,b)=>a.secIdx===b.secIdx&&a.paraIdx===b.paraIdx;
const region=(doc,page)=>{
  const {marginTop,marginHeader,footerArea,bodyLeft,bodyRight,height}=JSON.parse(doc.getPageInfo(page));
  return {top:marginTop+marginHeader,bottom:footerArea.y,bodyLeft,bodyRight,height};
};
const ink=r=>[r.charStart,r.text.replace(/\s/gu,'').length,r.x,r.y,r.w,r.h];
const signature=(doc,ref)=>JSON.parse(doc.getParagraphLineMetrics(ref.secIdx,ref.paraIdx,'[]'));

async function main(){
 const dir=path.dirname(require.resolve('@rhwp/core'));
 const {default:init,HwpDocument}=await import(pathToFileURL(path.join(dir,'rhwp.js')));
 await init({module_or_path:await fs.readFile(path.join(dir,'rhwp_bg.wasm'))});
 const [bin,zip]=await Promise.all(process.argv.slice(2).map(async f=>new HwpDocument(new Uint8Array(await fs.readFile(f)))));
 try{
  const crossings=[];
  for(let page=1;page<bin.pageCount();page++){
    const tail=body(bin,page), prev=body(bin,page-1), first=tail[0];
    if(!first||!prev.length||controls(bin,page).length||!tail.every(run=>sameOwner(first,run))||
       !prev.some(run=>sameOwner(run,first)))continue;
    const zipPages=[];
    for(let id=0;id<zip.pageCount();id++){
      const runs=body(zip,id).filter(run=>sameOwner(run,first));
      if(runs.length)zipPages.push({page:id+1,flow:region(zip,id),runs:runs.map(ink)});
    }
    const literal=bin.getTextRange(first.secIdx,first.paraIdx,0,0xffffffff);
    crossings.push({page:page+1,before:region(bin,page-1),after:region(bin,page),
      owner:[first.secIdx,first.paraIdx],sameLogical:literal===zip.getTextRange(first.secIdx,first.paraIdx,0,0xffffffff),
      count:literal.length,binary:signature(bin,first),package:signature(zip,first),
      beforeRuns:prev.filter(run=>sameOwner(run,first)).map(ink),afterRuns:tail.map(ink),zipPages});
  }
  console.log('WINDOWS_BINARY_PAGE_TAILS '+JSON.stringify({pages:[bin.pageCount(),zip.pageCount()],crossings}));
 } finally{bin.free();zip.free();}
}
main().catch(error=>{console.error(error);process.exitCode=1;});
