'use strict';

const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url');
const {unzipSync,zipSync,strFromU8,strToU8}=require('fflate');
const shape=(doc,page)=>JSON.parse(doc.getPageControlLayout(page)).controls.filter(c=>!c.headerFooter);
const box=c=>[c.type,c.rowCount??null,c.colCount??null,c.x,c.y,c.w,c.h];
const size=doc=>Array.from({length:doc.pageCount()},(_,i)=>shape(doc,i).map(box));
const text=doc=>Array.from({length:doc.pageCount()},(_,i)=>JSON.parse(doc.getPageTextLayout(i)).runs)
  .flat().filter(r=>!r.headerFooter&&r.paraIdx!==0xffffffff&&r.text.trim())
  .reduce((table,r)=>{
    const key=JSON.stringify([r.secIdx,r.parentParaIdx,r.paraIdx,r.cellPath]);
    table.set(key,(table.get(key)??'')+r.text);
    return table;
  },new Map());

function elements(xml,name){
  let first, level=0;
  const result=[];
  for(const m of xml.matchAll(new RegExp('<\\/?'+name+'\\b[^>]*>','g'))){
    if(m[0].startsWith('</')) {
      if(--level===0)result.push({start:first.index,open:first[0],inner:xml.slice(first.index+first[0].length,m.index)});
    } else if (!m[0].endsWith('/>')&&level++===0){ first=m; }
  }
  return result;
}

async function main(){
  const dir=path.dirname(require.resolve('@rhwp/core'));
  const {default:init,HwpDocument}=await import(pathToFileURL(path.join(dir,'rhwp.js')));
  await init({module_or_path:await fs.readFile(path.join(dir,'rhwp_bg.wasm'))});
  const [original,saved]=await Promise.all(process.argv.slice(2).map(async file=>new HwpDocument(new Uint8Array(await fs.readFile(file)))));
  try {
    const input=shape(original,0).filter(c=>c.type==='table'&&!c.cellPath?.length).at(-1);
    const onPage=shape(saved,Math.min(1,saved.pageCount()-1)), target=onPage.find(c=>c.type==='table'&&c.rowCount===1&&c.colCount===2&&!c.cellPath?.length);
    if(!input||!target)throw Error('Source or saved inline wrapper was unavailable');
    const p=(doc,c)=>JSON.parse(doc.getTableProperties(c.secIdx,c.paraIdx,c.controlIdx));
    const originalStyle=p(original,input), savedStyle=p(saved,target);
    const props=JSON.parse(saved.getPageInfo(0));
    const pkg=unzipSync(saved.exportHwpx()), file='Contents/section'+target.secIdx+'.xml', xml=strFromU8(pkg[file]);
    const para=elements(xml,'hp:p')[target.paraIdx], table=elements(para.inner,'hp:tbl')[0];
    const tag=/<hp:outMargin\b[^>]*\/>/.exec(table.inner);
    if(!tag)throw Error('The readback did not contain a measured native table margin');
    const at=para.start+para.open.length+table.start+table.open.length+tag.index;
    const prior=JSON.stringify([...text(saved)].sort());
    const attempts=[];
    const host=JSON.parse(saved.getParaPropertiesAt(target.secIdx,target.paraIdx));
    const pads={top:originalStyle.outerTop,bottom:originalStyle.outerBottom,left:originalStyle.outerLeft};
    for(const [mode,fields,resetHost] of [['outer-left',{left:pads.left},false],['host-indent',{},true],
      ['inline-position',{left:pads.left},true],['all-authored',pads,true]]){
      const patched=Object.entries(fields).reduce((s,[k,v])=>s.replace(new RegExp('\\b'+k+'="[^"]*"'),k+'="'+v+'"'),tag[0]);
      const out=new HwpDocument(zipSync({...pkg,[file]:strToU8(xml.slice(0,at)+patched+xml.slice(at+tag[0].length))},{level:1}));
      try{
        if(resetHost&&!JSON.parse(out.applyParaFormat(target.secIdx,target.paraIdx,JSON.stringify({marginLeft:0}))).ok)
          throw Error('Diagnostic wrapper indent change rejected');
        attempts.push({mode,pages:out.pageCount(),sameText:JSON.stringify([...text(out)].sort())===prior,shape:size(out)});
      }finally{out.free();}
    }
    const flow = (doc, owner) => {
      const paper = JSON.parse(doc.getPageInfo(0));
      const rows = Array.from({length:doc.getParagraphCount(owner.secIdx)}, (_, i) => {
        const {pageBreakBefore, lineSpacingType, lineSpacing, spacingBefore, spacingAfter} = JSON.parse(doc.getParaPropertiesAt(owner.secIdx,i));
        const {stored, controlCount, lines} = JSON.parse(doc.getParagraphLineMetrics(owner.secIdx,i,'[]'));
        return {i, length:doc.getParagraphLength(owner.secIdx,i), stored, controlCount, pageBreakBefore, lineSpacingType, lineSpacing,
          spacingBefore, spacingAfter, font:JSON.parse(doc.getCharPropertiesAt(owner.secIdx,i,0)).fontSize,
          lines:lines.map(line=>[line.verticalPosition,line.height,line.spacing])};
      });
      const pages=Array.from({length:doc.pageCount()},(_,i)=>({
        objects:shape(doc,i).map(box),
        lines:JSON.parse(doc.getPageTextLayout(i)).runs.filter(run=>!run.headerFooter && run.paraIdx!==0xffffffff &&
          !run.cellPath?.length && run.text.trim()).map(run=>[run.paraIdx??null,run.y,run.h,run.text.replace(/\s/gu,'').length])
      }));
      return {top:paper.marginTop+paper.marginHeader,bottom:paper.footerArea.y,rows,pages};
    };
    console.log('WINDOWS_RETURN_FORM_FLOW '+JSON.stringify({original:flow(original,input),target:flow(saved,target)}));
    console.log('WINDOWS_RETURN_FORM_INSET_TRIAL '+JSON.stringify({original:box(input),target:box(target),bounds:[props.width,props.height,props.footerArea.y],
      originalInsets:[originalStyle.outerTop,originalStyle.outerBottom,originalStyle.outerLeft],currentInsets:[savedStyle.outerTop,savedStyle.outerBottom,savedStyle.outerLeft],hostLeft:host.marginLeft,control:size(saved),attempts}));
  } finally {original.free();saved.free();}
}
main().catch(error=>{console.error(error);process.exitCode=1;});
