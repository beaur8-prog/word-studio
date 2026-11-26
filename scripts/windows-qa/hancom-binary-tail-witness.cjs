'use strict';
const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url');
const {unzipSync,zipSync,strFromU8,strToU8}=require('fflate');

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

const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const styleKey=({text,charStart,x,w,h,fontSize})=>[text,charStart,x,w,h,fontSize];
const units=value=>Math.round(value*75);
function measuredBaseline(doc,page,run){
 const points=[...doc.renderPageSvg(page).matchAll(/<text\s[^>]*>/g)].map(([tag])=>{
  const val=name=>Number(new RegExp(name+'="([-+\\d.]+)"').exec(tag)?.[1]);
  return {x:val('x'),y:val('y'),size:val('font-size')};
 });
 const found=new Set(points.filter(at=>Math.abs(at.size-run.fontSize)<=.065&&at.x>=run.x-.065&&at.x<=run.x+run.w+.065&&
   at.y>run.y&&at.y<=run.y+run.h).map(at=>at.y));
 return found.size===1?[...found][0]:undefined;
}
function proof(Core, source, binary, beforeSheet, spillSheet, ref, targetSheet){
 const {secIdx:section,paraIdx:paragraph}=ref, own=run=>sameOwner(run,ref), staged=[];
 const candidates=new Map();
 for(let page=0;page<source.pageCount();page++){
  const runs=body(source,page).filter(own);
  if(runs.length)candidates.set(page,runs);
 }
 for(const [page,runs] of candidates)for(const r of runs)
  if(!staged.some(v=>v.page===page&&v.run.y===r.y))staged.push({page,run:r});
 const first=staged[0]?.run, firstPage=staged[0]?.page;
 if(!first||first.charStart!==0||staged.length<2)return {status:'first-line'};
 const before=body(binary,beforeSheet).filter(own), tail=body(binary,spillSheet);
 const matched=candidates.get(targetSheet);
 if(!matched||!same(matched.map(styleKey),[...before,...tail].map(styleKey)))return {status:'target-ink'};
 const props=JSON.parse(source.getParaPropertiesAt(section,paragraph));
 const font=JSON.parse(source.getCharPropertiesAt(section,paragraph,0)).fontSize;
 if(props.lineSpacingType!=='Percent'||signature(source,ref).controlCount||props.lineSpacing<100||
    staged.some(v=>Math.abs(v.run.h-font/75)>.065))return {status:'typography'};
 const baselines=staged.map(v=>measuredBaseline(source,v.page,v.run));
 if(!baselines.every(Number.isFinite))return {status:'unmeasured-baselines'};
 const gap=Math.round(font*(props.lineSpacing/100-1));
 const top=region(source,firstPage).top;
 const ascents=staged.map((v,i)=>units(baselines[i]-v.run.y)).sort((a,b)=>a-b);
 const ascent=ascents[Math.floor(ascents.length/2)];
 const offset=units(baselines[0]-top)-ascent;
 if(!(offset>=0&&ascent>0)||staged.some((v,i)=>i&&v.page===staged[i-1].page&&
    Math.abs(units(baselines[i]-baselines[i-1])-font-gap)>1))return {status:'pitch'};
 const width=units(region(source,firstPage).bodyRight-region(source,firstPage).bodyLeft);
 const xmlLines='<hp:linesegarray>'+staged.map(({run},i)=>`<hp:lineseg textpos="${run.charStart}" vertpos="${offset+i*(font+gap)}" vertsize="${font}" textheight="${font}" baseline="${ascent}" spacing="${gap}" horzpos="0" horzsize="${width}" flags="${i?0x160000:0x60000}"/>`).join('')+'</hp:linesegarray>';
 const archive=unzipSync(source.exportHwpx());
 const name='Contents/section'+section+'.xml',xml=strFromU8(archive[name]);
 const parts=[...xml.matchAll(/<\/?hp:p\b[^>]*>/g)];
 let depth=0,index=0,end=-1,start=-1;
 for(const m of parts){
  if(m[0].startsWith('</')){if(--depth===0){if(index++===paragraph){end=m.index;break;}}}
  else if(!m[0].endsWith('/>')&&depth++===0)start=m.index;
 }
 if(end<0||xml.slice(start,end).includes('<hp:linesegarray'))return {status:'native-paragraph-cache'};
 archive[name]=strToU8(xml.slice(0,end)+xmlLines+xml.slice(end));
 const withCache=new Core(zipSync(archive,{level:1}));
 let out;
 try {
  out=new Core(withCache.exportHwp());
  const exact=[];
  for(let i=0;i<out.pageCount();i++){
   const expected=i>=spillSheet?i+1:i;
   if(expected>=binary.pageCount())break;
   const expectedRows=i===beforeSheet?body(binary,expected).filter(r=>!own(r)).concat(matched):body(binary,expected);
   exact.push({page:i+1,text:same(body(out,i),expectedRows),shapes:same(controls(out,i),controls(binary,expected))});
  }
  return {status:'readback',pages:out.pageCount(),sourceWithCache:withCache.pageCount(),offset,ascent,exact,
   relocated:body(out,beforeSheet).filter(own).map(ink)};
 }finally{out?.free();withCache.free();}
}
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
    const identity=[...prev.filter(run=>sameOwner(run,first)),...tail].map(styleKey);
    const matching=zipPages.find(({page})=>same(body(zip,page-1).filter(run=>sameOwner(run,first)).map(styleKey),identity));
    const replay=matching&&proof(HwpDocument,zip,bin,page-1,page,first,matching.page-1);
    crossings.push({replay,page:page+1,before:region(bin,page-1),after:region(bin,page),
      owner:[first.secIdx,first.paraIdx],sameLogical:literal===zip.getTextRange(first.secIdx,first.paraIdx,0,0xffffffff),
      count:literal.length,binary:signature(bin,first),package:signature(zip,first),
      beforeRuns:prev.filter(run=>sameOwner(run,first)).map(ink),afterRuns:tail.map(ink),zipPages});
  }
  console.log('WINDOWS_BINARY_PAGE_TAILS '+JSON.stringify({pages:[bin.pageCount(),zip.pageCount()],crossings}));
 } finally{bin.free();zip.free();}
}
main().catch(error=>{console.error(error);process.exitCode=1;});
