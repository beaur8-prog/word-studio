'use strict';
// Compare fixed source paragraphs with the real print layout. Matching uses
// complete normalized paragraph text; ambiguous/repeated text stays explicit.
const fs=require('node:fs/promises'),path=require('node:path'),z=require('node:zlib');
const {chromium}=require('@playwright/test');
const normalize=s=>(s||'').normalize('NFKC').replace(/[\s\u0000-\u001f\ufffc]/gu,'');
// Compare painted line boundaries, never insert cached source wrapping into
// the document. Mixed-size glyphs sharing a baseline must remain one line.
function sourceVisualLines(runs) {
 const lines=[];
 for(const run of runs){
  if(!normalize(run.text))continue;
  const center=run.y+run.height/2;
  let line=lines.find(l=>l.page===run.page&&Math.abs(l.center-center)<Math.max(2,Math.min(l.height,run.height)*.45));
  if(!line){line={page:run.page,y:run.y,center,height:run.height,text:''};lines.push(line)}
  line.text+=run.text;
 }
 return lines;
}
function compareLines(paragraph) {
 const source=sourceVisualLines(paragraph.source.lines),actual=paragraph.actual.flatMap(p=>p.visualLines||[]);
 // Some controls paint text that is absent from the logical paragraph API.
 // Such paragraphs cannot be used as evidence of a line-break mismatch.
 const text=normalize(paragraph.text);
 if(normalize(source.map(l=>l.text).join(''))!==text||normalize(actual.map(l=>l.text).join(''))!==text)
  return{comparable:false,reason:'paintedTextDoesNotMatchLogicalParagraph'};
 const ends=lines=>{let at=0;return lines.map(l=>at+=normalize(l.text).length)};
 const a=ends(source),b=ends(actual),index=a.findIndex((end,i)=>end!==b[i]);
 const exact=a.length===b.length&&index<0;
 const first=index<0&&!exact?Math.min(a.length,b.length):index;
 return{comparable:true,sourceLines:a.length,actualLines:b.length,exactLineBoundaries:exact,sourceEnds:a,actualEnds:b,
  ...(exact?{placement:paragraph.pageDifference===0?'samePage':'sameWrappingDifferentPage'}:{firstDifference:{line:first+1,sourceEnd:a[first]??null,actualEnd:b[first]??null,
   sourceLine:source[first]?.text,actualLine:actual[first]?.text}})};
}
(async()=>{
 const directory=process.argv[2]||'analysis/hwp-12h-2026-10-28',modelDirectory=process.argv[3]||path.join(directory,'converted'),label=process.argv[4]||'baseline';
 const auditLines=process.argv.includes('--lines');
 if(auditLines&&label==='baseline')throw Error('Use a new report label with --lines; frozen baseline anchors must not be regenerated.');
 const ids=['02-work-rules','04-language-plan','06-mixed-document','07-english-exam'];
 const browser=await chromium.launch({channel:process.env.PLAYWRIGHT_CHANNEL||'chrome',headless:true});
 const out=[];
 try{
  const page=await browser.newPage({viewport:{width:1440,height:1000}});page.setDefaultTimeout(180000);await page.goto(process.env.HWP_AUDIT_URL||'http://127.0.0.1:5196');
  for(const id of ids){
   const source=JSON.parse(await fs.readFile(path.join(directory,'source',id+'.json'),'utf8'));
   let reference;try{reference=JSON.parse(await fs.readFile(path.join(directory,'pdf-reference',id+'.json'),'utf8'));}catch{}
   const model=JSON.parse(z.gunzipSync(await fs.readFile(path.join(modelDirectory,id+'.wordstudio.gz'))));
   const result=await page.evaluate(async({model,source,auditLines})=>{
    const norm=s=>(s||'').normalize('NFKC').replace(/[\s\u0000-\u001f\ufffc]/gu,'');
    const{generateHTML}=await import('/src/audit-exports.ts'),{editorExtensions}=await import('/src/extensions.ts'),{createPrintLayout}=await import('/src/print-layout.ts');
    const root=document.createElement('div');root.innerHTML=generateHTML(model.content,editorExtensions);
    const map=new Map();for(const g of source.paragraphs){const key=norm(g.text),values=map.get(key)||[];values.push(g);map.set(key,values)}
    const used=new Set(),matches=[];let unmatched=0,ambiguous=0;
    for(const p of root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')){
     if(p.closest('[data-document-note]'))continue;const text=norm(p.textContent);if(!text)continue;
     const candidates=map.get(text)||[];if(candidates.length!==1){if(candidates.length)ambiguous++;else unmatched++;continue}
     const g=candidates[0];if(used.has(g.key)){ambiguous++;continue}used.add(g.key);p.dataset.auditKey=g.key;matches.push({key:g.key,text:p.textContent,source:g});
    }
    const tables=[...root.querySelectorAll('table')].filter(t=>!t.parentElement.closest('table,[data-document-note]'));
    const sourceTables=[...source.tables].sort((a,b)=>{const aa=a.key.split(':').map(Number),bb=b.key.split(':').map(Number);return aa[0]-bb[0]||aa[1]-bb[1]||aa[2]-bb[2]});
    const tableCountMatches=tables.length===sourceTables.length;
    if(tableCountMatches)tables.forEach((t,i)=>t.dataset.auditTable=sourceTables[i].key);
    const layout=await createPrintLayout(root.innerHTML,model.page);
    try{
     const actual=new Map();
     for(const p of layout.element.querySelectorAll('[data-audit-key]')){
      const paper=p.closest('.print-page'),box=p.getBoundingClientRect(),origin=paper.getBoundingClientRect(),values=actual.get(p.dataset.auditKey)||[];
      const rects=[],visualLines=[],walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);
      while(walker.nextNode()){
       if(!walker.currentNode.textContent.trim())continue;
       const node=walker.currentNode,range=document.createRange();range.selectNodeContents(node);
       for(const r of range.getClientRects())if(r.width>.2&&r.height>1)rects.push({x:r.x-origin.x,y:r.y-origin.y,width:r.width,height:r.height});
       if(auditLines){
        let offset=0;
        for(const character of node.textContent){
         const start=offset;offset+=character.length;if(!norm(character))continue;
         range.setStart(node,start);range.setEnd(node,offset);const r=range.getBoundingClientRect();if(r.height<1)continue;
         const center=(r.top+r.bottom)/2-origin.y;
         let line=visualLines.find(l=>Math.abs(l.center-center)<Math.max(2,Math.min(l.height,r.height)*.45));
         if(!line){line={page:Number(paper.dataset.page),y:r.y-origin.y,center,height:r.height,text:''};visualLines.push(line)}
         line.text+=character;
        }
       }
      }
      values.push({page:Number(paper.dataset.page),x:box.x-origin.x,y:box.y-origin.y,width:box.width,height:box.height,inkTop:rects.length?Math.min(...rects.map(r=>r.y)):null,inkBottom:rects.length?Math.max(...rects.map(r=>r.y+r.height)):null,lineHeight:getComputedStyle(p).lineHeight,css:p.style.cssText,text:p.textContent,rects,...(auditLines?{visualLines}:{})});actual.set(p.dataset.auditKey,values);
     }
     const paragraphs=matches.map(m=>({...m,actual:actual.get(m.key)||[]}));
     const actualTables=[...layout.element.querySelectorAll('[data-audit-table]')].map(t=>{const paper=t.closest('.print-page'),r=t.getBoundingClientRect(),p=paper.getBoundingClientRect();return{key:t.dataset.auditTable,page:Number(paper.dataset.page),x:r.x-p.x,y:r.y-p.y,width:r.width,height:r.height}});
     return{printPages:layout.pages,matched:matches.length,unmatched,ambiguous,tableCountMatches,paragraphs,tables:actualTables};
    }finally{layout.dispose()}
   },{model,source,auditLines});
   const paragraphs=result.paragraphs.map(p=>{
    const sourceStart=p.source.lines[0],actualStart=p.actual[0];const text=normalize(p.text),pdfPages=reference?.pages.filter(page=>normalize(page.text).includes(text)).map(p=>p.page)||[];
    return{...p,sourcePage:sourceStart?.page,actualPage:actualStart?.page,pageDifference:actualStart&&sourceStart?actualStart.page-sourceStart.page:null,yDifference:actualStart&&sourceStart?actualStart.y-sourceStart.y:null,pdfPages};
   });
   if(auditLines)for(const p of paragraphs)if(normalize(p.text).length>=20)p.lineComparison=compareLines(p);
   let anchors;
   const anchorFile=path.join(directory,'anchors-'+id+'.json');
   if(label==='baseline'){
    // Prefer unique substantive paragraphs, spread over the entire document.
    const pool=paragraphs.filter(p=>normalize(p.text).length>=20&&p.source.lines.length&&p.actual.length&&(!reference||p.pdfPages.length===1));
    const count=Math.min(40,pool.length),selected=[];
    for(let i=0;i<count;i++){const p=pool[Math.floor(i*(pool.length-1)/Math.max(1,count-1))];selected.push({key:p.key,text:p.text,sourcePage:p.sourcePage,pdfPages:p.pdfPages,baselinePage:p.actualPage});}
    anchors=selected;await fs.writeFile(anchorFile,JSON.stringify(anchors,null,2)+'\n');
   }else anchors=JSON.parse(await fs.readFile(anchorFile,'utf8'));
   const checks=anchors.map(a=>{const p=paragraphs.find(p=>p.key===a.key);return{...a,currentPage:p?.actualPage,pageDifference:p?.pageDifference,yDifference:p?.yDifference,missing:!p}});
   const summary={id,sourcePages:source.sourcePages,printPages:result.printPages,matched:result.matched,unmatched:result.unmatched,ambiguous:result.ambiguous,anchors:checks.length,missingAnchors:checks.filter(c=>c.missing).length,mismatchedAnchors:checks.filter(c=>c.pageDifference!==0).length,meanAbsolutePageDifference:checks.reduce((n,c)=>n+Math.abs(c.pageDifference||0),0)/Math.max(1,checks.length),firstMismatch:paragraphs.find(p=>p.pageDifference!==0&&normalize(p.text).length>=20)?.key};
   if(auditLines){
    const eligible=paragraphs.filter(p=>p.lineComparison?.comparable),first=eligible.find(p=>!p.lineComparison.exactLineBoundaries);
    summary.lineAudit={scope:'Same unique logical paragraphs, 20+ normalized characters. Native-source coordinates are not independent Microsoft Word/Hancom rendering.',
     comparable:eligible.length,excluded:paragraphs.filter(p=>p.lineComparison&&!p.lineComparison.comparable).length,
     exact:eligible.filter(p=>p.lineComparison.exactLineBoundaries).length,
     changedLineCount:eligible.filter(p=>p.lineComparison.sourceLines!==p.lineComparison.actualLines).length,
     sameWrappingDifferentPage:eligible.filter(p=>p.lineComparison.placement==='sameWrappingDifferentPage').length,
     firstMismatch:first?{key:first.key,...first.lineComparison.firstDifference}:null};
   }
   out.push({...summary,checks});await fs.writeFile(path.join(directory,`${label}-correspondence-${id}.json`),JSON.stringify({...summary,paragraphs,tables:result.tables,tableCountMatches:result.tableCountMatches},null,2));console.log(JSON.stringify(summary));
  }
  await fs.writeFile(path.join(directory,label+'-page-correspondence.json'),JSON.stringify(out,null,2)+'\n');
 }finally{await browser.close()}
})().catch(e=>{console.error(e);process.exitCode=1});
