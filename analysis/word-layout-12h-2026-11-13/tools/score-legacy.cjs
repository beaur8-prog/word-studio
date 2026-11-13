const fs=require('node:fs/promises'),path=require('node:path'),z=require('node:zlib'),{createRequire}=require('node:module'),{pathToFileURL}=require('node:url'),{createHash}=require('node:crypto');
const root=path.resolve(__dirname,'../../..'),B=path.dirname(__dirname),req=createRequire(root+'/package.json'),fixed=root+'/analysis/hwp-libreoffice-8h-2026-10-30',label=process.env.GOV_LEGACY_LABEL||'legacy-fresh',fixedIds=['02-work-rules','04-language-plan','06-mixed-document','07-english-exam'];
const hash=bytes=>createHash('sha256').update(bytes).digest('hex');
(async()=>{
 const {createServer}=await import(pathToFileURL(req.resolve('vite'))),{chromium}=req('@playwright/test'),browser=await chromium.launch({channel:'chrome',headless:true});
 const reports=[],proofs={};let server,cache;
 try{
  for(const version of (process.env.GOV_LEGACY_VERSIONS||'baseline,current').split(',')){
   const source=root,input='fresh-legacy-reference';
  const isolation={name:'isolated-current-source',enforce:'pre',load(id){
 const rel=path.relative(root,id.split('?')[0]);const local=JSON.parse(process.env.LEGACY_LOCAL_SOURCE_OVERRIDES||'{}');if(local[rel])return require('node:fs').readFileSync(local[rel],'utf8');const changed=['src/App.tsx','src/docx-stories.ts','src/editor-cell-pagination.ts','src/editor-paragraph-pagination.ts','src/pagination.ts','src/paragraph-fragments.ts','src/section-node-view.ts','src/styles.css','src/print-layout.ts'];
 const partial=JSON.parse(process.env.LEGACY_FILE_REFS||'{}');if(partial[rel])return require('node:child_process').execFileSync('git',['show',partial[rel]+':'+rel],{cwd:root,encoding:'utf8'});
 if(process.env.LEGACY_REF&&rel.startsWith('src/')&&/\.(ts|tsx|css|json)$/.test(rel))return require('node:child_process').execFileSync('git',['show',process.env.LEGACY_REF+':'+rel],{cwd:root,encoding:'utf8'});
 if(rel==='src/styles.css'){
  if(process.env.LEGACY_CSS_REF)return require('node:child_process').execFileSync('git',['show',process.env.LEGACY_CSS_REF+':'+rel],{cwd:root,encoding:'utf8'});
  if(process.env.LEGACY_DEFAULT_OLD==='1')return require('node:fs').readFileSync(id,'utf8').replace('font-family: "Malgun Gothic", "맑은 고딕", Arial, "AppleGothic", sans-serif;','font-family: "Malgun Gothic", "맑은 고딕", Arial, sans-serif;');
 }
}};
 cache=await fs.mkdtemp('/private/tmp/ws-legacy-score-');server=await createServer({root:source,cacheDir:cache,plugins:version==='baseline'?[{name:'frozen-product',enforce:'pre',load(id){const rel=path.relative(root,id.split('?')[0]);if(['src/document-io.ts','src/print-layout.ts','src/paragraph-fit.ts','src/pagination.ts','src/types.ts','src/section-node-view.ts','src/App.tsx','src/android/EditorCanvas.tsx','src/row-split-policy.ts','src/table-fragments.ts','src/editor-cell-pagination.ts'].includes(rel))return require('node:child_process').execFileSync('git',['show','5cb4246:'+rel],{cwd:root,encoding:'utf8'});}}]:[isolation],configFile:source+'/vite.config.ts',server:{host:'127.0.0.1',port:Number(process.env.LEGACY_PORT||5448),strictPort:true,hmr:false,watch:{ignored:['**']},fs:{allow:[root,source]}},logLevel:'error'});await server.listen();
   const retained=root+'/analysis/word-layout-4h-2026-11-10';const rows=JSON.parse(await fs.readFile(retained+'/'+input+'/results.json')).rows;
   for(const row of rows.filter(r=>!process.env.LEGACY_IDS||process.env.LEGACY_IDS.split(',').includes(r.id))){
    if(row.status!=='converted')throw Error('Incomplete original conversion: '+row.id);
    const started=Date.now(),page=await browser.newPage({viewport:{width:1440,height:1000},deviceScaleFactor:1});page.setDefaultTimeout(300000);
    try{
     await page.goto('http://127.0.0.1:'+(process.env.LEGACY_PORT||5448)+'/');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();if(process.env.LEGACY_EXPERIMENT_CSS)await page.addStyleTag({path:process.env.LEGACY_EXPERIMENT_CSS});
     const rt=JSON.parse(await fs.readFile(retained+'/roundtrip-legacy-candidate/results.json')),rr=rt.rows.find(r=>r.id===row.id),saved=rr?.files?.find(f=>f.save===2);if(!saved)throw Error('Saved DOCX is unavailable: '+row.id);const bytes=await fs.readFile(retained+'/'+saved.file);if(hash(bytes)!==saved.sha256)throw Error('Saved DOCX bytes changed');const sourceModel=[...bytes];
     let cohort=[],anchors=[];
     if(fixedIds.includes(row.id)){
      const file=fixed+'/table-indent-docx2-correspondence-'+row.id+'.json.gz',bytes=await fs.readFile(file);cohort=JSON.parse(z.gunzipSync(bytes)).paragraphs.map(p=>({key:p.key,text:p.text,sourcePage:p.sourcePage}));
      const af=fixed+'/anchors-'+row.id+'.json',ab=await fs.readFile(af);anchors=JSON.parse(ab).map(p=>({key:p.key,text:p.text,sourcePage:p.sourcePage}));
      proofs[row.id]={inputSHA256:row.sha256,cohortSHA256:hash(bytes),cohortCount:cohort.length,anchorsSHA256:hash(ab),anchorCount:anchors.length};
     }
     console.log(JSON.stringify({version,id:row.id,phase:'start',fixed:cohort.length}));
     const result=await page.evaluate(async({sourceModel,cohort,anchors,id})=>{
      const{prepareEditableHwp}=await import('/src/hwp/svg-images.ts'),{importEditableHwp}=await import('/src/hwp/editing.ts'),{generateHTML}=await import('/node_modules/@tiptap/core/dist/index.js'),{editorExtensions}=await import('/src/extensions.ts'),{createPrintLayout}=await import('/src/print-layout.ts');
      const{importFile}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');const loaded=await importFile(new File([new Uint8Array(sourceModel)],id+'.docx'));const doc={...loaded.document,content:parseDocumentHtml(loaded.html)};await document.fonts.ready;
      const counts={},images=[],cells=[],text=[];
      const logical=n=>n.type==='text'?n.text:n.type==='hardBreak'?'\n':n.type==='documentTab'?'\t':(n.content||[]).map(logical).join('');
      async function sha(bytes){return[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(n=>n.toString(16).padStart(2,'0')).join('');}
      function walk(n){counts[n.type]=(counts[n.type]||0)+1;if(n.type==='text')text.push(n.text);if(n.type==='hardBreak')text.push('\n');if(n.type==='documentTab')text.push('\t');if(['image','inlineImage'].includes(n.type))images.push(n.attrs.src);if(['tableCell','tableHeader'].includes(n.type))cells.push({text:logical(n),rowspan:n.attrs?.rowspan||1,colspan:n.attrs?.colspan||1,paragraphs:(n.content||[]).filter(c=>c.type==='paragraph').length});(n.content||[]).forEach(walk);}walk(doc.content);
      const encoder=new TextEncoder(),plain=text.join(''),imageHashes=[];
      for(const src of images){const m=src.match(/^data:([^;,]+);base64,([\s\S]*)$/);if(!m)throw Error('Unexpected image source');imageHashes.push({mime:m[1],sha256:await sha(Uint8Array.from(atob(m[2]),c=>c.charCodeAt(0)))});}
      const report={logicalCharacters:plain.length,logicalSHA256:await sha(encoder.encode(plain)),semanticSpaceSHA256:await sha(encoder.encode(plain.replace(/\u2005/g,'\u2007'))),spaceCensus:{figure:(plain.match(/\u2007/g)||[]).length,quarter:(plain.match(/\u2005/g)||[]).length,breaks:(plain.match(/\n/g)||[]).length,tabs:(plain.match(/\t/g)||[]).length},cellCount:cells.length,cellsSHA256:await sha(encoder.encode(JSON.stringify(cells))),semanticCellsSHA256:await sha(encoder.encode(JSON.stringify(cells).replace(/\u2005/g,'\u2007'))),images:imageHashes,counts};

      const norm=s=>(s||'').normalize('NFKC').replace(/[\s\u0000-\u001f\ufffc]/gu,''),root=document.createElement('div');root.innerHTML=generateHTML(doc.content,editorExtensions);
      const paragraphs=new Map();for(const p of root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')){if(p.closest('[data-document-note]'))continue;const key=norm(p.textContent);if(!key)continue;const a=paragraphs.get(key)||[];a.push(p);paragraphs.set(key,a);}
      const all=new Map([...cohort,...anchors].map(p=>[p.key,p])),missing=[],ambiguous=[];
      for(const p of all.values()){const nodes=paragraphs.get(norm(p.text))||[];if(nodes.length===1)nodes[0].dataset.frozenLegacyKey=p.key;else(nodes.length?ambiguous:missing).push(p.key);}
      const layout=await createPrintLayout(root.innerHTML,doc.page);try{
       const actual=new Map();for(const[pageIndex,paper]of[...layout.element.querySelectorAll('.print-page')].entries()){const base=paper.getBoundingClientRect();for(const p of paper.querySelectorAll('[data-frozen-legacy-key]')){const r=p.getBoundingClientRect(),at=actual.get(p.dataset.frozenLegacyKey)||[];at.push({page:pageIndex+1,x:r.x-base.x,y:r.y-base.y,width:r.width,height:r.height});actual.set(p.dataset.frozenLegacyKey,at);}}
       const score=set=>{const rows=set.map(p=>{const a=actual.get(p.key)||[],first=a[0];return{key:p.key,sourcePage:p.sourcePage,actual:first||null,difference:first?first.page-p.sourcePage:null,fragments:a.length,missing:missing.includes(p.key),ambiguous:ambiguous.includes(p.key)};});return{total:set.length,found:rows.filter(p=>p.actual).length,missing:rows.filter(p=>!p.actual).length,mismatched:rows.filter(p=>p.actual&&p.difference!==0).length,absolutePageError:rows.reduce((s,p)=>s+(p.difference===null?0:Math.abs(p.difference)),0),rows};};
       return{...report,printPages:layout.pages,warnings:layout.warnings,fixed3007:score(cohort),fixed160:score(anchors)};
      }finally{layout.dispose();}
     },{sourceModel,cohort,anchors,id:row.id});
     const report={version,id:row.id,inputSHA256:row.sha256,savedDocxSHA256:saved.sha256,sourceRoot:source,durationMs:Date.now()-started,...result};reports.push(report);
     await fs.writeFile(B+'/'+label+'.json.gz',z.gzipSync(JSON.stringify({method:'same-second-saved-DOCX/prepared-HWP-models-not-fresh-parsing/frozen-native-reference-keys-v1',experimentalCSS:process.env.LEGACY_EXPERIMENT_CSS?path.basename(process.env.LEGACY_EXPERIMENT_CSS):null,fontEnvironment:'actual app root startup, same native aliases and bundled fonts for both stages; not Microsoft Word',sourceCommit:process.env.LEGACY_REF||null,sourceCode:Object.fromEntries(['src/styles.css','src/word-font-fallbacks.ts','src/print-layout.ts','src/print-fonts.ts'].map(f=>[f,hash(require('node:fs').readFileSync(root+'/'+f))])),proofs,rows:reports})));
     console.log(JSON.stringify({...report,images:result.images.length,fixed3007:result.fixed3007&&{...result.fixed3007,rows:undefined},fixed160:result.fixed160&&{...result.fixed160,rows:undefined}}));
    }finally{await page.close();}
   }
   await server.close();server=undefined;await fs.rm(cache,{recursive:true,force:true});cache=undefined;
  }
 }finally{await browser.close();await server?.close();if(cache)await fs.rm(cache,{recursive:true,force:true});}
})().catch(e=>{console.error(e);process.exitCode=1});
