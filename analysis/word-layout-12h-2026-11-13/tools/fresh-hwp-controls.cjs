const fs=require('node:fs/promises'),path=require('node:path'),z=require('node:zlib'),{execFileSync}=require('node:child_process'),{createRequire}=require('node:module'),{pathToFileURL}=require('node:url'),crypto=require('node:crypto');
const root=path.resolve(__dirname,'../../..'),base=path.dirname(__dirname),req=createRequire(root+'/package.json'),hash=bytes=>crypto.createHash('sha256').update(bytes).digest('hex');
(async()=>{
 const out=base+'/'+(process.env.HWP_FRESH_STAGE||'fresh-hwp-controls');await fs.mkdir(out,{recursive:true});
 const inputs=JSON.parse(await fs.readFile(root+'/analysis/word-layout-4h-2026-11-10/fresh-legacy-reference/results.json')).rows;
 const rows=[],errors=[];
 const refs=JSON.parse(process.env.HWP_FRESH_REFS||'["eee8aa0","a4685e2"]');
 for(const ref of refs){
  const {createServer}=await import(pathToFileURL(req.resolve('vite'))),sources=new Map(),cache=await fs.mkdtemp('/private/tmp/ws-layout-fresh-hwp-');
  const frozen={name:'frozen-source',enforce:'pre',load(id){const file=path.relative(root,id.split('?')[0]);if(file.startsWith('src/')&&/\.(ts|tsx|css|json)$/.test(file)){if(!sources.has(file))sources.set(file,execFileSync('git',['show',ref+':'+file],{cwd:root,encoding:'utf8'}));return sources.get(file);}}};
  const server=await createServer({root,cacheDir:cache,configFile:root+'/vite.config.ts',plugins:[frozen],server:{host:'127.0.0.1',port:5643,strictPort:true,hmr:false,watch:{ignored:['**']}},logLevel:'error'});await server.listen();
  const browser=await req('@playwright/test').chromium.launch({channel:'chrome',headless:true});
  try{for(const input of inputs){
   const bytes=await fs.readFile(input.file);if(hash(bytes)!==input.sha256)throw Error('Original changed '+input.id);
   const started=Date.now(),page=await browser.newPage({viewport:{width:1440,height:1000},deviceScaleFactor:1});page.setDefaultTimeout(360000);page.on('pageerror',error=>errors.push({ref,id:input.id,error:String(error)}));
   try{
    await page.goto('http://127.0.0.1:5643');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
    console.log(JSON.stringify({ref,id:input.id,phase:'fresh-original-start'}));
    const result=await page.evaluate(async({bytes,id,current})=>{
     const {HwpReaderSession}=await import('/src/hwp/reader.ts'),{importEditableHwp}=await import('/src/hwp/editing.ts'),{exportDocx,importFile}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/@tiptap/core/dist/index.js'),{createPrintLayout}=await import('/src/print-layout.ts');
     const session=new HwpReaderSession();let doc,info;try{info=await session.open(new File([new Uint8Array(bytes)],id+'.hwp'));doc=importEditableHwp(await session.editable(),id+'.hwp');}finally{session.close();}
     const sha=async bytes=>[...new Uint8Array(await crypto.subtle.digest('SHA-256',typeof bytes==='string'?new TextEncoder().encode(bytes):bytes))].map(v=>v.toString(16).padStart(2,'0')).join('');
     async function summary(doc){const pictures=[],cells=[],counts={},floats=[];function visit(n){counts[n.type]=(counts[n.type]||0)+1;const text=n.type==='hardBreak'?'\n':n.type==='documentTab'?'\t':n.type==='text'?n.text:(n.content||[]).map(visit).join('');if(['tableCell','tableHeader'].includes(n.type))cells.push([n.attrs?.rowspan||1,n.attrs?.colspan||1,text]);if(['image','inlineImage'].includes(n.type)){pictures.push(n.attrs.src);if(n.attrs.floating)floats.push(JSON.parse(n.attrs.floating));}return text;}const text=visit(doc.content),images=[];for(const src of pictures){const m=src.match(/^data:([^;,]+);base64,(.*)$/s);if(!m)throw Error('Non-raster image');images.push([m[1],await sha(Uint8Array.from(atob(m[2]),c=>c.charCodeAt(0)))]);}return{characters:text.length,bodySHA256:await sha(text),counts,cellsSHA256:await sha(JSON.stringify(cells)),images,floats};}
     const states=[await summary(doc)],saves=[];
     const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);let printed;try{printed={pages:layout.pages,warnings:layout.warnings,pictures:[...layout.element.querySelectorAll('.print-content img[data-image-floating]')].map(image=>{const r=image.getBoundingClientRect(),p=image.closest('.print-page'),o=p.getBoundingClientRect();return{page:Number(p.dataset.page),x:r.x-o.x,y:r.y-o.y,width:r.width,height:r.height,display:getComputedStyle(image).display};})};}finally{layout.dispose();}
     if(current)for(let pass=1;pass<=2;pass++){const blob=await exportDocx(doc),bytes=new Uint8Array(await blob.arrayBuffer());let binary='';for(let i=0;i<bytes.length;i+=8192)binary+=String.fromCharCode(...bytes.subarray(i,i+8192));saves.push(btoa(binary));const loaded=await importFile(new File([blob],id+'.docx'));doc={...loaded.document,content:parseDocumentHtml(loaded.html)};states.push(await summary(doc));}
     return{nativePages:info.count,states,printed,saves};
    },{bytes:[...bytes],id:input.id,current:ref===refs.at(-1)});
    const saves=[];for(const[index,data]of result.saves.entries()){const bytes=Buffer.from(data,'base64'),file=out+'/textbox-save-fresh-'+input.id+'-'+(index+1)+'.docx';await fs.writeFile(file,bytes);saves.push({file:path.relative(base,file),sha256:hash(bytes),bytes:bytes.length});}
    delete result.saves;const row={ref,id:input.id,inputSHA256:input.sha256,...result,saves,durationMs:Date.now()-started};rows.push(row);console.log(JSON.stringify({ref,id:input.id,pages:row.printed.pages,characters:row.states[0].characters,images:row.states[0].images.length,finished:true,durationMs:row.durationMs}));
   }catch(error){errors.push({ref,id:input.id,error:String(error)});console.error(String(error));}finally{await page.close();}
   await fs.writeFile(out+'/results.json',JSON.stringify({method:'Fresh original HWP bytes through production Worker and editable import, not retained prepared models. Source refs recorded per row; last source model additionally exported and reopened as real DOCX twice. Native HWP parser page count is not Microsoft Word.',rows,errors},null,2));
  }}finally{await browser.close();await server.close();await fs.rm(cache,{recursive:true,force:true});}
 }
 if(errors.length)process.exitCode=1;
})().catch(error=>{console.error(error);process.exitCode=1});
