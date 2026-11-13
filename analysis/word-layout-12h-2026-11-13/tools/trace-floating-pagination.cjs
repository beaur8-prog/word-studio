const fs = require('node:fs/promises');
const path = require('node:path');
const z = require('node:zlib');
const { execFileSync } = require('node:child_process');
const { createRequire } = require('node:module');
const { pathToFileURL } = require('node:url');
const root = path.resolve(__dirname, '../../..'), out = path.dirname(__dirname), req = createRequire(root + '/package.json');
const mode = process.env.OSC_MODE || 'frozen', port = Number(process.env.OSC_PORT || 5191), ref = process.env.OSC_REF || '6ac3f2b';
const frozen = file => ref === 'current' ? require('node:fs').readFileSync(path.join(root,file),'utf8') : execFileSync('git', ['show',ref + ':' + file], {cwd:root,encoding:'utf8'});
(async () => {
  const candidateCode=process.env.OSC_FLOATING_FILE?await fs.readFile(process.env.OSC_FLOATING_FILE,'utf8'):null;
  if(candidateCode)await fs.writeFile(out+'/candidate-'+mode+'.ts.gz',z.gzipSync(candidateCode));
  const paginationCode=process.env.OSC_PAGINATION_FILE?await fs.readFile(process.env.OSC_PAGINATION_FILE,'utf8'):null;
  if(paginationCode)await fs.writeFile(out+'/pagination-candidate-'+mode+'.ts.gz',z.gzipSync(paginationCode));
  const candidateSHA256=candidateCode?require('node:crypto').createHash('sha256').update(candidateCode).digest('hex'):null;
  const plugin = { name:'read-only-pagination-trace', enforce:'pre', load(id) {
    const file = path.relative(root,id.split('?')[0]);
    if (['src/App.tsx','src/editor-paragraph-pagination.ts','src/paragraph-fragments.ts','src/docx-stories.ts','src/section-node-view.ts'].includes(file)) return frozen(file);
    if(file==='src/floating-layout.ts')return (candidateCode||frozen(file)).replace('const relevant = () =>', "if(p.textContent?.includes('\uc694\uccad\uad00\ub9ac \uc808\ucc28\ub294')){const h=((window as any).__guideFloats||=[]);h.push({box:{top:box.top,bottom:box.bottom,left:box.left,width:box.width},paragraphText:p.textContent,flowing:flowing.filter(pic=>pic.top+pic.height>=box.top-1&&pic.top<box.bottom+1).map(pic=>({...pic,sourceCellText:pic.index>=0?images[pic.index]?.closest('td,th')?.textContent?.slice(0,120):null,sourceParagraphText:pic.index>=0?images[pic.index]?.closest('p,h1,h2,h3')?.textContent?.slice(0,160):null,sameCell:pic.index>=0&&images[pic.index]?.closest('td,th')===p.closest('td,th'),sameSection:pic.index>=0&&images[pic.index]?.closest('.document-section-content')===p.closest('.document-section-content')}))});if(h.length>18)h.shift();}\n"+'const relevant = () =>');
    if (file !== 'src/pagination.ts') {if(process.env.OSC_FULL_FREEZE==='1'&&file.startsWith('src/')&&/\.(ts|tsx|css|json)$/.test(file))return frozen(file);return;}
    let source = paginationCode||frozen(file);
    source = source.replace('const dec: Decoration[] = [], signatures: string[] = [];','const dec: Decoration[] = [], signatures: string[] = []; const tracePieces: any[] = [];');
    const needle = 'const flow = pieces(view, usable, scale, removed, context.node, context.start, context.dom);';
    if (!source.includes(needle)) throw Error('Trace insertion changed');
    source = source.replace(needle, needle + "\ntracePieces.push({start:context.start,usable,removed,flow:flow.map(p=>({pos:p.pos,end:p.end,top:p.top,height:p.height,columns:p.columns,explicit:p.explicit,cellRow:!!p.cellRow,merged:!!p.mergedRows,text:view.state.doc.nodeAt(p.pos)?.textContent.slice(0,90),detail:[35775,35839,35972].includes(p.pos)&&p.cellRow?[p.cellRow,...p.cellRow.querySelectorAll('td,p,span,img')].map(el=>{const c=getComputedStyle(el),b=el.getBoundingClientRect();return {tag:el.tagName,cls:el.className,text:el.tagName==='IMG'?'image':el.textContent?.slice(0,50),inline:el.style.cssText,box:{top:b.top,height:b.height,width:b.width},css:{height:c.height,minHeight:c.minHeight,width:c.width,maxWidth:c.maxWidth,position:c.position,verticalAlign:c.verticalAlign,paddingTop:c.paddingTop,paddingBottom:c.paddingBottom},attrs:[...el.attributes].filter(a=>a.name==='width'||a.name==='height'||a.name.startsWith('data-')).map(a=>[a.name,a.value])}}):undefined}))});");
    source = source.replace("const next = signatures.join('|');", [
      'const snapshot = {pages:totalCount,pieces:tracePieces,signatures};',
      'const history = ((window as any).__layoutTransitions ||= []);',
      'if (!history.length || history[history.length-1].pages !== totalCount) { history.push(snapshot); if(history.length>12)history.shift(); }',
      '(window as any).__lastLayoutTrace = snapshot;',
      "const next = signatures.join('|');"
    ].join('\n'));
    return source;
  } };
  const { createServer } = await import(pathToFileURL(req.resolve('vite')));
  const cache=await fs.mkdtemp('/private/tmp/ws-layout-osc-vite-');
  const server = await createServer({ root,cacheDir:cache, plugins:[plugin],configFile:root+'/vite.config.ts',
    server:{host:'127.0.0.1',port,strictPort:true,hmr:false,watch:{ignored:['**']}},logLevel:'error'});
  await server.listen();
  let browser,app,close,profile;
  try {
    let page;
    if (process.env.OSC_ELECTRON !== '0') {
      const mainPath=out+'/native-app/electron/main-oscillation.cjs';
      await fs.writeFile(mainPath,(await fs.readFile(root+'/electron/main.cjs','utf8')).replace("const DEV_URL = 'http://127.0.0.1:5191';","const DEV_URL = 'http://127.0.0.1:"+port+"';"));
      profile=await fs.mkdtemp('/tmp/ws-layout-osc-');
      const entry=profile+'/entry.cjs';
      await fs.writeFile(entry,"const{app,BrowserWindow,systemPreferences}=require('electron');const set=systemPreferences.setUserDefault.bind(systemPreferences);systemPreferences.setUserDefault=(k,...a)=>k==='NSOSPLastRootDirectory'?undefined:set(k,...a);app.whenReady().then(()=>app.setActivationPolicy('prohibited'));BrowserWindow.prototype.show=function(){};BrowserWindow.prototype.focus=function(){};app.on('browser-window-created',(_e,w)=>{w.setSkipTaskbar(true);w.webContents.setBackgroundThrottling(false)});app.whenReady().then(()=>app.dock?.hide());require("+JSON.stringify(mainPath)+");");
      const env={...process.env,WORD_STUDIO_DEV:'1',STUDIO_DISABLE_UPDATES:'1'};delete env.ELECTRON_RUN_AS_NODE;
      app=await req('@playwright/test')._electron.launch({executablePath:req('electron'),args:[entry,'--force-device-scale-factor=1','--user-data-dir='+profile+'/profile'],env,timeout:60000});
      close=await req('./scripts/native-smoke-lifetime.cjs')(app,'Pagination oscillation probe');
      browser=app; page=await app.firstWindow();
    } else {
      browser=await req('@playwright/test').chromium.launch({channel:'chrome',headless:true});
      page=await browser.newPage();await page.goto('http://127.0.0.1:'+port);
    }
    const errors=[];page.on('pageerror',e=>{errors.push(String(e));console.error('RENDERER '+String(e))});page.on('console',m=>{if(m.type()==='error'||m.text().startsWith('FORWARD'))console.log(m.text())});
    await page.setViewportSize({width:1440,height:1000});
    await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
    if(mode==='hide-row')await page.addStyleTag({content:'.pagination-measuring-cells .pagination-row-gap{display:none!important}'});
    if(mode==='hide-all')await page.addStyleTag({content:'.pagination-measuring-cells .pagination-gap{display:none!important}'});
    const previous=root+'/analysis/word-layout-4h-2026-11-10';
    const row=JSON.parse(await fs.readFile(previous+'/government-final-comparison.json')).rows.find(r=>r.id===(process.env.OSC_ID||'mss-funding-2026-0'));
    const input=path.resolve(previous,row.sourceDOCX),bytes=await fs.readFile(input);
    if(require('node:crypto').createHash('sha256').update(bytes).digest('hex')!==row.inputSHA256)throw Error('Input changed');
    if(process.env.OSC_CLIP_FLOATS) {
      const clip=Number(process.env.OSC_CLIP_FLOATS);
      const model=await page.evaluate(async({data,clip})=>{
        const {importFile}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');
        const loaded=await importFile(new File([Uint8Array.from(atob(data),c=>c.charCodeAt(0))],'SOP.docx'));
        const doc={...loaded.document,content:loaded.html===undefined?loaded.document.content:parseDocumentHtml(loaded.html)};
        let floats=0,done=false;
        const visit=n=>{if(done)return null;if(n.type==='inlineImage'&&n.attrs?.floating)floats++;if(n.content){const next=[];for(const child of n.content){if(done)break;const v=visit(child);if(v)next.push(v);if(floats>=clip && ['paragraph','heading','table','pageBreak','sectionBreak'].includes(child.type))done=true;}return {...n,content:next}}return n};
        return {...doc,content:visit(doc.content)};
      },{data:bytes.toString('base64'),clip});
      await page.locator('input[type=file]').first().setInputFiles({name:'SOP-diagnostic-prefix.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(model))});
    } else await page.locator('input[type=file]').first().setInputFiles(input);
    await page.waitForFunction(()=>((document.querySelector('.tiptap[contenteditable=true]')?.textContent||'').trim().length>100),{},{timeout:120000});
    await page.locator('.busy-indicator').waitFor({state:'hidden',timeout:120000});
    await page.evaluate(()=>document.fonts.ready);await page.waitForTimeout(2000);
    await page.evaluate(()=>{window.__layoutTransitions=[];});
    const samples=[];
    for(let i=0;i<80;i++){
      await page.waitForTimeout(500);
      const state=await page.evaluate(()=>{
        const root=document.querySelector('.tiptap[contenteditable=true]'),origin=root.getBoundingClientRect();
        const frames=[...root.querySelectorAll('img[data-image-floating]')].map(el=>{const b=el.getBoundingClientRect();return [b.x-origin.x,b.y-origin.y,b.width,b.height].map(v=>Math.round(v*64)/64)});
        const gaps=[...root.querySelectorAll('.pagination-gap,.pagination-cell-gap')].map(g=>g.style.height);
        const value=JSON.stringify([frames,gaps]);let hash=2166136261;for(let i=0;i<value.length;i++)hash=Math.imul(hash^value.charCodeAt(i),16777619);
        return {pages:document.querySelector('.document-stack')?.dataset.pages,transitions:window.__layoutTransitions?.length||0,height:origin.height,geometrySignature:(hash>>>0).toString(16),frames};
      });
      samples.push({i,...state});if(i%10===0)console.log(JSON.stringify({mode,...samples.at(-1),frames:undefined}));
      if(state.transitions>=8)break;
    }
    const trace=await page.evaluate(()=>({transitions:window.__layoutTransitions,last:window.__lastLayoutTrace,guideFloats:window.__guideFloats,forwardPlan:window.__lastForwardPlan,forwardFailure:window.__forwardOwnerFailure}));
    const report={method:'Read-only referenced source instrumentation, isolated Electron developer mode unless specified. Public original exact Word DOCX. Product mode adds no experimental CSS.',sourceCommit:execFileSync('git',['rev-parse','HEAD'],{cwd:root,encoding:'utf8'}).trim(),productFloatingSHA256:require('node:crypto').createHash('sha256').update(frozen('src/floating-layout.ts')).digest('hex'),ref,mode,diagnosticClipFloats:process.env.OSC_CLIP_FLOATS?Number(process.env.OSC_CLIP_FLOATS):null,candidateSHA256,inputSHA256:row.inputSHA256,samples,errors,...trace};
    await fs.writeFile(out+'/oscillation-'+mode+'.json.gz',z.gzipSync(JSON.stringify(report)));
    if(errors.length)process.exitCode=1;
    console.log(JSON.stringify({mode,errors,pages:[...new Set(samples.map(s=>s.pages))],transitions:trace.transitions?.map(t=>t.pages),geometrySignatures:[...new Set(samples.map(s=>s.geometrySignature))]}));
  } finally {
    if(close)await close();else if(browser)await browser.close();
    await server.close();await fs.rm(cache,{recursive:true,force:true});if(profile)await fs.rm(profile,{recursive:true,force:true});
  }
})().catch(e=>{console.error(e);process.exitCode=1});
