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
  const plugin = { name:'read-only-pagination-trace', enforce:'pre', load(id) {
    const file = path.relative(root,id.split('?')[0]);
    if (['src/App.tsx','src/editor-paragraph-pagination.ts','src/paragraph-fragments.ts','src/docx-stories.ts','src/section-node-view.ts'].includes(file)) return frozen(file);
    if (file !== 'src/pagination.ts') return;
    let source = frozen(file);
    source = source.replace('const dec: Decoration[] = [], signatures: string[] = [];','const dec: Decoration[] = [], signatures: string[] = []; const tracePieces: any[] = [];');
    const needle = 'const flow = pieces(view, usable, scale, removed, context.node, context.start, context.dom);';
    if (!source.includes(needle)) throw Error('Trace insertion changed');
    source = source.replace(needle, needle + '\ntracePieces.push({start:context.start,usable,removed,flow:flow.map(p=>({pos:p.pos,end:p.end,top:p.top,height:p.height,columns:p.columns,explicit:p.explicit,cellRow:!!p.cellRow,merged:!!p.mergedRows,text:view.state.doc.nodeAt(p.pos)?.textContent.slice(0,90)}))});');
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
  const server = await createServer({ root, plugins:[plugin],configFile:root+'/vite.config.ts',
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
    await page.setViewportSize({width:1440,height:1000});
    await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
    if(mode==='hide-row')await page.addStyleTag({content:'.pagination-measuring-cells .pagination-row-gap{display:none!important}'});
    if(mode==='hide-all')await page.addStyleTag({content:'.pagination-measuring-cells .pagination-gap{display:none!important}'});
    const previous=root+'/analysis/word-layout-4h-2026-11-10';
    const row=JSON.parse(await fs.readFile(previous+'/government-final-comparison.json')).rows.find(r=>r.id==='mss-funding-2026-0');
    const input=path.resolve(previous,row.sourceDOCX),bytes=await fs.readFile(input);
    if(require('node:crypto').createHash('sha256').update(bytes).digest('hex')!==row.inputSHA256)throw Error('Input changed');
    await page.locator('input[type=file]').first().setInputFiles(input);
    await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).filter({hasText:'정책자금'}).waitFor({timeout:120000});
    await page.locator('.busy-indicator').waitFor({state:'hidden',timeout:120000});
    await page.evaluate(()=>document.fonts.ready);await page.waitForTimeout(2000);
    await page.evaluate(()=>{window.__layoutTransitions=[];});
    const samples=[];
    for(let i=0;i<80;i++){
      await page.waitForTimeout(500);
      const state=await page.evaluate(()=>({pages:document.querySelector('.document-stack')?.dataset.pages, transitions:window.__layoutTransitions?.length||0,height:document.querySelector('.tiptap[contenteditable=true]')?.getBoundingClientRect().height}));
      samples.push({i,...state});if(i%10===0)console.log(JSON.stringify({mode,...samples.at(-1)}));
      if(state.transitions>=8)break;
    }
    const trace=await page.evaluate(()=>({transitions:window.__layoutTransitions,last:window.__lastLayoutTrace}));
    const report={method:'Read-only referenced source instrumentation, isolated Electron developer mode unless specified. Public original exact Word DOCX. Product mode adds no experimental CSS.',sourceCommit:execFileSync('git',['rev-parse','HEAD'],{cwd:root,encoding:'utf8'}).trim(),ref,mode,inputSHA256:row.inputSHA256,samples,...trace};
    await fs.writeFile(out+'/oscillation-'+mode+'.json.gz',z.gzipSync(JSON.stringify(report)));
    console.log(JSON.stringify({mode,pages:[...new Set(samples.map(s=>s.pages))],transitions:trace.transitions?.map(t=>t.pages)}));
  } finally {
    if(close)await close();else if(browser)await browser.close();
    await server.close();if(profile)await fs.rm(profile,{recursive:true,force:true});
  }
})().catch(e=>{console.error(e);process.exitCode=1});
