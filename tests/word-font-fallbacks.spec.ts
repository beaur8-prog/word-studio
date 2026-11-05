import { test, expect } from '@playwright/test';

// These are font/interaction supplements, separate from the frozen Word originals.
test('missing Century Gothic keeps four real faces and the measured Word line box', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', {name:'문서 편집 영역',exact:true}).waitFor();
  const faces = await page.evaluate(async () => {
    const host=document.createElement('div'); host.style.cssText='position:fixed;left:-10000px;top:0'; document.body.append(host);
    const rows=[];
    try {
      for(const [name,weight,style,ascent] of [['regular','400','normal','97.119140625%'],['bold','700','normal','99.21875%'],['italic','400','italic','97.119140625%'],['bolditalic','700','italic','100.5859375%']] as const){
        let installed=false;
        try{await new FontFace('Native font check',`local("Century Gothic${weight==='700'?' Bold':''}${style==='italic'?' Italic':''}")`,{weight,style}).load();installed=true}catch{}
        if(installed){rows.push({name,installed});continue}
        const reference=new FontFace('ReferenceAdventor_'+name,`url(/src/assets/fonts/texgyreadventor-${name}.otf)`,{weight,style,ascentOverride:ascent,descentOverride:'22.021484375%',lineGapOverride:'0%'});
        await reference.load();document.fonts.add(reference);
        const boxes=[];
        for(const family of ['Century Gothic',reference.family]){
          const p=document.createElement('span');p.style.cssText=`display:inline-block;font-family:"${family}";font-weight:${weight};font-style:${style};font-size:11pt;line-height:normal;white-space:pre`;
          p.textContent='Building Association Rules and Regulations AV To ( )';host.append(p);boxes.push(p);
        }
        await document.fonts.ready;
        rows.push({name,installed,actual:boxes[0].getBoundingClientRect().toJSON(),expected:boxes[1].getBoundingClientRect().toJSON(),source:boxes[0].style.fontFamily});
      }
      return rows;
    }finally{host.remove()}
  });
  for(const face of faces){
    if(face.installed)continue;
    expect(face.source).toBe('"Century Gothic"');
    expect(face.actual!.width,face.name+' advance width').toBeCloseTo(face.expected!.width,2);
    expect(face.actual!.height,face.name+' normal line').toBeCloseTo(face.expected!.height,2);
  }
});

test('an installed face keeps its original metrics and never requests the fallback asset',async({page})=>{
  await page.goto('/');
  await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
  const value=await page.evaluate(async()=>{
    const {initializeWordFontFallbacks}=await import('/src/word-font-fallbacks.ts');
    let local='';for(const name of ['Arial','Liberation Sans','DejaVu Sans']){try{await new FontFace('Native check',`local("${name}")`).load();local=name;break}catch{}}
    if(!local)throw new Error('A known native sans font is required for this font priority check');
    const p=document.createElement('span');p.style.cssText=`font:11pt "${local}";position:fixed;white-space:pre`;p.textContent='Native AV To metrics';document.body.append(p);
    const before=p.getBoundingClientRect().toJSON();
    const registeredBefore=[...document.fonts].filter(f=>f.family.replace(/^["']|["']$/g,'')===local).length;
    await initializeWordFontFallbacks([{family:local,faces:[{names:[local],weight:'400',style:'normal',url:'/must-not-download-installed-font.otf',ascent:'200%',descent:'80%'}]}]);
    const after=p.getBoundingClientRect().toJSON(),registeredAfter=[...document.fonts].filter(f=>f.family.replace(/^["']|["']$/g,'')===local).length;
    p.remove();return{before,after,registeredBefore,registeredAfter,resources:performance.getEntriesByType('resource').filter(r=>r.name.includes('must-not-download')).length};
  });
  expect(value.after).toEqual(value.before);expect(value.registeredAfter).toBe(value.registeredBefore+1);expect(value.resources).toBe(0);
});

test('fallback loading preserves native editing history and the two standard DOCX font names',async({page})=>{
  await page.goto('/');
  const doc=await page.evaluate(async()=>{
    const{defaultPage}=await import('/src/types.ts');const{parseDocumentHtml}=await import('/src/extensions.ts');
    return{format:'word-studio',version:1,id:'font-roundtrip',title:'Font roundtrip',updatedAt:new Date().toISOString(),page:defaultPage,comments:[],content:parseDocumentHtml('<p data-word-line-rule="auto" style="font-family:Century Gothic;font-size:11pt;line-height:1.15;margin:0">Regular <strong>Bold</strong> <em>Italic</em> <strong><em>Both</em></strong> 한글 Ω Ж</p>')};
  });
  await page.locator('input[type=file]').first().setInputFiles({name:'fonts.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});
  const p=page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).locator('p').first();
  await expect(p).toContainText('Regular Bold Italic Both 한글 Ω Ж');
  await page.evaluate(()=>document.fonts.ready);
  await p.click();await page.keyboard.press('End');await page.keyboard.type(' EDIT');await expect(p).toContainText(' EDIT');
  await page.keyboard.press('ControlOrMeta+z');await expect(p).not.toContainText(' EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z');await expect(p).toContainText(' EDIT');
  const result=await page.evaluate(async()=>{
    const{exportDocx,importFile}=await import('/src/document-io.ts');const{parseDocumentHtml}=await import('/src/extensions.ts');const{defaultPage}=await import('/src/types.ts');const{default:JSZip}=await import('/node_modules/.vite/deps/jszip.js');
    const el=document.querySelector('.tiptap')!;let model={format:'word-studio',version:1,id:'fonts',title:'Fonts',updatedAt:new Date().toISOString(),page:defaultPage,comments:[],content:parseDocumentHtml(el.innerHTML)};
    const states=[];
    for(let i=0;i<2;i++){
      const blob=await exportDocx(model as any);const zip=await JSZip.loadAsync(await blob.arrayBuffer());const xml=await zip.file('word/document.xml')!.async('string');
      const opened=await importFile(new File([blob],'fonts.docx'));model={...opened.document,content:parseDocumentHtml(opened.html!)} as any;
      states.push({xml,model});
    }
    return states;
  });
  for(const state of result){expect(state.xml).toContain('Century Gothic');expect(state.xml).not.toContain('TeXGyre');expect(state.xml).not.toContain('Adventor');expect(JSON.stringify(state.model)).toContain('Century Gothic');expect(JSON.stringify(state.model)).toContain('EDIT')}
});

test('a partial native font family keeps its installed regular face beside a missing bold face',async({page})=>{
  await page.goto('/');
  await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
  const result=await page.evaluate(async()=>{
    const{initializeWordFontFallbacks}=await import('/src/word-font-fallbacks.ts');
    let local='';for(const name of ['Arial','Liberation Sans','DejaVu Sans']){try{await new FontFace('Native check',`local("${name}")`).load();local=name;break}catch{}}
    if(!local)throw new Error('A known installed sans font is required');
    await initializeWordFontFallbacks([{family:'MixedNativePriority',faces:[
      {names:[local],weight:'400',style:'normal',url:'/must-not-fetch-regular.otf',ascent:'200%',descent:'80%'},
      {names:['MissingBoldWS2026'],weight:'700',style:'normal',url:'/src/assets/fonts/texgyreadventor-bold.otf',ascent:'99.21875%',descent:'22.021484375%'},
    ]}]);
    const host=document.createElement('div');host.style.cssText='position:fixed;left:-10000px';document.body.append(host);
    const boxes=[];for(const family of [local,'MixedNativePriority']){const span=document.createElement('span');span.style.cssText=`display:inline-block;font:400 11pt "${family}";white-space:pre`;span.textContent='Native regular AV To';host.append(span);boxes.push(span)}
    await document.fonts.ready;const result=boxes.map(s=>({width:s.getBoundingClientRect().width,height:s.getBoundingClientRect().height}));host.remove();return result;
  });
  expect(result[1]).toEqual(result[0]);
});

test('measured fallback faces retain fractional Word leading without changing native or unrelated faces',async({page})=>{
 await page.goto('/');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
 const rows=await page.evaluate(async()=>{
  const{initializeWordFontFallbacks,wordFallbackLineHeight}=await import('/src/word-font-fallbacks.ts'),{wordLineHeight,layoutWordLines}=await import('/src/word-line-metrics.ts');
  const host=document.createElement('div');host.style.cssText='position:fixed;left:-10000px';document.body.append(host);const rows=[];
  try{
   for(const[name,weight,style,ascent]of[['regular','400','normal',1989],['bold','700','normal',2032],['italic','400','italic',1989],['bolditalic','700','italic',2060]]as const){
    const family='WSFractional_'+name;
    await initializeWordFontFallbacks([{family,faces:[{names:['NoInstalledWSFractionalFace20261105'],weight,style,url:`/src/assets/fonts/texgyreadventor-${name}.otf`,ascent:`${ascent/2048*100}%`,descent:`${451/2048*100}%`}]}]);
    const p=document.createElement('p');p.style.cssText=`font-family:${family};font-size:11.3pt;font-weight:${weight};font-style:${style};line-height:1.15;margin:0`;p.dataset.wordLineRule='auto';p.dataset.wordLine='1.15';p.textContent='Fractional line metrics';host.append(p);await document.fonts.load(`${style} ${weight} 11.3pt ${family}`);await document.fonts.ready;
    const css=getComputedStyle(p),expected=parseFloat(css.fontSize)*(ascent+451)/2048;
    rows.push({name,metric:wordFallbackLineHeight(css,p.textContent),line:wordLineHeight(p),expected,source:p.dataset.wordLine});
    p.dataset.wordLineRule='exact';p.style.lineHeight='8pt';layoutWordLines(host);rows[rows.length-1].exact=p.style.lineHeight;
   }
   const external=new FontFace('WSUnregisteredOverride','url(/src/assets/fonts/texgyreadventor-regular.otf)',{ascentOverride:'97.119140625%',descentOverride:'22.021484375%'});document.fonts.add(external);await external.load();
   const p=document.createElement('p');p.style.cssText='font-family:WSUnregisteredOverride;font-size:11.3pt';host.append(p);rows.push({name:'unrelated',metric:wordFallbackLineHeight(getComputedStyle(p),'Ag')});
   return rows;
  }finally{host.remove()}
 });
 for(const row of rows){if(row.name==='unrelated'){expect(row.metric).toBeNull();continue}expect(row.metric,row.name).toBeCloseTo(row.expected!,5);expect(row.line,row.name).toBeCloseTo(row.expected!*1.15,5);expect(row.source).toBe('1.15');expect(row.exact).toBe('8pt')}
});

test('an optional stalled local font probe cannot keep startup waiting indefinitely',async({page})=>{
 await page.goto('/');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
 const result=await page.evaluate(async()=>{
  const{initializeWordFontFallbacks}=await import('/src/word-font-fallbacks.ts');const original=globalThis.FontFace;
  const start=performance.now();
  try{
   globalThis.FontFace=class{load(){return new Promise(()=>{})}}as unknown as typeof FontFace;
   await initializeWordFontFallbacks([{family:'WSStalledOptionalFont',faces:[{names:['Stalled'],weight:'400',style:'normal',url:'/must-not-fetch-an-unresolved-probe.otf',ascent:'100%',descent:'20%'}]}]);
   return{elapsed:performance.now()-start,registered:[...document.fonts].some(face=>face.family.includes('WSStalledOptionalFont'))};
  }finally{globalThis.FontFace=original}
 });
 expect(result.elapsed).toBeLessThan(5000);expect(result.registered).toBe(false);
 await expect(page.getByRole('textbox',{name:'문서 편집 영역',exact:true})).toBeVisible();
});
