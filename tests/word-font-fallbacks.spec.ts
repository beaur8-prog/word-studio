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

test('missing Gill Sans MT uses measured local regular and italic widths without replacing installed originals',async({page})=>{
 await page.goto('/');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
 const value=await page.evaluate(async()=>{
  for(const name of ['Gill Sans MT','GillSansMT','GillSansMT-Bold','GillSansMT-Italic','GillSansMT-BoldItalic'])try{await new FontFace('Original Gill check',`local(${JSON.stringify(name)})`).load();return{skip:'Original Gill Sans MT is installed'};}catch{}
  const host=document.createElement('div');host.style.cssText='position:fixed;left:-10000px';document.body.append(host);const faces=[],rows=[];
  try{
   for(const[style,name]of[['normal','GillSans'],['italic','GillSans-Italic']]){
    const ref=new FontFace('WSGillReference_'+style,`local(${JSON.stringify(name)})`,{weight:'400',style,ascentOverride:'92.919921875%',descentOverride:'23.046875%',lineGapOverride:'0%',unicodeRange:'U+20-024F,U+2000-206F'});
    try{await ref.load();}catch{return{skip:'Measured local Gill Sans faces are unavailable'};}
    document.fonts.add(ref);faces.push(ref);const boxes=[];
    for(const family of ['Gill Sans MT',ref.family]){const span=document.createElement('span');span.style.cssText=`display:inline-block;font-family:"${family}";font-weight:400;font-style:${style};font-size:11pt;line-height:normal;white-space:pre`;span.textContent='Purchases and customer tables AV To';host.append(span);boxes.push(span);}
    await document.fonts.ready;const{wordFallbackLineHeight}=await import('/src/word-font-fallbacks.ts');rows.push({style,actual:boxes[0].getBoundingClientRect().width,expected:boxes[1].getBoundingClientRect().width,font:boxes[0].style.fontFamily,metric:wordFallbackLineHeight(getComputedStyle(boxes[0]),boxes[0].textContent!),expectedMetric:parseFloat(getComputedStyle(boxes[0]).fontSize)*2375/2048});
   }
   return{rows};
  }finally{host.remove();faces.forEach(f=>document.fonts.delete(f));}
 });
 test.skip(!!value.skip,value.skip);for(const row of value.rows!){expect(row.actual,row.style).toBeCloseTo(row.expected,2);expect(row.metric,row.style+' fractional line').toBeCloseTo(row.expectedMetric,4);expect(row.font).toBe('"Gill Sans MT"');}
});

test('regular-only fallbacks preserve every unmeasured weight and unsupported script through CSS matching',async({page})=>{
 await page.goto('/');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
 const result=await page.evaluate(async()=>{
  const{initializeRegularFontFallback,wordFallbackLineHeight}=await import('/src/word-font-fallbacks.ts');const family='WSPartialRegularSupplement',host=document.createElement('div');host.style.cssText='position:fixed;left:-10000px';document.body.append(host);const boxes=[];
  for(const weight of [100,300,400,500,600,700,900])for(const style of ['normal','italic'])for(const text of ['Source text AV To 123','한글ΩЖ😀']){const span=document.createElement('span');span.style.cssText=`display:inline-block;font-family:${family};font-weight:${weight};font-style:${style};font-size:16px;line-height:normal;white-space:pre`;span.textContent=text;host.append(span);boxes.push({span,weight,style,text});}
  const measure=()=>boxes.map(({span,...label})=>({...label,width:span.getBoundingClientRect().width,height:span.getBoundingClientRect().height,font:span.style.fontFamily,metric:wordFallbackLineHeight(getComputedStyle(span),span.textContent!)}));
  const before=measure(),profile={family,nativeNames:['NoNativeWSPartialRegular20261105'],faces:[{style:'normal' as const,source:'url(/src/assets/fonts/texgyreadventor-regular.otf)',ascent:'97.119140625%',descent:'22.021484375%',unicodeRange:'U+20-024F,U+2000-206F'},{style:'italic' as const,source:'url(/src/assets/fonts/texgyreadventor-italic.otf)',ascent:'97.119140625%',descent:'22.021484375%',unicodeRange:'U+20-024F,U+2000-206F'}]};
  try{const status=await initializeRegularFontFallback(profile);await document.fonts.ready;const after=measure(),count=[...document.fonts].filter(f=>f.family.replace(/["']/g,'')===family).length;const again=await initializeRegularFontFallback(profile),afterCount=[...document.fonts].filter(f=>f.family.replace(/["']/g,'')===family).length;return{status,again,count,afterCount,before,after};}finally{host.remove();}
 });
 expect(result.status).toBe('fallback');expect(result.again).toBe('fallback');expect(result.count).toBe(6);expect(result.afterCount).toBe(result.count);
 let changed=0;for(const[i,before]of result.before.entries()){const after=result.after[i];expect(after.font).toBe(before.font);if(before.weight!==400){expect(after.width).toBe(before.width);expect(after.height).toBe(before.height);expect(after.metric).toBeNull();}else if(before.text==='한글ΩЖ😀'){expect(after.width).toBe(before.width);expect(after.metric).toBeNull();}else{expect(after.metric).toBeCloseTo(16*2440/2048,5);if(after.width!==before.width)changed++;}}
 expect(changed).toBe(2);
});

test('a partial alias leaves installed fonts alone and never registers an incomplete pair of styles',async({page})=>{
 await page.goto('/');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
 const result=await page.evaluate(async()=>{
  const{initializeRegularFontFallback}=await import('/src/word-font-fallbacks.ts');let local='';for(const name of ['Arial','Liberation Sans','DejaVu Sans'])try{await new FontFace('Native priority probe',`local(${JSON.stringify(name)})`).load();local=name;break;}catch{}
  if(!local)throw Error('A known installed font is required');
  const host=document.createElement('span');host.style.cssText=`display:inline-block;font:16px "${local}"`;host.textContent='Installed face remains unchanged';document.body.append(host);const before=host.getBoundingClientRect().toJSON();
  const never='url(/must-not-fetch-partial-original.otf)',native=await initializeRegularFontFallback({family:local,nativeNames:[local],faces:[{style:'normal',source:never,ascent:'200%',descent:'80%',unicodeRange:'U+20-7E'},{style:'italic',source:never,ascent:'200%',descent:'80%',unicodeRange:'U+20-7E'}]});
  const after=host.getBoundingClientRect().toJSON(),fetched=performance.getEntriesByType('resource').some(r=>r.name.includes('must-not-fetch-partial-original'));host.remove();
  const family='WSIncompleteRegular',normal={style:'normal' as const,source:'url(/src/assets/fonts/texgyreadventor-regular.otf)',ascent:'97.119140625%',descent:'22.021484375%',unicodeRange:'U+20-024F'},italic={...normal,style:'italic' as const,source:'local("MissingWSRegularItalic20261105")'};
  const missing=await initializeRegularFontFallback({family,nativeNames:['MissingWSRegularOriginal20261105'],faces:[normal,italic]}),countAfterMissing=[...document.fonts].filter(f=>f.family.replace(/["']/g,'')===family).length;
  const retry=await initializeRegularFontFallback({family,nativeNames:['MissingWSRegularOriginal20261105'],faces:[normal,{...italic,source:'url(/src/assets/fonts/texgyreadventor-italic.otf)'}]}),countAfterRetry=[...document.fonts].filter(f=>f.family.replace(/["']/g,'')===family).length;
  return{native,before,after,fetched,missing,countAfterMissing,retry,countAfterRetry};
 });
 expect(result.native).toBe('original');expect(result.after).toEqual(result.before);expect(result.fetched).toBe(false);expect(result.missing).toBe('unavailable');expect(result.countAfterMissing).toBe(0);expect(result.retry).toBe('fallback');expect(result.countAfterRetry).toBe(6);
});
