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
  const value=await page.evaluate(async()=>{
    const {initializeWordFontFallbacks}=await import('/src/word-font-fallbacks.ts');
    let local='';for(const name of ['Arial','Liberation Sans','DejaVu Sans']){try{await new FontFace('Native check',`local("${name}")`).load();local=name;break}catch{}}
    if(!local)throw new Error('A known native sans font is required for this font priority check');
    const p=document.createElement('span');p.style.cssText=`font:11pt "${local}";position:fixed;white-space:pre`;p.textContent='Native AV To metrics';document.body.append(p);
    const before=p.getBoundingClientRect().toJSON();
    const registeredBefore=[...document.fonts].filter(f=>f.family===local).length;
    await initializeWordFontFallbacks([{family:local,faces:[{names:[local],weight:'400',style:'normal',url:'/must-not-download-installed-font.otf',ascent:'200%',descent:'80%'}]}]);
    const after=p.getBoundingClientRect().toJSON(),registeredAfter=[...document.fonts].filter(f=>f.family===local).length;
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
