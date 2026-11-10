import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

test('cold print waits for its requested font and has the same pages and visible lines as warm print', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name:'문서 편집 영역',exact:true }).waitFor();
  await page.route(/texgyreadventor-(regular|bold)\.otf\?cold-print/, async route => {
    await new Promise(resolve=>setTimeout(resolve,500)); await route.continue();
  });
  const cold=await page.evaluate(async()=>{
    const {createPrintLayout}=await import('/src/print-layout.ts');
    const {defaultPage}=await import('/src/types.ts');
    const faces=[['regular','400'],['bold','700']].map(([name,weight])=>new FontFace('WS Deferred Print',
      'url("/src/assets/fonts/texgyreadventor-'+name+'.otf?cold-print")',{weight}));
    faces.forEach(face=>document.fonts.add(face));
    const paragraphs=Array.from({length:70},(_,i)=>'<p style="font-family:&quot;WS Deferred Print&quot;;font-size:12pt;line-height:normal;margin:0" data-word-line-rule="auto">'+
      'MARKER'+String(i).padStart(3,'0')+' The measured letters must stay visible across every page. <strong>Bold text</strong> and regular text share the same source.</p>');
    const html=paragraphs.join(''), settings={...defaultPage,size:'Custom' as const,width:140,height:130,margin:12,pageNumbers:false};
    const layout=await createPrintLayout(html,settings);
    const ready=faces.map(f=>f.status);
    await Promise.all(faces.map(f=>f.load())); await document.fonts.ready;
    const describe=(layout:any)=>({pages:layout.pages,paragraphs:[...layout.element.querySelectorAll('.print-content p')].map((p:any)=>{
      const r=p.getBoundingClientRect(),body=p.closest('.print-content').getBoundingClientRect();
      return {text:p.textContent,page:Number(p.closest('.print-page').dataset.page),top:r.top-body.top,bottom:r.bottom-body.top,body:body.height};
    })});
    (window as any).fontCold={layout,html,settings,describe};
    layout.element.classList.replace('measuring','preview');document.body.classList.add('studio-printing');
    return {ready,state:describe(layout)};
  });
  await page.pdf({path:test.info().outputPath('cold.pdf'),preferCSSPageSize:true,printBackground:true});
  const warm=await page.evaluate(async()=>{
    const {createPrintLayout}=await import('/src/print-layout.ts'),state=(window as any).fontCold;
    state.layout.dispose();
    const layout=await createPrintLayout(state.html,state.settings);
    layout.element.classList.replace('measuring','preview');
    return state.describe(layout);
  });
  await page.pdf({path:test.info().outputPath('warm.pdf'),preferCSSPageSize:true,printBackground:true});
  await writeFile(test.info().outputPath('font-layout-states.json'),JSON.stringify({cold,warm},null,2));
  expect(cold.ready).toEqual(['loaded','loaded']);
  expect(cold.state.pages).toBe(warm.pages);
  expect(cold.state.paragraphs).toEqual(warm.paragraphs);
  expect(cold.state.paragraphs.every(p=>p.top>=-.6&&p.bottom<=p.body+.6)).toBe(true);
});

test('font preparation includes footnotes, section furniture and empty page-number fields', async ({ page })=>{
  await page.goto('/');
  await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
  await page.route(/texgyreadventor-regular\.otf\?story-/,async route=>{
    await new Promise(resolve=>setTimeout(resolve,300));await route.continue();
  });
  const result=await page.evaluate(async()=>{
    const {createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
    const names=['WS Note Font','WS Header Font','WS Section Font','WS Field Font'];
    const faces=names.map((name,index)=>new FontFace(name,'url("/src/assets/fonts/texgyreadventor-regular.otf?story-'+index+'")',
      index===3?{unicodeRange:'U+30-39'}:{}));
    faces.forEach(face=>document.fonts.add(face));
    const page1={...defaultPage,differentFirstPage:true,firstHeaderHtml:'<p style="font-family:&quot;WS Header Font&quot;">First header</p>',
      footerHtml:'<span data-page-number="PAGE" style="font-family:&quot;WS Field Font&quot;"></span>'};
    const page2={...defaultPage,headerHtml:'<p style="font-family:&quot;WS Section Font&quot;">Section header</p>'};
    const source=document.createElement('div');
    source.innerHTML='<p>Body note <sup data-note-reference="n">1</sup></p><aside class="document-note" data-document-note="n" data-note-kind="footnote"><p style="font-family:&quot;WS Note Font&quot;">Note text</p></aside>';
    const boundary=document.createElement('div');boundary.dataset.sectionPage=JSON.stringify(page2);source.append(boundary);
    const after=document.createElement('p');after.textContent='Second section body';source.append(after);
    const layout=await createPrintLayout(source.innerHTML,page1);
    try{return {statuses:faces.map(face=>face.status),pages:layout.pages,notes:layout.element.querySelector('.print-footnotes')?.textContent,
      headers:[...layout.element.querySelectorAll('.print-header')].map(n=>n.textContent)};}
    finally{layout.dispose();}
  });
  expect(result.statuses).toEqual(Array(4).fill('loaded'));
  expect(result.pages).toBe(2);expect(result.notes).toContain('Note text');
  expect(result.headers).toEqual(['First header','Section header']);
});

test('cancelling while a requested print font loads removes the temporary layout promptly',async({page})=>{
  await page.goto('/');
  await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
  await page.route(/texgyreadventor-italic\.otf\?abort-print/,async route=>{
    await new Promise(resolve=>setTimeout(resolve,1000));await route.continue().catch(()=>{});
  });
  const result=await page.evaluate(async()=>{
    const {createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
    const face=new FontFace('WS Cancel Font','url("/src/assets/fonts/texgyreadventor-italic.otf?abort-print")');document.fonts.add(face);
    const controller=new AbortController(),started=performance.now();setTimeout(()=>controller.abort(),50);
    try{
      const layout=await createPrintLayout('<p style="font-family:&quot;WS Cancel Font&quot;">Cancel this print</p>',defaultPage,undefined,controller.signal);
      layout.dispose();return {aborted:false,elapsed:performance.now()-started,layouts:document.querySelectorAll('.print-layout').length,probes:document.querySelectorAll('.print-font-probe').length};
    }catch(error:any){
      return {aborted:error.name==='AbortError',elapsed:performance.now()-started,layouts:document.querySelectorAll('.print-layout').length,probes:document.querySelectorAll('.print-font-probe').length};
    }finally{document.fonts.delete(face);}
  });
  expect(result.aborted).toBe(true);expect(result.elapsed).toBeLessThan(600);
  expect(result.layouts).toBe(0);expect(result.probes).toBe(0);
});

test('an actual web-font failure uses fallback text and reports it instead of losing the print',async({page})=>{
  await page.goto('/');
  await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
  await page.route('**/failed-print-font.woff2',route=>route.abort('failed'));
  const result=await page.evaluate(async()=>{
    const {createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
    const face=new FontFace('WS Failed Font','url("/failed-print-font.woff2")');document.fonts.add(face);
    const layout=await createPrintLayout('<p style="font-family:&quot;WS Failed Font&quot;,Arial;">Fallback text remains visible.</p>',defaultPage);
    try{return {status:face.status,pages:layout.pages,warnings:layout.warnings,text:layout.element.querySelector('.print-content')?.textContent};}
    finally{layout.dispose();document.fonts.delete(face);}
  });
  expect(result.status).toBe('error');expect(result.pages).toBe(1);
  expect(result.text).toContain('Fallback text remains visible.');
  expect(result.warnings.join(' ')).toContain('글꼴');
});
