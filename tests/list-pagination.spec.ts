import {test,expect,type Page} from '@playwright/test';
import fs from 'node:fs/promises';
import {Document,Packer,Paragraph,LevelFormat} from 'docx';

async function loadList(page:Page, html:string) {
  await page.goto('/');
  const state=await page.evaluate(async html=>{
    const {defaultPage}=await import('/src/types.ts'),{importFile}=await import('/src/document-io.ts');
    const {parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js');
    const {createPrintLayout}=await import('/src/print-layout.ts');
    const opened=await importFile(new File([html],'list.html'));
    const doc={...opened.document,page:{...defaultPage,size:'Custom',width:150,height:180,margin:10},content:parseDocumentHtml(opened.html!)};
    const print=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
    try{return{doc,pages:[...print.element.querySelectorAll('.print-content')].map(p=>p.textContent!.replace(/\s/g,''))}}
    finally{print.dispose()}
  },html);
  await page.locator('input[type=file]').first().setInputFiles({name:'list.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(state.doc))});
  return state;
}
async function editorPages(page:Page,zoom=100) {
  return page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).evaluate((root,zoom)=>{
    const sheet=root.closest<HTMLElement>('.document-stack')!,origin=root.getBoundingClientRect().top;
    const h=Number(sheet.dataset.pageHeight),top=Number(sheet.dataset.margin),bottom=Number(sheet.dataset.bottomMargin),pages:string[]=[];
    const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);
    while(walker.nextNode()){
      const n=walker.currentNode as Text;if(n.parentElement?.closest('.pagination-gap,.pagination-cell-gap,[data-flow-widget]'))continue;
      for(const s of new Intl.Segmenter(undefined,{granularity:'grapheme'}).segment(n.data)){
        if(/^\s+$/.test(s.segment))continue;const r=document.createRange();r.setStart(n,s.index);r.setEnd(n,s.index+s.segment.length);const b=r.getBoundingClientRect(),y=(b.top-origin)/(zoom/100);
        const p=Math.max(0,Math.floor((y+.5)/(h+24))),local=y-p*(h+24);
        if(local<top-1||local+b.height/(zoom/100)>h-bottom+1)return['INK IN MARGIN'];
        pages[p]=(pages[p]||'')+s.segment;
      }
    }
    return pages;
  },zoom);
}
const p=(text:string)=>`<p style="font-size:12px;line-height:24px;margin:0">${text}</p>`;
const lead='<p style="line-height:400px;margin:0">LEAD</p>';

test('a list continues at item boundaries instead of moving the whole list to the next sheet',async({page})=>{
  const state=await loadList(page,lead+'<ol start="4">'+Array.from({length:18},(_,i)=>'<li>'+p(`ITEM${String(i).padStart(2,'0')} words on one line` )+'</li>').join('')+'</ol>'+p('END'));
  expect(state.pages.length).toBeGreaterThan(1);expect(state.pages[0]).toContain('ITEM00');
  for(const zoom of [100,50,125,200,100]){
    const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');
    await expect.poll(()=>editorPages(page,zoom)).toEqual(state.pages);
  }
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await editor.click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowDown':'Control+End');await page.keyboard.type('EDIT');await expect(editor).toContainText('EDIT');await page.keyboard.press('ControlOrMeta+z');await expect(editor).not.toContainText('EDIT');
  await expect.poll(()=>editorPages(page)).toEqual(state.pages);
  const download=page.waitForEvent('download');await page.keyboard.press('ControlOrMeta+s');const saved=JSON.parse(await fs.readFile((await(await download).path())!,'utf8'));expect(saved.content).toEqual(state.doc.content);
});

for(const kind of ['long-paragraph','nested-items'] as const)test(`${kind}: an item taller than a sheet continues without shrinking text or repeating its parent number`,async({page})=>{
  const body=kind==='long-paragraph'?p(Array.from({length:340},(_,i)=>`TOKEN${String(i).padStart(3,'0')} `).join('')):p('PARENT')+'<ol start="3">'+Array.from({length:40},(_,i)=>'<li>'+p(`CHILD${String(i).padStart(2,'0')}`)+'</li>').join('')+'</ol>';
  const state=await loadList(page,lead+'<ol start="5"><li>'+body+'</li><li>'+p('FOLLOWING ITEM')+'</li></ol>'+p('END'));
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});
  for(const zoom of [100,50,125,200,100]){
    const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');
    await expect.poll(()=>editorPages(page,zoom)).toEqual(state.pages);
  }
  await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();
  await expect.poll(()=>page.locator('.print-page').count()).toBeGreaterThan(1);
  expect(await page.locator('.print-content [style*="zoom"]').count()).toBe(0);
  expect(await page.locator('.print-content').allTextContents().then(a=>a.join('').replace(/\s/g,''))).toBe(state.pages.join(''));
  const visible=await page.locator('.print-content > ol').evaluateAll(lists=>lists.flatMap(ol=>[...ol.children].flatMap((li,i)=>(li as HTMLElement).dataset.printListContinuation?[]:[{number:(Number(ol.getAttribute('start'))||1)+i,text:li.textContent}])));
  expect(visible.map(x=>x.number)).toEqual([5,6]);
  expect(visible[1].text).toBe('FOLLOWING ITEM');
  if(kind==='nested-items'){
    const nested=await page.locator('.print-content > ol > li > ol').evaluateAll(lists=>lists.flatMap(ol=>[...ol.children].map((li,i)=>({number:(Number(ol.getAttribute('start'))||1)+i,text:li.textContent}))));
    expect(nested.map(x=>x.number)).toEqual(Array.from({length:40},(_,i)=>i+3));
    expect(nested.map(x=>x.text)).toEqual(Array.from({length:40},(_,i)=>`CHILD${String(i).padStart(2,'0')}`));
  }
  await page.getByRole('button',{name:'편집으로 돌아가기',exact:true}).click();
  const download=page.waitForEvent('download');await page.keyboard.press('ControlOrMeta+s');const saved=JSON.parse(await fs.readFile((await(await download).path())!,'utf8'));expect(saved.content).toEqual(state.doc.content);
  expect(await editor.textContent()).toBeTruthy();
});

test('imported numbering updates immediately after Enter, prints the same labels, and undo restores the original model',async({page})=>{
  const attrs='data-list-id="9" data-list-level="0" data-list-format="decimal" data-list-pattern="%1."';
  const state=await loadList(page,`<ol start="4"><li ${attrs} data-list-label="4.">${p('FIRST')}</li><li ${attrs} data-list-label="5.">${p('SECOND')}</li></ol>${p('END')}`);
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});
  await editor.locator('p').first().click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowRight':'End');await page.keyboard.press('Enter');await page.keyboard.type('INSERTED');
  await expect.poll(()=>editor.locator('li').evaluateAll(items=>items.map(li=>li.getAttribute('data-list-label')))).toEqual(['4.','5.','6.']);
  await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();await expect.poll(()=>page.locator('.print-content li').count()).toBe(3);
  expect(await page.locator('.print-content li').evaluateAll(items=>items.map(li=>li.getAttribute('data-list-label')))).toEqual(['4.','5.','6.']);
  await page.getByRole('button',{name:'편집으로 돌아가기',exact:true}).click();
  await page.keyboard.press('ControlOrMeta+z');await expect(editor).not.toContainText('INSERTED');
  const download=page.waitForEvent('download');await page.keyboard.press('ControlOrMeta+s');const saved=JSON.parse(await fs.readFile((await(await download).path())!,'utf8'));expect(saved.content).toEqual(state.doc.content);
});


test('edited standard DOCX list keeps current numbering through two standard DOCX saves',async({page},testInfo)=>{
  const source=await Packer.toBuffer(new Document({numbering:{config:[{reference:'Steps',levels:[{level:0,start:4,format:LevelFormat.DECIMAL,text:'%1.'}]}]},sections:[{children:[
    new Paragraph({text:'FIRST',numbering:{reference:'Steps',level:0}}),new Paragraph({text:'SECOND',numbering:{reference:'Steps',level:0}}),new Paragraph({text:'END'}),
  ]}]}));
  await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'steps.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:source});
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText('FIRST');await editor.locator('p').first().click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowRight':'End');await page.keyboard.press('Enter');await page.keyboard.type('INSERTED');
  await expect.poll(()=>editor.locator('li').evaluateAll(items=>items.map(li=>li.getAttribute('data-list-label')))).toEqual(['4.','5.','6.']);
  const download=page.waitForEvent('download');await page.keyboard.press('ControlOrMeta+s');const model=JSON.parse(await fs.readFile((await(await download).path())!,'utf8'));
  const result=await page.evaluate(async model=>{
    const{exportDocx,importFile}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');const saves=[],stages=[];
    for(let i=0;i<2;i++){const bytes=[...new Uint8Array(await(await exportDocx(model)).arrayBuffer())];saves.push(bytes);const opened=await importFile(new File([new Uint8Array(bytes)],'steps.docx'));model={...opened.document,content:parseDocumentHtml(opened.html!)};
      const root=document.createElement('div');root.innerHTML=opened.html!;stages.push({text:root.textContent,labels:[...root.querySelectorAll('li')].map(li=>li.dataset.listLabel)});}
    return{saves,stages};
  },model);
  for(const stage of result.stages){expect(stage.labels).toEqual(['4.','5.','6.']);expect(stage.text).toBe('FIRSTINSERTEDSECONDEND')}
  for(let i=0;i<2;i++)await fs.writeFile(testInfo.outputPath(`list-save${i+1}.docx`),Buffer.from(result.saves[i]));
});

for(const kind of ['hard-breaks','mixed-height'] as const)test(`${kind}: list line continuation preserves authored breaks and graphemes`,async({page})=>{
  const body=kind==='hard-breaks'?Array.from({length:64},(_,i)=>`줄${String(i).padStart(2,'0')}🙂e\u0301`).join('<br><br>')+'<br><br>':Array.from({length:340},(_,i)=>i%15===0?`<strong style="font-size:18px">LARGE${i}</strong> `:`word${i} `).join('');
  const content=kind==='mixed-height'?`<p style="font-size:12px;line-height:1.5;margin:0;text-indent:12px">${body}</p>`:p(body);
  const state=await loadList(page,lead+'<ul><li>'+content+'</li><li>'+p('AFTER')+'</li></ul>'+p('END'));
  await expect.poll(()=>editorPages(page)).toEqual(state.pages);
  const download=page.waitForEvent('download');await page.keyboard.press('ControlOrMeta+s');const saved=JSON.parse(await fs.readFile((await(await download).path())!,'utf8'));expect(saved.content).toEqual(state.doc.content);
});

test('list continuation keeps input visible during composition and preserves split/join undo and redo',async({page})=>{
  const state=await loadList(page,lead+'<ol>'+Array.from({length:18},(_,i)=>'<li>'+p(`ITEM${i}`)+'</li>').join('')+'</ol>'+p('END'));
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect.poll(()=>editorPages(page)).toEqual(state.pages);
  const last=editor.locator('li p').last();await last.click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowRight':'End');await page.keyboard.press('Enter');
  await page.waitForTimeout(650);await editor.dispatchEvent('compositionstart',{data:''});await page.keyboard.insertText('한글🙂');await expect(editor).toContainText('한글🙂');await editor.dispatchEvent('compositionend',{data:'한글🙂'});
  await expect(editor.locator('li')).toHaveCount(19);
  await page.keyboard.press('ControlOrMeta+z');await expect(editor).not.toContainText('한글🙂');await page.keyboard.press('ControlOrMeta+z');await expect(editor.locator('li')).toHaveCount(18);
  await expect.poll(()=>editorPages(page)).toEqual(state.pages);
  await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor.locator('li')).toHaveCount(19);await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor).toContainText('한글🙂');
  await page.waitForTimeout(650);await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowLeft':'Home');await page.keyboard.press('Backspace');await expect(editor.locator('li')).toHaveCount(18);await expect(editor.locator('li p').last()).toHaveText('ITEM17한글🙂');
  await page.keyboard.press('ControlOrMeta+z');await expect(editor.locator('li')).toHaveCount(19);await expect(editor).toContainText('한글🙂');
  await page.keyboard.press('ControlOrMeta+z');await expect(editor).not.toContainText('한글🙂');await page.keyboard.press('ControlOrMeta+z');await expect(editor.locator('li')).toHaveCount(18);
  await expect.poll(()=>editorPages(page)).toEqual(state.pages);
  const download=page.waitForEvent('download');await page.keyboard.press('ControlOrMeta+s');const saved=JSON.parse(await fs.readFile((await(await download).path())!,'utf8'));expect(saved.content).toEqual(state.doc.content);
});
