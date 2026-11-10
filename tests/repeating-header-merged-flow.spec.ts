import { test, expect } from '@playwright/test';

for (const alignment of ['top', 'middle', 'bottom']) for (const continuing of [true, false])
test('repeated header with ' + alignment + ' vertical merge ' + (continuing ? 'continues at rows' : 'keeps its protected block') + ' after preceding text', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const states = await page.evaluate(async ({ alignment, continuing }) => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const p = (text: string, height = 20) => '<p style="font-family:Arial;font-size:12pt;line-height:' + height + 'px;margin:0">' + text + '</p>';
    const table = '<table style="width:400px;margin:0"><colgroup><col style="width:140px"><col style="width:260px"></colgroup>' +
      '<tr data-row-split="avoid" style="height:30px"><th colwidth="140" style="padding:2px;border:1px solid black">' + p('HEADER LEFT') + '</th><th colwidth="260" style="padding:2px;border:1px solid black">' + p('HEADER RIGHT') + '</th></tr>' +
      Array.from({length:5}, (_, i) => '<tr data-row-split="' + (continuing ? 'allow' : 'avoid') + '" style="height:80px">' +
        (i ? '' : '<td colwidth="140" rowspan="5" ' + (continuing ? 'data-cell-split="allow"' : '') + ' style="padding:2px;border:1px solid black;vertical-align:' + alignment + '">' + p('LABEL A') + p('LABEL B') + '</td>') +
        '<td colwidth="260" style="padding:2px;border:1px solid black">' + p('BODY ROW ' + i) + '</td></tr>').join('') + '</table>';
    const loaded = await importFile(new File([p('LEAD',450) + table + p('TAIL',200)], 'header-merge.html'));
    let doc = { ...loaded.document, content:parseDocumentHtml(loaded.html!),
      page:{...defaultPage,size:'Custom',width:150,height:180,margin:10,pageNumbers:false}};
    const result = [];
    for (let save=0;save<3;save++) {
      doc.title = 'MERGED HEADER ' + alignment + ' ' + continuing + ' ' + save;
      const layout = await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
      try {
        const starts = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].filter(p=>!p.closest('th')).map(p=>{
          const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);walker.nextNode();
          const range=document.createRange();range.setStart(walker.currentNode,0);range.setEnd(walker.currentNode,1);
          const box=range.getBoundingClientRect(),paper=p.closest<HTMLElement>('.print-page')!,origin=paper.getBoundingClientRect();
          return {text:p.textContent!,page:Number(paper.dataset.page),y:box.top-origin.top};
        });
        result.push({doc,pages:layout.pages,starts,headers:layout.element.querySelectorAll('th').length});
      } finally {layout.dispose();}
      if (save<2) {const next=await importFile(new File([await exportDocx(doc)],'saved.docx'));doc={...next.document,content:parseDocumentHtml(next.html!)};}
    }
    return result;
  }, {alignment,continuing});
  for (const state of states) {
    expect(state.starts).toHaveLength(9);
    expect(state.starts.map(p=>p.text).sort()).toEqual(['LEAD','TAIL','LABEL A','LABEL B',...Array.from({length:5},(_,i)=>'BODY ROW '+i)].sort());
    expect(state.pages).toBe(continuing ? 2 : 3);
    expect(state.starts.find(p=>p.text==='BODY ROW 0')?.page).toBe(continuing ? 1 : 2);
    await page.locator('input[type=file]').first().setInputFiles({name:'header-merge.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(state.doc))});
    await expect(page.getByRole('textbox',{name:'문서 제목',exact:true})).toHaveValue(state.doc.title);
    for (const zoom of [100,50,125,200,100]) {
      const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');
      await expect.poll(()=>page.evaluate(({zoom,expected})=>{
        const root=document.querySelector<HTMLElement>('.tiptap[contenteditable=true]')!,sheet=root.closest<HTMLElement>('.document-stack')!;
        const origin=root.getBoundingClientRect().top,scale=zoom/100,stride=Number(sheet.dataset.pageHeight)+24;
        const actual=[...root.querySelectorAll<HTMLElement>('p')].filter(p=>!p.closest('th,.ProseMirror-widget,.pagination-headers,.section-papers')).map(p=>{
          const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);walker.nextNode();
          if(walker.currentNode===p)return null;
          const range=document.createRange();range.setStart(walker.currentNode,0);range.setEnd(walker.currentNode,1);
          const yy=(range.getBoundingClientRect().top-origin)/scale,page=Math.max(1,1+Math.floor((yy+.5)/stride));
          return {text:p.textContent!,page,y:yy-(page-1)*stride};
        }).filter(Boolean);
        return {pages:Number(sheet.dataset.pages),problems:actual.flatMap(a=>{
          const e=expected.find(e=>e.text===a!.text);return !e||e.page!==a!.page||Math.abs(e.y-a!.y)>1.25?[{actual:a,expected:e}]:[];
        })};
      },{zoom,expected:state.starts})).toEqual({pages:state.pages,problems:[]});
    }
    const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true}),before=await editor.evaluate(el=>(el as any).editor.getJSON());
    await editor.locator('td p').filter({hasText:'BODY ROW 0'}).click();await page.keyboard.press('End');await page.keyboard.insertText(' EDIT');
    await expect(editor).toContainText('BODY ROW 0 EDIT');await page.keyboard.press('ControlOrMeta+z');
    await expect.poll(()=>editor.evaluate(el=>(el as any).editor.getJSON())).toEqual(before);
    await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor).toContainText('BODY ROW 0 EDIT');
    await page.keyboard.press('ControlOrMeta+z');await expect.poll(()=>editor.evaluate(el=>(el as any).editor.getJSON())).toEqual(before);
  }
});
