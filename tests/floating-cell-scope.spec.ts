import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

for (const layoutInCell of [true, false]) test(`top-bottom floating picture ${layoutInCell ? 'stays in its cell' : 'keeps its explicit page-wide exclusion'} through two saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async layoutInCell => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { defaultPage } = await import('/src/types.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const floating={horizontalPosition:{relative:'column',offset:0},verticalPosition:{relative:'paragraph',offset:0},wrap:{type:3},layoutInCell,behindDocument:false,allowOverlap:true};
    const html=`<table style="width:500px"><colgroup><col style="width:250px"><col style="width:250px"></colgroup><tbody><tr><td colwidth="250" style="width:250px;vertical-align:top;padding:4px"><p style="font-family:Arial;font-size:11pt;line-height:20px;margin:0">LEFT CONTROL</p></td><td colwidth="250" style="width:250px;vertical-align:top;padding:4px"><p style="font-family:Arial;font-size:11pt;line-height:20px;margin:0"><img data-inline-image="true" data-image-floating="${JSON.stringify(floating).replaceAll('"','&quot;')}" src="data:image/png;base64,${png}" width="180" height="100" alt="CELL PICTURE">RIGHT CONTENT</p></td></tr></tbody></table><p style="line-height:20px;margin:0">AFTER TABLE</p>`;
    let doc:any={format:'word-studio',version:1,id:'cell-scope',title:'Cell scope',updatedAt:new Date().toISOString(),comments:[],page:{...defaultPage,pageNumbers:false},content:parseDocumentHtml(html)};
    const states=[],saves=[];
    for(let pass=0;pass<3;pass++){
      const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
      const img=layout.element.querySelector<HTMLImageElement>('img[alt="CELL PICTURE"]')!;
      const glyph=(label:string)=>{const p=[...layout.element.querySelectorAll('p')].find(p=>p.textContent===label)!,walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);while(walker.nextNode()){const n=walker.currentNode as Text;if(!n.data.trim())continue;const range=document.createRange();range.setStart(n,0);range.setEnd(n,1);const b=range.getBoundingClientRect();return{page:Number(p.closest<HTMLElement>('.print-page')!.dataset.page),y:b.y-p.closest('.print-page')!.getBoundingClientRect().y};}throw Error('Missing text '+label);};
      const box=img.getBoundingClientRect(),paper=img.closest('.print-page')!.getBoundingClientRect();
      states.push({doc,pages:layout.pages,left:glyph('LEFT CONTROL'),right:glyph('RIGHT CONTENT'),after:glyph('AFTER TABLE'),picture:{top:box.top-paper.top,bottom:box.bottom-paper.top,width:box.width,height:box.height},geometry:JSON.parse(img.dataset.imageFloating!)});layout.dispose();
      if(pass<2){const output=await exportDocx(doc);saves.push([...new Uint8Array(await output.arrayBuffer())]);const opened=await importFile(new File([output],'cell-scope.docx'));doc={...opened.document,content:parseDocumentHtml(opened.html!)};}
    }
    return{states,saves};
  },layoutInCell);
  for(const[i,bytes]of result.saves.entries())await writeFile(test.info().outputPath(`textbox-save-cell-scope-${layoutInCell}-${i+1}.docx`),Buffer.from(bytes));
  await writeFile(test.info().outputPath('cell-scope-coordinates.json'),JSON.stringify(result.states.map(({doc,...rest})=>rest),null,2));
  for(const state of result.states){
    expect(state.pages).toBe(1);expect(state.geometry.layoutInCell).toBe(layoutInCell);
    expect(state.picture.width).toBeCloseTo(180,0);expect(state.picture.height).toBeCloseTo(100,0);
    expect(state.right.y).toBeGreaterThanOrEqual(state.picture.bottom-1);
    if(layoutInCell)expect(state.picture.top-state.left.y).toBeGreaterThan(-20);
    else expect(state.left.y).toBeGreaterThanOrEqual(state.picture.bottom-1);
    expect(state.after.y).toBeGreaterThan(state.right.y);
    await page.locator('input[type=file]').first().setInputFiles({name:'cell-scope.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(state.doc))});
    const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText('RIGHT CONTENT');
    await expect.poll(()=>page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
    for(const zoom of [100,50,125,100]){const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');
      await expect.poll(async()=>editor.evaluate((root,printed)=>{let scale=1;for(let a:Element|null=root;a;a=a.parentElement){const c=getComputedStyle(a);scale*=parseFloat(c.zoom)||1;if(c.transform!=='none'){const m=new DOMMatrixReadOnly(c.transform);scale*=Math.hypot(m.a,m.b);}}const origin=root.getBoundingClientRect(),ys=[];for(const text of ['LEFT CONTROL','RIGHT CONTENT']){const p=[...root.querySelectorAll('p')].find(p=>p.textContent===text)!,walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);while(walker.nextNode()){const n=walker.currentNode as Text;if(!n.data.trim()||n.parentElement?.closest('[data-flow-widget]'))continue;const range=document.createRange();range.setStart(n,0);range.setEnd(n,1);ys.push((range.getBoundingClientRect().top-origin.top)/scale);break;}}return Math.max(Math.abs(ys[0]-printed[0]),Math.abs(ys[1]-printed[1]));},[state.left.y,state.right.y]),{timeout:8000}).toBeLessThan(1.25);
    }
    const before=await editor.evaluate((root:any)=>root.editor.getJSON());const left=editor.locator('p').filter({hasText:/^LEFT CONTROL/});await left.fill('LEFT CONTROL EDIT');await expect(left).toHaveText('LEFT CONTROL EDIT');await page.keyboard.press('ControlOrMeta+z');await expect.poll(()=>editor.evaluate((root:any)=>root.editor.getJSON())).toEqual(before);
  }
});
