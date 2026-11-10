import {test,expect}from'@playwright/test';
import{writeFile}from'node:fs/promises';
for(const kind of ['cell','cell-inline','cell-outside','paper','partial-paper'] as const)test(`off-paper ${kind} drawing preserves its ownership, pixels and source coordinates through two saves`,async({page})=>{
 await page.goto('/');
 const result=await page.evaluate(async kind=>{
  const{exportDocx,importFile}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{defaultPage}=await import('/src/types.ts'),{generateHTML}=await import('/node_modules/@tiptap/core/dist/index.js'),{createPrintLayout}=await import('/src/print-layout.ts');
  const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
  const floating={horizontalPosition:{relative:'page',offset:(kind==='partial-paper'?780:1500)*9525},verticalPosition:{relative:'page',offset:60*9525},wrap:{type:1,side:'bothSides'},layoutInCell:kind!=='cell-outside',behindDocument:false,allowOverlap:true};
  const picture=`<p style="font-family:Arial;font-size:11pt;line-height:20px;margin:0"><img data-inline-image="true" src="data:image/png;base64,${png}" width="96" height="24" alt="OWNED LOGO" data-image-floating="${JSON.stringify(floating).replaceAll('"','&quot;')}">${kind==='cell-inline'?`<img data-inline-image="true" src="data:image/png;base64,${png}" width="80" height="20" alt="INLINE HEADER">`:''}</p>`;
  const box=kind.startsWith('cell')?`<table style="width:320px"><colgroup><col style="width:320px"></colgroup><tr style="height:40px"><td colwidth="320" style="width:320px;padding:4px;vertical-align:middle">${picture}</td></tr></table>`:picture;
  let doc:any={format:'word-studio',version:1,id:'cell-offpaper',title:'Cell offpaper',updatedAt:new Date().toISOString(),comments:[],page:{...defaultPage,pageNumbers:false},content:parseDocumentHtml('<p style="font-family:Arial;font-size:11pt;line-height:180px;margin:0">FIRST</p>'+box+'<p style="font-family:Arial;font-size:11pt;line-height:20px;margin:0">AFTER</p>')};
  const states=[],saves=[];
  for(let pass=0;pass<3;pass++){
   const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page),img=layout.element.querySelector<HTMLImageElement>('img[alt="OWNED LOGO"]')!,paper=img.closest<HTMLElement>('.print-page')!,origin=paper.getBoundingClientRect(),rect=img.getBoundingClientRect(),cell=img.closest('td,th')?.getBoundingClientRect();
   states.push({doc,pages:layout.pages,display:getComputedStyle(img).display,image:{x:rect.x-origin.x,y:rect.y-origin.y,width:rect.width,height:rect.height},cell:cell&&{x:cell.x-origin.x,y:cell.y-origin.y,width:cell.width,height:cell.height},floating:JSON.parse(img.dataset.imageFloating!),src:img.getAttribute('src'),paperWidth:origin.width});layout.dispose();
   if(pass<2){const out=await exportDocx(doc);saves.push([...new Uint8Array(await out.arrayBuffer())]);const loaded=await importFile(new File([out],'again.docx'));doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};}
  }
  return{states,saves};
 },kind);
 for(const[i,bytes]of result.saves.entries())await writeFile(test.info().outputPath(`textbox-save-offpaper-${kind}-${i+1}.docx`),Buffer.from(bytes));
 await writeFile(test.info().outputPath('offpaper-cell-geometry.json'),JSON.stringify(result.states.map(({doc,...rest})=>rest),null,2));
 for(const state of result.states){
  expect(state.pages).toBe(1);expect(state.floating).toMatchObject(result.states[0].floating);expect(state.src).toBe(result.states[0].src);
  if(kind==='cell'||kind==='cell-inline'){
   expect(state.display).not.toBe('none');expect(state.image.width).toBeCloseTo(96,1);expect(state.image.height).toBeCloseTo(24,1);
   expect(state.image.x).toBeGreaterThanOrEqual(state.cell!.x);expect(state.image.x+96).toBeLessThanOrEqual(state.cell!.x+state.cell!.width+.5);expect(state.image.y).toBeGreaterThanOrEqual(state.cell!.y);expect(state.image.y+24).toBeLessThanOrEqual(state.cell!.y+state.cell!.height+.5);
  }else if(kind==='partial-paper'){expect(state.display).not.toBe('none');expect(state.image.x).toBeGreaterThan(state.paperWidth-30);expect(state.image.width).toBeCloseTo(96,1);}
  else{expect(state.display).toBe('none');expect(state.image.width).toBe(0);}
  await page.locator('input[type=file]').first().setInputFiles({name:'offpaper.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(state.doc))});
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText('AFTER');
  for(const zoom of [100,50,125,100]){const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');await expect.poll(()=>page.locator('.document-stack').getAttribute('data-pages')).toBe('1');if(kind==='cell'||kind==='cell-inline')await expect(editor.locator('img[alt="OWNED LOGO"]')).toBeVisible();}
  const after=editor.locator('p').filter({hasText:/^AFTER/}),before=await editor.evaluate((root:any)=>root.editor.getJSON());await after.fill('AFTER EDIT');await expect(after).toHaveText('AFTER EDIT');await page.keyboard.press('ControlOrMeta+z');await expect.poll(()=>editor.evaluate((root:any)=>root.editor.getJSON())).toEqual(before);
 }
});
