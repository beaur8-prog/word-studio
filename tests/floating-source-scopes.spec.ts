import { test, expect, type Page } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
const editor=(page:Page)=>page.getByRole('textbox',{name:'문서 편집 영역',exact:true});
async function geometry(page:Page){return page.evaluate(()=>{
 const root=document.querySelector<HTMLElement>('.tiptap[contenteditable=true]')!,frame=[...root.querySelectorAll<HTMLElement>('.document-section-content')].at(-1)!,origin=frame.getBoundingClientRect();
 const rect=(el:Element)=>{const b=el.getBoundingClientRect();return[b.x-origin.x,b.y-origin.y,b.width,b.height].map(x=>Math.round(x*64)/64)};
 return{pages:Number(frame.dataset.pages),flags:frame.querySelectorAll('[data-floating-start-page]').length,before:[...frame.querySelectorAll('p')].find(p=>p.textContent?.startsWith('BEFORE '))!.getBoundingClientRect().height,
 images:[...frame.querySelectorAll('img')].map(rect),paragraphs:[...frame.querySelectorAll('p')].map(rect)};
})}
for(const change of ['text','picture-index','remove-drawing']as const)test(`independent framed drawing layout survives ${change} editing and undo`,async({page},info)=>{
 test.setTimeout(120000);const errors:string[]=[];page.on('pageerror',e=>errors.push(String(e)));await page.goto('/');await expect(editor(page)).toBeVisible();
 const doc=await page.evaluate(async()=>{
  const{parseDocumentHtml}=await import('/src/extensions.ts'),{defaultPage}=await import('/src/types.ts');
  const c=document.createElement('canvas');c.width=c.height=4;c.getContext('2d')!.fillRect(0,0,4,4);const png=c.toDataURL();
  const p=(t:string)=>`<p style="font-family:Arial;font-size:12pt;line-height:20px;margin:0">${t}</p>`;
  const image=(alt:string,y:number,h:number)=>{const g={horizontalPosition:{relative:'page',offset:96*9525},verticalPosition:{relative:'page',offset:y*9525},wrap:{type:1,side:'bothSides'},layoutInCell:true,allowOverlap:true};return `<img data-inline-image="true" src="${png}" width="600" height="${h}" alt="${alt}" data-image-floating="${JSON.stringify(g).replaceAll('"','&quot;')}">`};
  const body=parseDocumentHtml(p(image('EARLY',96,120)+'EARLY ANCHOR')+Array.from({length:8},(_,i)=>p('LEADING '+i)).join('')+p('BEFORE '+'Ordinary text words for a later anchored picture. '.repeat(18))+`<p style="font-family:Arial;font-size:12pt;line-height:.75pt;margin:0" data-word-line-rule="exact">${image('LATER',174,650)}</p>`+p('AFTER '+'Following text must remain visible and in order. '.repeat(8)));
  const cover={...defaultPage,margin:30,pageNumbers:false},next={...defaultPage,pageNumbers:false,breakType:'nextPage' as const};
  return{format:'word-studio',version:1,id:'independent-floats',title:'Independent floats',comments:[],updatedAt:new Date().toISOString(),page:cover,content:{type:'doc',content:[{type:'sectionContent',attrs:{page:JSON.stringify(cover)},content:parseDocumentHtml(p('COVER TITLE')).content},{type:'sectionBreak',attrs:{page:JSON.stringify(next)}},{type:'sectionContent',attrs:{page:JSON.stringify(next)},content:body.content}]}};
 });
 await page.locator('input[type=file]').first().setInputFiles({name:'independent.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});await expect(page.locator('.busy-indicator')).toBeHidden();await page.evaluate(()=>document.fonts.ready);
 await expect.poll(async()=>{const g=await geometry(page);return[g.pages,Math.round(g.before),g.flags]}).toEqual([2,220,1]);await page.waitForTimeout(500);const original=await geometry(page),model=await editor(page).evaluate((el:any)=>el.editor.getJSON());
 if(change==='text'){
  await editor(page).evaluate((el:any)=>{let pos=0;el.editor.state.doc.descendants((n:any,p:number)=>{if(n.isTextblock&&n.textContent==='COVER TITLE')pos=p+1+n.content.size});el.editor.commands.setTextSelection(pos);el.editor.commands.focus()});await expect(editor(page)).toBeFocused();await page.keyboard.insertText(' EDIT');await expect(editor(page)).toContainText('COVER TITLE EDIT');
 }else await editor(page).evaluate((el:any,change)=>{
  const e=el.editor;if(change==='picture-index'){let end=0;e.state.doc.descendants((n:any,p:number)=>{if(n.isTextblock&&n.textContent==='COVER TITLE')end=p+1+n.content.size});let source:any;e.state.doc.descendants((n:any)=>{if(!source&&n.type.name==='inlineImage')source=n});const image=e.schema.nodes.inlineImage.create({...source.attrs,width:16,height:16,alt:'COVER PICTURE',floating:JSON.stringify({...JSON.parse(source.attrs.floating),wrap:{type:0}})});e.view.dispatch(e.state.tr.insert(end,image));
  }else{let pos=-1,size=0;e.state.doc.descendants((n:any,p:number)=>{if(n.type.name==='inlineImage'&&n.attrs.alt==='LATER'){pos=p;size=n.nodeSize}});e.view.dispatch(e.state.tr.delete(pos,pos+size))}
 },change);
 if(change==='remove-drawing'){await expect.poll(async()=> (await geometry(page)).flags).toBe(0);await expect.poll(async()=> (await geometry(page)).pages).toBe(1)}else{
  const samples=[];for(let i=0;i<20;i++){samples.push(await geometry(page));await page.waitForTimeout(100)}for(const s of samples){expect(s.pages).toBe(original.pages);expect(s.flags).toBe(original.flags);expect(Math.abs(s.before-original.before)).toBeLessThan(.05);expect(s.images).toEqual(original.images)}await writeFile(info.outputPath('unrelated-frame-observations.json'),JSON.stringify(samples));
 }
 await editor(page).evaluate((el:any)=>el.editor.commands.focus());await expect(editor(page)).toBeFocused();await page.keyboard.press('ControlOrMeta+z');await expect.poll(()=>editor(page).evaluate((el:any)=>el.editor.getJSON())).toEqual(model);
 await expect.poll(async()=>{const g=await geometry(page);return[g.pages,Math.round(g.before),g.flags]}).toEqual([original.pages,Math.round(original.before),original.flags]);await page.waitForTimeout(500);const undone=await geometry(page);
 expect(undone.images).toEqual(original.images);expect(undone.paragraphs).toEqual(original.paragraphs);expect(errors).toEqual([]);
});
