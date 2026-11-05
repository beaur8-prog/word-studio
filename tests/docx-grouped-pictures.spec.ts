import {writeFile} from 'node:fs/promises';
import {test,expect} from '@playwright/test';
import JSZip from 'jszip';
import {groupedDrawing,GROUP_PNG,type GroupScenario} from './helpers/grouped-drawing';

for(const scenario of ['mixed','pictures','nested','repeated','fallback','paragraph'] as GroupScenario[]) test(`grouped ${scenario} pictures and editable stories survive import and two standard DOCX saves`,async({page})=>{
 const bytes=await groupedDrawing(scenario);await writeFile(test.info().outputPath('textbox-save-input-group.docx'),bytes);
 await page.goto('/');const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx,validateDocument}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');
  let file=new File([new Uint8Array(bytes)],'group.docx');const states:any[]=[],saves:number[][]=[];
  for(let i=0;i<3;i++){
   const opened=await importFile(file),content=parseDocumentHtml(opened.html!),doc={...opened.document,content};validateDocument(doc);
   const images:any[]=[],boxes:any[]=[];let text='';const walk=(n:any)=>{if(['image','inlineImage'].includes(n.type))images.push(n);if(n.type==='documentTextBox')boxes.push(n);text+=n.text||'';for(const c of n.content||[])walk(c);};walk(content);
   states.push({images,boxes,text,warnings:opened.warnings});if(i<2){const output=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(output);file=new File([new Uint8Array(output)],'again.docx');}
  }return{states,saves};
 },[...bytes]);
 for(const state of result.states){
  expect(state.images).toHaveLength(scenario==='repeated'?2:1);expect(state.boxes).toHaveLength(scenario==='pictures'?0:1);
  expect(state.text).toContain('BODY BEFORE');expect(state.text).toContain('BODY AFTER');
  for(const [i,image]of state.images.entries()){
   expect(image.attrs.src.split(',')[1]).toBe(GROUP_PNG);expect(image.attrs.alt).toBe(`GROUP PICTURE ${14+i}`);
   expect(Number(image.attrs.width)).toBeCloseTo(80,2);expect(Number(image.attrs.height)).toBeCloseTo(80,2);
   const g=JSON.parse(image.attrs.floating);expect(g.drawingGroup).toBeTruthy();expect(g.horizontalPosition.offset/9525).toBeCloseTo(scenario==='nested'?256:396+i*100,2);expect(g.verticalPosition.offset/9525).toBeCloseTo(scenario==='paragraph'?70:166,2);
   expect(g.drawingGroup.width).toBeCloseTo(480,2);expect(g.drawingGroup.height).toBeCloseTo(100,2);
  }
  if(state.boxes.length){expect(state.text.match(/EDITABLE GROUP STORY/g)).toHaveLength(1);const f=JSON.parse(state.boxes[0].attrs.frame);expect(f.x).toBeCloseTo(116,2);expect(f.y).toBeCloseTo(scenario==='paragraph'?60:156,2);expect(f.drawingGroup.id).toBe(JSON.parse(state.images[0].attrs.floating).drawingGroup.id);}
 }
 for(const [i,save]of result.saves.entries()){
  const output=Buffer.from(save);await writeFile(test.info().outputPath(`textbox-save-group-${i+1}.docx`),output);const zip=await JSZip.loadAsync(output),xml=await zip.file('word/document.xml')!.async('string');
  expect(xml.match(/<wp:anchor\b/g)).toHaveLength(1);expect(xml.match(/<wpg:wgp\b/g)).toHaveLength(1);expect(xml.match(/<pic:pic\b/g)).toHaveLength(scenario==='repeated'?2:1);expect(xml).not.toContain('data-ws-');
 }
});

test('mixed group draws its picture and editable box at the same positions in editor and print',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'mixed-group.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await groupedDrawing('mixed')});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('img[data-inline-image]')).toHaveCount(1);await expect(editor.locator('.document-textbox')).toHaveCount(1);
 const measure=(selector:string)=>page.locator(selector).evaluate(el=>{const sheet=el.closest('.document-stack,.print-page')!,s=sheet.getBoundingClientRect(),r=el.getBoundingClientRect();return{x:r.x-s.x,y:r.y-s.y,width:r.width,height:r.height};});
 await expect.poll(async()=>(await measure('.tiptap img[data-inline-image]')).x).toBeCloseTo(396,0);const image=await measure('.tiptap img[data-inline-image]'),box=await measure('.tiptap .document-textbox-content');expect(image.y).toBeCloseTo(166,0);expect(box.x).toBeCloseTo(116,0);expect(box.y).toBeCloseTo(156,0);
 await editor.locator('.document-textbox-content p').click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowRight':'End');await page.keyboard.insertText(' EDIT');await expect(editor.locator('.document-textbox-content')).toContainText('STORY EDIT');await page.keyboard.press('ControlOrMeta+z');await expect(editor.locator('.document-textbox-content')).toHaveText('EDITABLE GROUP STORY');await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor.locator('.document-textbox-content')).toContainText('STORY EDIT');
 await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();await expect(page.locator('.print-content img[data-inline-image]')).toHaveCount(1);const printedImage=await measure('.print-content img[data-inline-image]'),printedBox=await measure('.print-content .document-textbox-content');
 for(const key of ['x','y','width','height'] as const){expect(printedImage[key]).toBeCloseTo(image[key],0);expect(printedBox[key]).toBeCloseTo(box[key],0);}
});


test('body text respects the entire group wrap rectangle, including the space between children',async({page})=>{
 const zip=await JSZip.loadAsync(await groupedDrawing('mixed'));const xml=await zip.file('word/document.xml')!.async('string');
 const text=('Visible body text must flow around the whole picture and textbox group. ').repeat(65);zip.file('word/document.xml',xml.replace('BODY AFTER',text));const bytes=await zip.generateAsync({type:'nodebuffer'});
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'wrapped-group.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:bytes});const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('img[data-inline-image]')).toHaveCount(1);
 const inspect=(root:HTMLElement)=>{const img=root.querySelector<HTMLImageElement>('img[data-image-floating]')!,g=JSON.parse(img.dataset.imageFloating!).drawingGroup,box=img.closest('.document-stack,.print-page')!.getBoundingClientRect(),rect={x:box.x+g.x,y:box.y+g.y,w:g.width,h:g.height};let overlaps=0,gap=0;const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);while(walker.nextNode()){const n=walker.currentNode;if(n.parentElement?.closest('.document-textbox,[data-flow-widget]'))continue;for(let i=0;i<(n.textContent||'').length;i++){if(/\s/.test(n.textContent![i]))continue;const range=document.createRange();range.setStart(n,i);range.setEnd(n,i+1);const r=range.getBoundingClientRect();if(r.right>rect.x+.5 && r.left<rect.x+rect.w-.5 && r.bottom>rect.y+.5 && r.top<rect.y+rect.h-.5){overlaps++;if(r.left>rect.x+240 && r.right<rect.x+280)gap++;}}}return{overlaps,gap,text:root.textContent};};
 await expect.poll(async()=>(await editor.evaluate(inspect)).overlaps).toBe(0);expect((await editor.evaluate(inspect)).text).toContain(text);
 await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();await expect(page.locator('.print-page').first().locator('img[data-inline-image]')).toHaveCount(1);const result=await page.locator('.print-page').first().evaluate(inspect);expect(result.overlaps).toBe(0);
});

test('a mixed group in a header keeps its own picture relationship and story through two saves',async({page})=>{
 const zip=await JSZip.loadAsync(await groupedDrawing('mixed'));let xml=await zip.file('word/document.xml')!.async('string');const drawing=xml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/)![0];
 const root=xml.slice(0,xml.indexOf('<w:body>')).replace('<w:document','<w:hdr');zip.file('word/header1.xml',root+'<w:p><w:r>'+drawing+'</w:r></w:p></w:hdr>');
 xml=xml.replace(drawing,'').replace('<w:sectPr>','<w:sectPr><w:headerReference w:type="default" r:id="header"/>');zip.file('word/document.xml',xml);
 zip.file('word/_rels/header1.xml.rels',(await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('Id="groupPicture"','Id="groupPicture"'));
 zip.file('word/_rels/document.xml.rels',(await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>','<Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
 zip.file('[Content_Types].xml',(await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>','<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
 const bytes=await zip.generateAsync({type:'nodebuffer'});await writeFile(test.info().outputPath('textbox-save-input-header-group.docx'),bytes);await page.goto('/');
 const result=await page.evaluate(async bytes=>{const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let file=new File([new Uint8Array(bytes)],'header.docx');const states:any[]=[],saves:number[][]=[];for(let i=0;i<3;i++){const opened=await importFile(file);const html=opened.document.page.headerHtml||'';const t=document.createElement('template');t.innerHTML=html;states.push({html,images:[...t.content.querySelectorAll('img')].map(e=>e.getAttribute('src')),boxes:t.content.querySelectorAll('[data-document-textbox]').length,text:t.content.textContent});if(i<2){const save=[...new Uint8Array(await(await exportDocx({...opened.document,content:parseDocumentHtml(opened.html!)})).arrayBuffer())];saves.push(save);file=new File([new Uint8Array(save)],'again.docx');}}return{states,saves};},[...bytes]);
 for(const s of result.states){expect(s.images).toHaveLength(1);expect(s.images[0]!.split(',')[1]).toBe(GROUP_PNG);expect(s.boxes).toBe(1);expect(s.text).toContain('EDITABLE GROUP STORY');}
 for(const[i,save]of result.saves.entries())await writeFile(test.info().outputPath(`textbox-save-header-group-${i+1}.docx`),Buffer.from(save));
});
