import {writeFile} from 'node:fs/promises';
import {test,expect} from '@playwright/test';
import JSZip from 'jszip';
import {groupedDrawing,headerGroupedDrawing,boxGroupWithNestedDrawing,tableGroupedDrawing,GROUP_PNG,TABLE_GROUP_INNER_PNG,type GroupScenario} from './helpers/grouped-drawing';

for(const scenario of ['mixed','pictures','nested','repeated','fallback','paragraph','inline','inline-box'] as GroupScenario[]) test(`grouped ${scenario} pictures and editable stories survive import and two standard DOCX saves`,async({page})=>{
 const bytes=await groupedDrawing(scenario);await writeFile(test.info().outputPath('textbox-save-input-group.docx'),bytes);
 await page.goto('/');const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx,validateDocument}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');
  let file=new File([new Uint8Array(bytes)],'group.docx');const states:any[]=[],saves:number[][]=[];
  for(let i=0;i<3;i++){
   const opened=await importFile(file),content=parseDocumentHtml(opened.html!),doc={...opened.document,content};validateDocument(doc);
   const images:any[]=[],boxes:any[]=[],anchors:any[]=[];let text='';const walk=(n:any)=>{if(['image','inlineImage'].includes(n.type))images.push(n);if(n.type==='documentTextBox')boxes.push(n);if(n.type==='documentDrawingAnchor')anchors.push(n);text+=n.text||'';for(const c of n.content||[])walk(c);};walk(content);
   states.push({images,boxes,anchors,text,warnings:opened.warnings});if(i<2){const output=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(output);file=new File([new Uint8Array(output)],'again.docx');}
  }return{states,saves};
 },[...bytes]);
 for(const state of result.states){
  expect(state.images).toHaveLength(scenario==='inline-box'?0:scenario==='repeated'?2:1);expect(state.boxes).toHaveLength(scenario==='pictures'?0:1);
  expect(state.text).toContain('BODY BEFORE');expect(state.text).toContain('BODY AFTER');expect(state.anchors).toHaveLength(scenario.startsWith('inline')?1:0);
  for(const [i,image]of state.images.entries()){
   expect(image.attrs.src.split(',')[1]).toBe(GROUP_PNG);expect(image.attrs.alt).toBe(`GROUP PICTURE ${14+i}`);
   expect(Number(image.attrs.width)).toBeCloseTo(80,2);expect(Number(image.attrs.height)).toBeCloseTo(80,2);
   const g=JSON.parse(image.attrs.floating);expect(g.drawingGroup).toBeTruthy();expect(g.horizontalPosition.offset/9525).toBeCloseTo(scenario.startsWith('inline')?280:scenario==='nested'?256:396+i*100,2);expect(g.verticalPosition.offset/9525).toBeCloseTo(scenario.startsWith('inline')?10:scenario==='paragraph'?70:166,2);
   expect(g.drawingGroup.width).toBeCloseTo(480,2);expect(g.drawingGroup.height).toBeCloseTo(100,2);
  }
  if(state.boxes.length){expect(state.text.match(/EDITABLE GROUP STORY/g)).toHaveLength(1);const f=JSON.parse(state.boxes[0].attrs.frame);expect(f.x).toBeCloseTo(scenario.startsWith('inline')?0:116,2);expect(f.y).toBeCloseTo(scenario.startsWith('inline')?0:scenario==='paragraph'?60:156,2);expect(f.drawingGroup.id).toBe(state.images.length?JSON.parse(state.images[0].attrs.floating).drawingGroup.id:JSON.parse(state.anchors[0].attrs.frame).id);}
 }
 for(const [i,save]of result.saves.entries()){
  const output=Buffer.from(save);await writeFile(test.info().outputPath(`textbox-save-group-${i+1}.docx`),output);const zip=await JSZip.loadAsync(output),xml=await zip.file('word/document.xml')!.async('string');
  expect(xml.match(scenario.startsWith('inline')?/<wp:inline\b/g:/<wp:anchor\b/g)).toHaveLength(1);expect(xml.match(/<wpg:wgp\b/g)).toHaveLength(1);expect(xml.match(/<pic:pic\b/g)||[]).toHaveLength(scenario==='inline-box'?0:scenario==='repeated'?2:1);expect(xml).not.toContain('data-ws-');
 }
});

test('mixed group draws its picture and editable box at the same positions in editor and print',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'mixed-group.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await groupedDrawing('mixed')});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('img[data-inline-image]')).toHaveCount(1);await expect(editor.locator('.document-textbox')).toHaveCount(1);
 const measure=(selector:string)=>page.locator(selector).evaluate(el=>{const sheet=el.closest('.document-stack,.print-page')!,s=sheet.getBoundingClientRect(),r=el.getBoundingClientRect();return{x:r.x-s.x,y:r.y-s.y,width:r.width,height:r.height};});
 await expect.poll(async()=>(await measure('.tiptap img[data-inline-image]')).x).toBeCloseTo(396,0);const image=await measure('.tiptap img[data-inline-image]'),box=await measure('.tiptap .document-textbox-content');expect(image.y).toBeCloseTo(166,0);expect(box.x).toBeCloseTo(116,0);expect(box.y).toBeCloseTo(156,0);
 await editor.locator('.document-textbox-content p').click();
 // Append at the paragraph end, independent of font-driven line wrapping.
 await editor.locator('.document-textbox-content p').evaluate(el=>{const r=document.createRange();r.selectNodeContents(el);r.collapse(false);const s=getSelection()!;s.removeAllRanges();s.addRange(r);document.dispatchEvent(new Event('selectionchange'));});
 await page.keyboard.press('ArrowLeft');await page.keyboard.press('ArrowRight');await page.keyboard.insertText(' EDIT');await expect(editor.locator('.document-textbox-content')).toContainText('STORY EDIT');await page.keyboard.press('ControlOrMeta+z');await expect(editor.locator('.document-textbox-content')).toHaveText('EDITABLE GROUP STORY');await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor.locator('.document-textbox-content')).toContainText('STORY EDIT');
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
 const bytes=await headerGroupedDrawing();await writeFile(test.info().outputPath('textbox-save-input-header-group.docx'),bytes);await page.goto('/');
 const result=await page.evaluate(async bytes=>{const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let file=new File([new Uint8Array(bytes)],'header.docx');const states:any[]=[],saves:number[][]=[];for(let i=0;i<3;i++){const opened=await importFile(file);const html=opened.document.page.headerHtml||'';const t=document.createElement('template');t.innerHTML=html;states.push({html,images:[...t.content.querySelectorAll('img')].map(e=>e.getAttribute('src')),boxes:t.content.querySelectorAll('[data-document-textbox]').length,text:t.content.textContent});if(i<2){const save=[...new Uint8Array(await(await exportDocx({...opened.document,content:parseDocumentHtml(opened.html!)})).arrayBuffer())];saves.push(save);file=new File([new Uint8Array(save)],'again.docx');}}return{states,saves};},[...bytes]);
 for(const s of result.states){expect(s.images).toHaveLength(1);expect(s.images[0]!.split(',')[1]).toBe(GROUP_PNG);expect(s.boxes).toBe(1);expect(s.text).toContain('EDITABLE GROUP STORY');}
 for(const[i,save]of result.saves.entries())await writeFile(test.info().outputPath(`textbox-save-header-group-${i+1}.docx`),Buffer.from(save));
});


test('header textboxes paint their source background and white text on each editor and print paper',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'visible-header-group.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await headerGroupedDrawing(true)});
 const editorBoxes=page.locator('.paper-header [data-textbox-content]');await expect(editorBoxes).toHaveCount(2);
 for(const box of await editorBoxes.all())await expect(box).toHaveCSS('background-color','rgb(32, 80, 128)');
 const measure=(el:Element)=>{const p=el.closest('.paper,.print-page')!.getBoundingClientRect(),r=el.getBoundingClientRect();return{x:r.x-p.x,y:r.y-p.y,width:r.width,height:r.height,text:el.textContent,color:getComputedStyle(el.querySelector('span')!).color};};
 const before=[];for(const box of await editorBoxes.all()){const b=await box.evaluate(measure);expect(b.x).toBeCloseTo(116,0);expect(b.y).toBeCloseTo(156,0);expect(b.color).toBe('rgb(255, 255, 255)');before.push(b);}
 await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();const printBoxes=page.locator('.print-header [data-textbox-content]');await expect(printBoxes).toHaveCount(2);
 for(const[i,box]of (await printBoxes.all()).entries()){await expect(box).toHaveCSS('background-color','rgb(32, 80, 128)');const b=await box.evaluate(measure);for(const key of ['x','y','width','height']as const)expect(b[key]).toBeCloseTo(before[i][key],0);expect(b.text).toBe(before[i].text);}
});

test('an otherwise empty grouped drawing anchor keeps its real paragraph line in print',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'group-anchor-line.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await groupedDrawing('mixed')});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('img[data-inline-image]')).toHaveCount(1);
 const measure=(el:Element)=>{const s=el.closest('.document-stack,.print-page')!.getBoundingClientRect(),r=el.getBoundingClientRect();return{y:r.top-s.top,height:r.height};};
 const before=await editor.locator('p').filter({hasText:'BODY AFTER'}).evaluate(measure);await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();const after=await page.locator('.print-content p').filter({hasText:'BODY AFTER'}).evaluate(measure);expect(after.y).toBeCloseTo(before.y,0);
});


test('a box-only outer group retains its wrap when a separate mixed drawing is inside a textbox story',async({page})=>{
 const bytes=await boxGroupWithNestedDrawing();await writeFile(test.info().outputPath('textbox-save-input-nested-story-group.docx'),bytes);await page.goto('/');
 const result=await page.evaluate(async bytes=>{const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let file=new File([new Uint8Array(bytes)],'nested-story-group.docx');const states:any[]=[],saves:number[][]=[];for(let i=0;i<3;i++){const d=await importFile(file),content=parseDocumentHtml(d.html!),boxes=content.content!.filter(n=>n.type==='documentTextBox');let pictures=0;const walk=(n:any)=>{if(n.type==='inlineImage')pictures++;for(const c of n.content||[])walk(c);};walk(content);states.push({frames:boxes.map(n=>JSON.parse(n.attrs!.frame)),pictures});if(i<2){const save=[...new Uint8Array(await(await exportDocx({...d.document,content})).arrayBuffer())];saves.push(save);file=new File([new Uint8Array(save)],'again.docx');}}return{states,saves};},[...bytes]);
 for(const s of result.states){expect(s.pictures).toBe(1);expect(s.frames).toHaveLength(2);for(const f of s.frames){expect(f.wrap).toBe('square');expect(f.drawingGroup).toBeUndefined();}expect(s.frames[0].group).toBeTruthy();expect(s.frames[0].group).toBe(s.frames[1].group);}
 for(const[i,save]of result.saves.entries())await writeFile(test.info().outputPath(`textbox-save-nested-story-group-${i+1}.docx`),Buffer.from(save));
});

test('an inline picture and textbox group reserves its actual inline box without covering surrounding text',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'inline-group.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await groupedDrawing('inline')});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('[data-drawing-group-anchor]')).toHaveCount(1);await expect(editor.locator('img[data-inline-image]')).toHaveCount(1);
 const geometry=(root:Element)=>{const a=root.querySelector('[data-drawing-group-anchor]')!.getBoundingClientRect(),i=root.querySelector('img[data-inline-image]')!.getBoundingClientRect(),b=root.querySelector('.document-textbox-content')!.getBoundingClientRect();return{width:a.width,height:a.height,image:[i.x-a.x,i.y-a.y,i.width,i.height],box:[b.x-a.x,b.y-a.y,b.width,b.height],text:root.textContent};};
 await expect.poll(async()=>(await editor.evaluate(geometry)).image[0]).toBeCloseTo(280,0);const before=await editor.evaluate(geometry);expect(before.width).toBeCloseTo(480,0);expect(before.height).toBeCloseTo(100,0);expect(before.image).toEqual([280,10,80,80]);expect(before.box).toEqual([0,0,240,80]);expect(before.text).toContain('INLINE BEFORE');expect(before.text).toContain('INLINE AFTER');
 await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();const after=await page.locator('.print-content').evaluate(geometry);expect(after.image).toEqual(before.image);expect(after.box).toEqual(before.box);expect(after.width).toBe(before.width);expect(after.height).toBe(before.height);
});

for(const scenario of ['mixed','inline'] as const)test(`grouped ${scenario} merged table and two picture roles survive two DOCX saves`,async({page})=>{
 const bytes=await tableGroupedDrawing(scenario);await writeFile(test.info().outputPath(`textbox-save-input-table-group-${scenario}.docx`),bytes);await page.goto('/');
 const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let file=new File([new Uint8Array(bytes)],'table-group.docx');const states:any[]=[],saves:number[][]=[];
  for(let i=0;i<3;i++){
   const opened=await importFile(file),doc={...opened.document,content:parseDocumentHtml(opened.html!)};const nodes:any[]=[];const walk=(n:any)=>{nodes.push(n);for(const c of n.content||[])walk(c);};walk(doc.content);
   const table=nodes.find(n=>n.type==='table'),boxes=nodes.filter(n=>n.type==='documentTextBox');
   states.push({text:nodes.map(n=>n.text||'').join(''),boxes:boxes.length,rows:table?.content.map((r:any)=>r.content.map((c:any)=>({colspan:c.attrs.colspan,rowspan:c.attrs.rowspan,text:c.content.map((p:any)=>(p.content||[]).map((n:any)=>n.text||'').join('')).join('')}))),pictures:nodes.filter(n=>['image','inlineImage'].includes(n.type)).map(n=>({alt:n.attrs.alt,png:n.attrs.src.split(',')[1],width:Number(n.attrs.width),height:Number(n.attrs.height),group:!!JSON.parse(n.attrs.floating||'null')?.drawingGroup}))});
   if(i<2){const output=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(output);file=new File([new Uint8Array(output)],'again.docx');}
  }return{states,saves};
 },[...bytes]);
 for(const s of result.states){
  expect(s.boxes).toBe(1);for(const marker of ['BODY BEFORE','BODY AFTER','MERGED HEADER','ROWSPAN CELL','UPPER CELL','TAIL CELL','BOX TAIL'])expect(s.text.split(marker)).toHaveLength(2);
  expect(s.rows).toEqual([[{colspan:2,rowspan:1,text:'MERGED HEADER'}],[{colspan:1,rowspan:2,text:'ROWSPAN CELL'},{colspan:1,rowspan:1,text:'UPPER CELL'}],[{colspan:1,rowspan:1,text:'TAIL CELL'}]]);
  expect(s.pictures).toHaveLength(2);expect(s.pictures.find(p=>p.alt==='GROUP PICTURE 14')).toMatchObject({png:GROUP_PNG,width:80,height:80,group:true});expect(s.pictures.find(p=>p.alt==='INNER TABLE PICTURE')).toMatchObject({png:TABLE_GROUP_INNER_PNG,width:12,height:12,group:false});
 }
 // Internal editable stories need not be traversed in XML reference order.
 // Check the saved standard group itself, including the cell's own picture.
 for(const [i,data]of [bytes,...result.saves.map(s=>Buffer.from(s))].entries()){
  const zip=await JSZip.loadAsync(data),xml=await zip.file('word/document.xml')!.async('string'),rels=await zip.file('word/_rels/document.xml.rels')!.async('string');
  const paths=new Map([...rels.matchAll(/<Relationship\b[^>]*>/g)].map(m=>[m[0].match(/\bId="([^"]+)"/)![1],m[0].match(/\bTarget="([^"]+)"/)![1]])),pictures=[...xml.matchAll(/<pic:pic\b[\s\S]*?<\/pic:pic>/g)];const ordered=[];
  for(const p of pictures){const id=p[0].match(/\br:embed="([^"]+)"/)![1],target=paths.get(id)!;ordered.push((await zip.file(target.startsWith('/')?target.slice(1):'word/'+target)!.async('base64')));}
  expect(ordered).toEqual([TABLE_GROUP_INNER_PNG,GROUP_PNG]);expect(xml.match(/<wpg:wgp\b/g)).toHaveLength(1);expect(xml.match(/<w:tbl\b/g)).toHaveLength(1);expect(xml).toMatch(/<w:gridSpan\b[^>]*w:val="2"/);expect(xml).toMatch(/<w:vMerge\b[^>]*w:val="restart"/);
  if(i)await writeFile(test.info().outputPath(`textbox-save-table-group-${scenario}-${i}.docx`),Buffer.from(data));
 }
});

for(const scenario of ['mixed','inline']as const)test(`grouped ${scenario} table cell stays visible and editable in editor and print`,async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'table-group.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await tableGroupedDrawing(scenario)});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true}),box=editor.locator('.document-textbox-content');await expect(box.locator('table')).toHaveCount(1);await expect(box.locator('img[data-inline-image]')).toHaveCount(1);await expect(editor.locator('img[data-inline-image]')).toHaveCount(2);await expect(box).toHaveCSS('background-color','rgb(32, 80, 128)');
 const inspect=(root:Element)=>{const frame=root.getBoundingClientRect(),table=root.querySelector('table')!.getBoundingClientRect(),cells=[...root.querySelectorAll('td,th')];let outside=0;const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);while(walker.nextNode()){const n=walker.currentNode;for(let i=0;i<(n.textContent||'').length;i++){if(/\s/.test(n.textContent![i]))continue;const range=document.createRange();range.setStart(n,i);range.setEnd(n,i+1);const r=range.getBoundingClientRect();if(r.left<frame.left-.5||r.right>frame.right+.5||r.top<frame.top-.5||r.bottom>frame.bottom+.5)outside++;}}return{outside,table:[table.x-frame.x,table.y-frame.y,table.width,table.height],cells:cells.map(e=>({text:e.textContent,colspan:e.getAttribute('colspan')||'1',rowspan:e.getAttribute('rowspan')||'1'}))};};
 await expect.poll(async()=>(await box.evaluate(inspect)).outside).toBe(0);const before=await editor.evaluate((el:any)=>el.editor.getJSON()),p=box.locator('tr').nth(2).locator('td p').first();await p.click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowRight':'End');
 await expect.poll(()=>editor.evaluate((el:any)=>{const s=el.editor.state.selection;return s.empty?el.editor.state.doc.textBetween(s.$from.start(),s.from,'','\ufffc'):null;})).toBe('TAIL CELL');await page.keyboard.insertText(' EDIT');await expect(p).toHaveText('TAIL CELL EDIT');await page.keyboard.press('ControlOrMeta+z');expect(await editor.evaluate((el:any)=>el.editor.getJSON())).toEqual(before);await page.keyboard.press('ControlOrMeta+Shift+z');await expect(p).toHaveText('TAIL CELL EDIT');
 const layout=await box.evaluate(inspect);await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();const printed=page.locator('.print-content [data-textbox-content]');await expect(printed).toHaveCount(1);await expect(printed.locator('img[data-inline-image]')).toHaveCount(1);await expect(page.locator('.print-content img[data-inline-image]')).toHaveCount(2);await expect(printed).toHaveCSS('background-color','rgb(32, 80, 128)');const after=await printed.evaluate(inspect);expect(after.outside).toBe(0);expect(after.cells).toEqual(layout.cells);for(const[i,v]of after.table.entries())expect(v).toBeCloseTo(layout.table[i],0);
});
