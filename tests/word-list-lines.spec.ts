import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
import {Document,Packer,Paragraph,TextRun,LevelFormat} from 'docx';

// Synthetic counterexamples, separate from the frozen Word-authored originals.
for(const spacing of [false,true])test(`imported list shells add no marker strut or ${spacing?'authored':'zero'} paragraph space`,async({page})=>{
 await page.goto('/');
 const result=await page.evaluate(async(spacing)=>{
  const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js');
  const{editorExtensions,parseDocumentHtml}=await import('/src/extensions.ts');
  const{createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
  const p=(text:string,size=10,before=0,after=0)=>`<p data-word-line-rule="auto" data-word-line="1.15" style="font-family:Arial;font-size:${size}pt;line-height:1.15;margin:${before}px 0 ${after}px">${text}</p>`;
  const li=(text:string,level=0,size=10,before=0,after=0)=>`<li data-list-id="Counter" data-list-level="${level}" data-list-format="decimal" data-list-pattern="%${level+1}.">${p(text,size,before,after)}</li>`;
  const values=spacing?[4,8,3,6]:[0,0,0,0];
  const rows=[p('FIRST',10,values[0],values[1]),p('SECOND',10,values[2],values[3]),p('NESTED',14,values[0],values[1]),p('AFTER',10,values[2],values[3])];
  const html='<ol>'+li('FIRST',0,10,values[0],values[1])+li('SECOND',0,10,values[2],values[3]).replace('</li>','<ol>'+li('NESTED',1,14,values[0],values[1])+'</ol></li>')+li('AFTER',0,10,values[2],values[3])+'</ol>';
  const snapshot=(root:HTMLElement)=>{
   const ps=[...root.querySelectorAll('p')].filter(p=>p.textContent),top=ps[0].getBoundingClientRect().top;
   return{paragraphs:ps.map(p=>{const s=getComputedStyle(p);return{text:p.textContent,top:p.getBoundingClientRect().top-top,height:p.getBoundingClientRect().height,before:s.marginTop,after:s.marginBottom}}),lists:[...root.querySelectorAll('ol')].map(l=>getComputedStyle(l).marginBottom),items:[...root.querySelectorAll('li')].map(li=>({font:getComputedStyle(li).fontSize,pFont:getComputedStyle(li.querySelector(':scope > p')!).fontSize,leading:getComputedStyle(li).lineHeight,pLeading:getComputedStyle(li.querySelector(':scope > p')!).lineHeight})),labels:[...root.querySelectorAll('li')].map(li=>(li as HTMLElement).dataset.listLabel)};
  };
  const mount=document.createElement('div');document.body.append(mount);
  const editor=new Editor({element:mount,extensions:editorExtensions,content:parseDocumentHtml(html)});
  try{
   await document.fonts.ready;await new Promise(r=>setTimeout(r,120));
   const editing=snapshot(editor.view.dom),model=editor.getJSON(),serialized=editor.getHTML();
   const print=await createPrintLayout(serialized,defaultPage),flat=await createPrintLayout(rows.join(''),defaultPage);
   try{return{editing,printing:snapshot(print.element),flat:snapshot(flat.element),model,serialized}}
   finally{print.dispose();flat.dispose()}
  }finally{editor.destroy();mount.remove()}
 },spacing);
 for(const stage of [result.editing,result.printing]){
  expect(stage.paragraphs.map(p=>p.text)).toEqual(['FIRST','SECOND','NESTED','AFTER']);
  for(let i=0;i<stage.paragraphs.length;i++){
   expect(stage.paragraphs[i].top,JSON.stringify(stage)).toBeCloseTo(result.flat.paragraphs[i].top,1);
   expect(stage.paragraphs[i].before).toBe(result.flat.paragraphs[i].before);
   expect(stage.paragraphs[i].after).toBe(result.flat.paragraphs[i].after);
  }
  expect(stage.lists).toEqual(['0px','0px']);
  for(const item of stage.items){expect(item.font).toBe(item.pFont);expect(item.leading).toBe(item.pLeading)}
  expect(stage.labels).toEqual(['1.','2.','1.','3.']);
 }
 expect(JSON.stringify(result.model)).not.toContain('21.25px');
 expect(result.serialized).not.toMatch(/<li[^>]*style=/);
});

test('authored lists retain their normal marker typography and default margins',async({page})=>{
 await page.goto('/');
 const actual=await page.evaluate(async()=>{
  const{createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
  const layout=await createPrintLayout('<ol><li><p>AUTHORED</p><ul><li><p>NESTED</p></li></ul></li></ol>',defaultPage);
  try{return[...layout.element.querySelectorAll('ol,ul')].map(l=>({margin:getComputedStyle(l).marginBottom,item:getComputedStyle(l.querySelector('li')!).fontSize,root:getComputedStyle(l.closest('.tiptap')!).fontSize}))}finally{layout.dispose()}
 });
 expect(actual.map(l=>l.margin)).toEqual(['14px','4px']);for(const l of actual)expect(l.item).toBe(l.root);
});

test('imported marker height follows edits, undo and redo without persisting view pixels',async({page})=>{
 await page.goto('/');
 const result=await page.evaluate(async()=>{
  const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{editorExtensions}=await import('/src/extensions.ts');
  const{closeHistory}=await import('/node_modules/.vite/deps/@tiptap_pm_history.js');
  const mount=document.createElement('div');document.body.append(mount);
  const p=(text:string)=>({type:'paragraph',attrs:{paragraphFont:'LateListFont,Arial',paragraphSize:'10pt',lineSpacing:'1.15',wordLineRule:'auto',spaceBefore:'0px',spaceAfter:'0px'},content:[{type:'text',text}]});
  const li=(text:string)=>({type:'listItem',attrs:{wordListId:'Edit',wordListLevel:0,wordListPattern:'%1.',wordListFormat:'decimal'},content:[p(text)]});
  const editor=new Editor({element:mount,extensions:editorExtensions,content:{type:'doc',content:[{type:'orderedList',content:[li('FIRST'),li('SECOND')]}]}});
  const settle=()=>new Promise(r=>setTimeout(r,100)),snap=()=>{const li=editor.view.dom.querySelector('li')!,p=li.querySelector('p')!;return{height:li.getBoundingClientRect().height,pHeight:p.getBoundingClientRect().height,font:getComputedStyle(li).fontSize,pFont:getComputedStyle(p).fontSize,leading:getComputedStyle(li).lineHeight,pLeading:getComputedStyle(p).lineHeight}};
  try{
   await settle();const before=snap(),original=editor.getJSON();editor.view.dispatch(closeHistory(editor.state.tr));
   editor.commands.setTextSelection(3);editor.commands.updateAttributes('paragraph',{paragraphSize:'24pt'});await settle();const enlarged=snap();
   editor.commands.undo();await settle();const undone=snap(),undoModel=editor.getJSON();editor.commands.redo();await settle();const redone=snap();
   editor.commands.undo();await settle();const face=new FontFace('LateListFont','url(/src/assets/fonts/NotoSansMono.ttf)',{ascentOverride:'170%',descentOverride:'40%',lineGapOverride:'10%'});document.fonts.add(face);await face.load();await document.fonts.ready;await settle();const late=snap();return{before,enlarged,undone,redone,late,original,undoModel,html:editor.getHTML()};
  }finally{editor.destroy();mount.remove()}
 });
 for(const stage of [result.before,result.enlarged,result.undone,result.redone,result.late]){expect(stage.height).toBeCloseTo(stage.pHeight,1);expect(stage.font).toBe(stage.pFont);expect(stage.leading).toBe(stage.pLeading)}
 expect(result.enlarged.height).toBeGreaterThan(result.before.height*1.8);expect(result.late.height).toBeGreaterThan(result.before.height*1.3);expect(result.undone).toEqual(result.before);expect(result.redone).toEqual(result.enlarged);expect(result.undoModel).toEqual(result.original);expect(result.html).not.toMatch(/<li[^>]*style=/);
});

test('standard DOCX list spacing and counters survive editing and two saves',async({page},testInfo)=>{
 const input=await Packer.toBuffer(new Document({styles:{default:{document:{run:{font:'Arial',size:20}}},paragraphStyles:[{id:'ListBody',name:'ListBody',paragraph:{contextualSpacing:true,spacing:{before:120,after:160,line:276}}}]},numbering:{config:[{reference:'List',levels:[0,1].map(level=>({level,start:level?1:4,format:LevelFormat.DECIMAL,text:level?'%1.%2.':'%1.',style:{paragraph:{indent:{left:720+level*360,hanging:360}}}}))}]},sections:[{children:[new Paragraph({text:'FIRST',style:'ListBody',numbering:{reference:'List',level:0}}),new Paragraph({children:[new TextRun('SECOND')],style:'ListBody',numbering:{reference:'List',level:0}}),new Paragraph({text:'NESTED',style:'ListBody',numbering:{reference:'List',level:1}}),new Paragraph({text:'AFTER',style:'ListBody',numbering:{reference:'List',level:0}})]}]}));
 await page.goto('/');
 await page.locator('input[type=file]').first().setInputFiles({name:'list-lines.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:input});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText('FIRST');
 await editor.locator('li > p').filter({hasText:'SECOND'}).click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowRight':'End');await page.keyboard.type('-EDIT');await expect(editor).toContainText('SECOND-EDIT');
 await page.waitForTimeout(550);await page.keyboard.press('Tab');
 await expect.poll(()=>editor.locator('li').evaluateAll(items=>items.map(li=>li.getAttribute('data-list-label')))).toEqual(['4.','4.1.','4.1.1.','5.']);
 await expect.poll(()=>editor.locator('p').filter({hasText:'SECOND-EDIT'}).evaluate(p=>getComputedStyle(p).marginLeft)).toBe('72px');
 await page.keyboard.press('ControlOrMeta+z');await expect.poll(()=>editor.locator('li').evaluateAll(items=>items.map(li=>li.getAttribute('data-list-label')))).toEqual(['4.','5.','5.1.','6.']);
 await page.keyboard.press('ControlOrMeta+Shift+z');await expect.poll(()=>editor.locator('li').evaluateAll(items=>items.map(li=>li.getAttribute('data-list-label')))).toEqual(['4.','4.1.','4.1.1.','5.']);
 await page.waitForTimeout(550);await page.keyboard.press('Shift+Tab');await expect.poll(()=>editor.locator('li').evaluateAll(items=>items.map(li=>li.getAttribute('data-list-label')))).toEqual(['4.','5.','5.1.','6.']);
 await expect.poll(()=>editor.locator('p').filter({hasText:'SECOND-EDIT'}).evaluate(p=>getComputedStyle(p).marginLeft)).toBe('48px');
 const download=page.waitForEvent('download');await page.keyboard.press('ControlOrMeta+s');const importedModel=JSON.parse(await fs.readFile((await(await download).path())!,'utf8'));
 const result=await page.evaluate(async model=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts');
  const saves=[],stages=[];
  for(let i=0;i<3;i++){
   const html=generateHTML(model.content,editorExtensions),layout=await createPrintLayout(html,model.page);
   try{const ps=[...layout.element.querySelectorAll('.print-content li > p')];stages.push({labels:[...layout.element.querySelectorAll('.print-content li')].map(l=>(l as HTMLElement).dataset.listLabel),paragraphs:ps.map(p=>({text:p.textContent,before:getComputedStyle(p).marginTop,after:getComputedStyle(p).marginBottom,height:p.getBoundingClientRect().height,top:p.getBoundingClientRect().top})),model:model.content})}finally{layout.dispose()}
   if(i<2){const bytes=[...new Uint8Array(await(await exportDocx(model)).arrayBuffer())];saves.push(bytes);const opened=await importFile(new File([new Uint8Array(bytes)],'saved.docx'));model={...opened.document,content:parseDocumentHtml(opened.html!)} as typeof model}
  }
  return{saves,stages};
 },importedModel);
 for(const stage of result.stages){expect(stage.labels).toEqual(['4.','5.','5.1.','6.']);expect(stage.paragraphs.map(p=>p.text)).toEqual(['FIRST','SECOND-EDIT','NESTED','AFTER']);expect(stage.paragraphs[1].top-stage.paragraphs[0].top).toBeCloseTo(stage.paragraphs[0].height,1);expect(stage.paragraphs[0].after).toBe('0px');expect(stage.paragraphs[1].before).toBe('0px')}
 for(let i=0;i<2;i++)await fs.writeFile(testInfo.outputPath(`list-save-lines-${i+1}.docx`),Buffer.from(result.saves[i]));
});

test('header and footer list markers use their own first paragraph metrics',async({page})=>{
 await page.goto('/');
 const result=await page.evaluate(async()=>{
  const{createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
  const html='<ol><li data-list-id="Furniture" data-list-level="0" data-list-format="decimal" data-list-pattern="%1."><p data-word-line-rule="exact" style="font-family:Arial;font-size:8pt;line-height:12px;margin:0">FIRST</p></li><li data-list-id="Furniture" data-list-level="0" data-list-format="decimal" data-list-pattern="%1."><p data-word-line-rule="exact" style="font-family:Arial;font-size:8pt;line-height:12px;margin:0">SECOND</p></li></ol>';
  const layout=await createPrintLayout('<p>BODY</p>',{...defaultPage,headerHtml:html,footerHtml:html});
  try{return['.print-header','.print-footer .document-furniture'].map(selector=>{const root=layout.element.querySelector(selector)!,ps=[...root.querySelectorAll('p')];return{gap:ps[1].getBoundingClientRect().top-ps[0].getBoundingClientRect().top,height:ps[0].getBoundingClientRect().height,margin:getComputedStyle(root.querySelector('ol')!).marginBottom,items:[...root.querySelectorAll('li')].map(li=>({font:getComputedStyle(li).fontSize,pFont:getComputedStyle(li.querySelector('p')!).fontSize,line:getComputedStyle(li).lineHeight,pLine:getComputedStyle(li.querySelector('p')!).lineHeight}))}})}finally{layout.dispose()}
 });
 for(const story of result){expect(story.gap).toBeCloseTo(story.height,1);expect(story.margin).toBe('0px');for(const item of story.items){expect(item.font).toBe(item.pFont);expect(item.line).toBe(item.pLine)}}
});

test('each imported numbering level retains its own start across continuation shells and two saves',async({page},testInfo)=>{
 const source=await Packer.toBuffer(new Document({numbering:{config:[{reference:'Multi',levels:[4,2,7].map((start,level)=>({start,level,format:LevelFormat.DECIMAL,text:Array.from({length:level+1},(_,i)=>`%${i+1}`).join('.')+'.'}))}]},sections:[{children:[
  ...[['ROOT-A',0],['CHILD-A',1],['DEEP',2]].map(([text,level])=>new Paragraph({text:String(text),numbering:{reference:'Multi',level:Number(level)}})),
  new Paragraph({text:'INTERRUPTION'}),
  ...[['ROOT-B',0],['CHILD-B',1],['CHILD-C',1],['ROOT-C',0]].map(([text,level])=>new Paragraph({text:String(text),numbering:{reference:'Multi',level:Number(level)}})),
 ]}]}));
 await page.goto('/');
 const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts');
  let file=new File([new Uint8Array(bytes)],'levels.docx');const stages=[],saves=[];
  for(let i=0;i<3;i++){
   const opened=await importFile(file),model={...opened.document,content:parseDocumentHtml(opened.html!)};
   const layout=await createPrintLayout(generateHTML(model.content,editorExtensions),model.page);
   try{stages.push({text:layout.element.querySelector('.print-content')!.textContent,labels:[...layout.element.querySelectorAll('.print-content li')].map(li=>(li as HTMLElement).dataset.listLabel)})}finally{layout.dispose()}
   if(i<2){const bytes=[...new Uint8Array(await(await exportDocx(model)).arrayBuffer())];saves.push(bytes);file=new File([new Uint8Array(bytes)],'levels.docx')}
  }
  return{stages,saves};
 },[...source]);
 for(const stage of result.stages){expect(stage.labels).toEqual(['4.','4.2.','4.2.7.','5.','5.2.','5.3.','6.']);expect(stage.text).toBe('ROOT-ACHILD-ADEEPINTERRUPTIONROOT-BCHILD-BCHILD-CROOT-C')}
 for(let i=0;i<2;i++)await fs.writeFile(testInfo.outputPath(`list-save-levels-${i+1}.docx`),Buffer.from(result.saves[i]));
});

test('Word hanging indents place markers separately from body and nested continuation lines',async({page})=>{
 await page.goto('/');await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
 const result=await page.evaluate(async()=>{
  const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{editorExtensions,parseDocumentHtml}=await import('/src/extensions.ts'),{createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
  const p=(text:string,left:number)=>`<p data-word-line-rule="auto" style="font-family:Arial;font-size:10pt;line-height:1.15;margin:0 0 0 ${left}px;text-indent:-24px">${text}</p>`;
  const li=(text:string,level:number,left:number)=>`<li data-list-id="Geometry" data-list-level="${level}" data-list-format="decimal" data-list-pattern="%${level+1}.">${p(text,left)}</li>`;
  const html='<ol>'+li('First paragraph keeps its continuation at the source text indent. '.repeat(10),0,48)+li('PARENT',0,48).replace('</li>','<ol>'+li('Nested paragraph keeps the absolute source indent. '.repeat(8),1,72)+'</ol></li>')+'</ol>'+p('UNNUMBERED',48);
  const mount=document.createElement('div');mount.style.width='600px';document.body.append(mount);const editor=new Editor({element:mount,extensions:editorExtensions,content:parseDocumentHtml(html)});
  const snapshot=(root:HTMLElement)=>[...root.querySelectorAll('p')].filter(p=>p.textContent).map(p=>{
   const host=p.closest('.print-content,.tiptap')!,origin=host.getBoundingClientRect().left,lines=[];
   const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);while(walker.nextNode()){const text=walker.currentNode as Text;for(let i=0;i<text.length;i++){if(/\s/.test(text.data[i]))continue;const range=document.createRange();range.setStart(text,i);range.setEnd(text,i+1);const r=range.getBoundingClientRect();if(!lines.length||Math.abs(lines[lines.length-1].y-r.y)>2)lines.push({x:r.x-origin,y:r.y})}}
   const item=p.parentElement!.tagName==='LI'?p.parentElement:null;return{text:p.textContent,lines,indent:getComputedStyle(p).textIndent,marker:item?.hasAttribute('data-word-list-layout')?getComputedStyle(item,'::before').left:null};
  });
  try{await new Promise(r=>setTimeout(r,150));const editing=snapshot(editor.view.dom),model=editor.getJSON(),serialized=editor.getHTML();const layout=await createPrintLayout(serialized,defaultPage);try{return{editing,printing:snapshot(layout.element),model,serialized}}finally{layout.dispose()}}finally{editor.destroy();mount.remove()}
 });
 for(const stage of [result.editing,result.printing]){
  for(let i=0;i<3;i++){const expected=i===2?72:48;expect(stage[i].indent).toBe('0px');for(const line of stage[i].lines)expect(line.x).toBeCloseTo(expected,1);expect(parseFloat(stage[i].marker!)).toBe(expected-24)}
  expect(stage[0].lines.length).toBeGreaterThan(1);expect(stage[2].lines.length).toBeGreaterThan(1);expect(stage[3].indent).toBe('-24px');
 }
 expect(result.serialized).toContain('text-indent: -24px');expect(result.serialized).not.toContain('data-word-list-layout');expect(result.serialized).not.toContain('data-word-list-paragraph');expect(result.serialized).not.toContain('--word-list-marker');
});

test('a label wider than its hanging slot keeps a native marker instead of overlapping body text',async({page})=>{
 await page.goto('/');
 const result=await page.evaluate(async()=>{
  const{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts');
  const html='<ol><li data-list-id="Wide" data-list-level="0" data-list-format="decimal" data-list-pattern="LongMarkerThatCannotFit"><p data-word-line-rule="auto" style="font-family:Arial;font-size:10pt;line-height:1.15;margin:0 0 0 48px;text-indent:-24px">BODY</p></li></ol>';
  const layout=await createPrintLayout(generateHTML(parseDocumentHtml(html),editorExtensions),defaultPage);try{const li=layout.element.querySelector('li')!,p=li.querySelector('p')!;return{custom:li.hasAttribute('data-word-list-layout'),indent:getComputedStyle(p).textIndent,native:getComputedStyle(li,'::marker').content}}finally{layout.dispose()}
 });
 expect(result.custom).toBe(false);expect(result.indent).toBe('-24px');expect(result.native).toContain('LongMarkerThatCannotFit');
});

test('numbering space and nothing suffixes survive two standard saves and keep the safe native layout',async({page},testInfo)=>{
 const{LevelSuffix}=await import('docx');
 const input=await Packer.toBuffer(new Document({styles:{default:{document:{run:{font:'Arial',size:20}}}},numbering:{config:[{reference:'Suffix',levels:[LevelSuffix.SPACE,LevelSuffix.NOTHING].map((suffix,level)=>({level,format:LevelFormat.DECIMAL,text:level?'%1.%2.':'%1.',suffix,style:{paragraph:{indent:{left:720+level*360,hanging:360}}}}))}]},sections:[{children:[new Paragraph({text:'SPACE',numbering:{reference:'Suffix',level:0}}),new Paragraph({text:'NOTHING',numbering:{reference:'Suffix',level:1}})]}]}));
 await page.goto('/');
 const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts');let file=new File([new Uint8Array(bytes)],'suffix.docx');const stages=[],saves=[];
  for(let i=0;i<3;i++){
   const opened=await importFile(file),model={...opened.document,content:parseDocumentHtml(opened.html!)};const layout=await createPrintLayout(generateHTML(model.content,editorExtensions),model.page);
   try{stages.push([...layout.element.querySelectorAll('li')].map(li=>({suffix:(li as HTMLElement).dataset.listSuffix,custom:li.hasAttribute('data-word-list-layout'),indent:getComputedStyle(li.querySelector('p')!).textIndent})))}finally{layout.dispose()}
   if(i<2){const bytes=[...new Uint8Array(await(await exportDocx(model)).arrayBuffer())];saves.push(bytes);file=new File([new Uint8Array(bytes)],'suffix.docx')}
  }
  return{stages,saves};
 },[...input]);
 for(const stage of result.stages){expect(stage.map(li=>li.suffix)).toEqual(['space','nothing']);for(const li of stage){expect(li.custom).toBe(false);expect(li.indent).toBe('-24px')}}
 for(let i=0;i<2;i++)await fs.writeFile(testInfo.outputPath(`list-save-suffix-${i+1}.docx`),Buffer.from(result.saves[i]));
});
