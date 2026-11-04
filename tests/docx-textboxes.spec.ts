import { writeFile } from 'node:fs/promises';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const boxText='<w:p><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>VISIBLE BOX TEXT</w:t></w:r></w:p>';
async function fixture(kind:string,after=true) {
 const zip=new JSZip();
 zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
 zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
 zip.file('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
 zip.file('word/styles.xml',`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:color w:val="000000"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`);
 const modern=`<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>190500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="margin"><wp:posOffset>571500</wp:posOffset></wp:positionV><wp:extent cx="2286000" cy="762000"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="10" name="Box"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="205080"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>${boxText}</w:txbxContent></wps:txbx><wps:bodyPr wrap="square" lIns="91440" tIns="0" rIns="91440" bIns="0" anchor="ctr"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
 const legacy=`<w:pict><v:shape id="Box" style="position:absolute;margin-left:20px;margin-top:60px;width:240px;height:80px;mso-position-horizontal-relative:margin;mso-position-vertical-relative:margin" fillcolor="#205080" stroked="f"><v:textbox inset="9.6px,0px,9.6px,0px"><w:txbxContent>${boxText}</w:txbxContent></v:textbox></v:shape></w:pict>`;
 const drawing=kind==='legacy'?legacy:kind==='unwrapped'?modern:`<mc:AlternateContent><mc:Choice Requires="wps">${modern}</mc:Choice>${kind==='mixed'?`<mc:Fallback>${legacy}</mc:Fallback>`:''}</mc:AlternateContent>`;
 zip.file('word/document.xml',`<w:document xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="wps"><w:body><w:p><w:r><w:t>BODY BEFORE</w:t></w:r></w:p><w:p><w:r>${drawing}</w:r></w:p>${after?'<w:p><w:r><w:t>BODY AFTER</w:t></w:r></w:p>':''}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
 return zip.generateAsync({type:'nodebuffer'});
}
for(const kind of ['modern','mixed','unwrapped','legacy']) test(`${kind} textbox retains one editable story, background and geometry through two DOCX saves`,async({page})=>{
 await page.goto('/');const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx,validateDocument}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let file=new File([new Uint8Array(bytes)],'box.docx');const states=[],saves:number[][]=[];
  for(let stage=0;stage<3;stage++){
   const opened=await importFile(file),content=parseDocumentHtml(opened.html!),doc={...opened.document,content};validateDocument(doc);
   const frames:any[]=[];const walk=(n:any)=>{if(n.type==='documentTextBox')frames.push(n);for(const c of n.content||[])walk(c);};walk(content);
   states.push({frames,html:opened.html});if(stage<2){const bytes=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(bytes);file=new File([new Uint8Array(bytes)],'again.docx');}
  }return {states,saves};
 },[...await fixture(kind)]);
 for(let i=0;i<result.saves.length;i++)await writeFile(test.info().outputPath(`textbox-save${i+1}.docx`),Buffer.from(result.saves[i]));
 for(const state of result.states){expect(state.frames).toHaveLength(1);expect(state.html!.split('VISIBLE BOX TEXT')).toHaveLength(2);expect(state.html).toContain('BODY BEFORE');expect(state.html).toContain('BODY AFTER');
  const f=JSON.parse(state.frames[0].attrs.frame);expect(f.fill.toLowerCase()).toBe('#205080');expect(f.width).toBeCloseTo(240,2);expect(f.height).toBeCloseTo(80,2);expect(f.padding[0]).toBe(0);expect(f.x).toBeCloseTo(116,2);expect(f.y).toBeCloseTo(156,2);
 }
});

test('Ctrl/Command End stays in body, new input is visible, textbox editing exits explicitly, undo and redo preserve both stories',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'box.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await fixture('mixed',false)});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('.document-textbox')).toHaveCount(1);
 await editor.locator('p').first().click();await page.keyboard.press('ControlOrMeta+End');await page.keyboard.press('Enter');await page.keyboard.insertText('NEW BODY 글자가 보입니다.');
 const body=editor.locator('p').filter({hasText:'NEW BODY 글자가 보입니다.'});await expect(body).toHaveCSS('color',await editor.evaluate(e=>getComputedStyle(e).color));expect(await body.evaluate(e=>!!e.closest('.document-textbox'))).toBe(false);
 const content=editor.locator('.document-textbox-content');await expect(content).toHaveCSS('background-color','rgb(32, 80, 128)');const white=content.locator('span').filter({hasText:'VISIBLE BOX TEXT'});await expect(white).toHaveCSS('color','rgb(255, 255, 255)');
 await white.click();await page.keyboard.press('ControlOrMeta+End');await page.keyboard.insertText(' EDIT');await expect(content).toContainText('TEXT EDIT');await page.keyboard.press('ControlOrMeta+z');await expect(content).not.toContainText('EDIT');await page.keyboard.press('ControlOrMeta+Shift+z');await expect(content).toContainText('TEXT EDIT');
 await page.keyboard.press('ControlOrMeta+Enter');await page.keyboard.insertText('OUTSIDE BOX');const outside=editor.locator('p').filter({hasText:'OUTSIDE BOX'});expect(await outside.evaluate(e=>!!e.closest('.document-textbox'))).toBe(false);await expect(outside).toHaveCSS('color',await editor.evaluate(e=>getComputedStyle(e).color));
});

test('textbox stories and ordinary body have the same colors and coordinates in editor and print',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'box.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await fixture('modern')});
 const editor=page.locator('.tiptap');await expect(editor.locator('.document-textbox-content')).toBeVisible();
 const measure=async(selector:string)=>page.locator(selector).evaluate(el=>{const p=el.closest('.print-page,.document-stack')!.getBoundingClientRect(),r=el.getBoundingClientRect();return{x:r.left-p.left,y:r.top-p.top,width:r.width,height:r.height,fill:getComputedStyle(el).backgroundColor};});
 await expect.poll(async()=>Math.round((await measure('.tiptap .document-textbox-content')).x)).toBe(116);const before=await measure('.tiptap .document-textbox-content');
 await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();await expect(page.locator('.print-content .document-textbox-content')).toBeVisible();const after=await measure('.print-content .document-textbox-content');
 for(const key of ['x','y','width','height'] as const)expect(after[key]).toBeCloseTo(before[key],0);expect(after.fill).toBe(before.fill);
});

test('a grouped Word drawing remains one anchor with two independently editable stories through two saves',async({page})=>{
 const zip=await JSZip.loadAsync(await fixture('unwrapped'));let xml=await zip.file('word/document.xml')!.async('string');
 const shape=xml.match(/<wps:wsp>[\s\S]*?<\/wps:wsp>/)![0];
 const a=shape.replace('VISIBLE BOX TEXT','FIRST GROUP STORY'),b=shape.replace('VISIBLE BOX TEXT','SECOND GROUP STORY').replace('<a:off x="0" y="0"/>','<a:off x="2381250" y="0"/>');
 const group=`<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4667250" cy="762000"/><a:chOff x="0" y="0"/><a:chExt cx="4667250" cy="762000"/></a:xfrm></wpg:grpSpPr>${a}${b}</wpg:wgp>`;
 xml=xml.replace(shape,group).replace('uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"','uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"');zip.file('word/document.xml',xml);
 await page.goto('/');const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let file=new File([new Uint8Array(bytes)],'group.docx');const states=[],saves:number[][]=[];
  const text=(n:any):string=>(n.text||'')+(n.content||[]).map(text).join('');
  for(let i=0;i<3;i++){const opened=await importFile(file),content=parseDocumentHtml(opened.html!),frames=content.content!.filter(n=>n.type==='documentTextBox').map(n=>({frame:JSON.parse(n.attrs!.frame),text:text(n)}));states.push(frames);if(i<2){const b=[...new Uint8Array(await(await exportDocx({...opened.document,content})).arrayBuffer())];saves.push(b);file=new File([new Uint8Array(b)],'again.docx');}}return{states,saves};
 },[...await zip.generateAsync({type:'nodebuffer'})]);
 for(const frames of result.states){expect(frames.map(f=>f.text)).toEqual(['FIRST GROUP STORY','SECOND GROUP STORY']);expect(frames[0].frame.group).toBe(frames[1].frame.group);expect(frames[0].frame.group).toBeTruthy();expect(frames[1].frame.x-frames[0].frame.x).toBeCloseTo(250,2);}
 for(let i=0;i<result.saves.length;i++){const bytes=Buffer.from(result.saves[i]),z=await JSZip.loadAsync(bytes),xml=await z.file('word/document.xml')!.async('string');expect(xml.match(/<wp:anchor\b/g)).toHaveLength(1);expect(xml.match(/<wps:txbx\b/g)).toHaveLength(2);await writeFile(test.info().outputPath(`textbox-save${i+1}.docx`),bytes);}
});

test('nested modern textboxes retain every story without importing duplicate fallbacks',async({page})=>{
 const zip=await JSZip.loadAsync(await fixture('unwrapped'));let xml=await zip.file('word/document.xml')!.async('string');const drawing=xml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/)![0];
 xml=xml.replace('VISIBLE BOX TEXT','OUTER BOX TEXT').replace('</w:txbxContent>',`<w:p><w:r>${drawing.replace('VISIBLE BOX TEXT','INNER BOX TEXT')}</w:r></w:p></w:txbxContent>`);zip.file('word/document.xml',xml);
 await page.goto('/');const states=await page.evaluate(async bytes=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let f=new File([new Uint8Array(bytes)],'nested.docx');const states=[];const text=(n:any):string=>(n.text||'')+(n.content||[]).map(text).join('');
  for(let i=0;i<3;i++){const opened=await importFile(f),content=parseDocumentHtml(opened.html!);let boxes=0;const walk=(n:any)=>{if(n.type==='documentTextBox')boxes++;for(const c of n.content||[])walk(c);};walk(content);states.push({text:text(content),boxes});if(i<2)f=new File([await exportDocx({...opened.document,content})],'again.docx');}return states;
 },[...await zip.generateAsync({type:'nodebuffer'})]);
 for(const s of states){expect(s.boxes).toBe(2);expect(s.text.match(/OUTER BOX TEXT/g)).toHaveLength(1);expect(s.text.match(/INNER BOX TEXT/g)).toHaveLength(1);expect(s.text).toContain('BODY AFTER');}
});

test('moving to the body end inherits its real bold and color instead of the textbox or a forced default',async({page})=>{
 const zip=await JSZip.loadAsync(await fixture('mixed'));const xml=await zip.file('word/document.xml')!.async('string');zip.file('word/document.xml',xml.replace('<w:r><w:t>BODY AFTER','<w:r><w:rPr><w:b/><w:color w:val="0055AA"/></w:rPr><w:t>BODY AFTER'));
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'body-marks.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await zip.generateAsync({type:'nodebuffer'})});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('.document-textbox')).toHaveCount(1);await editor.locator('p').first().click();await page.keyboard.press('ControlOrMeta+End');await page.keyboard.insertText(' CONTINUED');
 const text=editor.locator('span').filter({hasText:'CONTINUED'});await expect(text).toHaveCSS('color','rgb(0, 85, 170)');await expect(text.locator('strong')).toHaveCSS('font-weight','700');expect(await text.evaluate(e=>!!e.closest('.document-textbox'))).toBe(false);
});

test('percentage anchors in compatibility Choice and wrap clearances retain their actual page coordinates through two saves',async({page})=>{
 const zip=await JSZip.loadAsync(await fixture('unwrapped'));let xml=await zip.file('word/document.xml')!.async('string');
 xml=xml.replace('xmlns:wp=', 'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp=')
  .replace('<wp:positionH relativeFrom="margin"><wp:posOffset>190500</wp:posOffset></wp:positionH>','<mc:AlternateContent><mc:Choice Requires="wp14"><wp:positionH relativeFrom="page"><wp14:pctPosHOffset>4000</wp14:pctPosHOffset></wp:positionH></mc:Choice><mc:Fallback><wp:positionH relativeFrom="page"><wp:posOffset>310515</wp:posOffset></wp:positionH></mc:Fallback></mc:AlternateContent>')
  .replace('distL="0" distR="0"','distL="228600" distR="228600"');zip.file('word/document.xml',xml);await page.goto('/');
 const frames=await page.evaluate(async bytes=>{const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let f=new File([new Uint8Array(bytes)],'percentage.docx');const frames=[];for(let i=0;i<3;i++){const d=await importFile(f),content=parseDocumentHtml(d.html!);let frame:any;const walk=(n:any)=>{if(n.type==='documentTextBox')frame=JSON.parse(n.attrs.frame);for(const c of n.content||[])walk(c);};walk(content);frames.push(frame);if(i<2)f=new File([await exportDocx({...d.document,content})],'again.docx');}return frames;},[...await zip.generateAsync({type:'nodebuffer'})]);
 for(const f of frames){expect(f.x).toBeCloseTo(32.64,2);expect(f.wrapMargins).toEqual([0,24,0,24]);}
});

test('legacy v:rect textbox stays readable instead of failing the whole document',async({page})=>{
 const zip=await JSZip.loadAsync(await fixture('legacy'));const xml=await zip.file('word/document.xml')!.async('string');zip.file('word/document.xml',xml.replace('<v:shape ','<v:rect ').replace('</v:shape>','</v:rect>'));await page.goto('/');
 const states=await page.evaluate(async bytes=>{const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let f=new File([new Uint8Array(bytes)],'rect.docx');const states=[];for(let i=0;i<3;i++){const d=await importFile(f),content=parseDocumentHtml(d.html!),boxes:any[]=[];const walk=(n:any)=>{if(n.type==='documentTextBox')boxes.push(n);for(const c of n.content||[])walk(c);};walk(content);states.push({boxes,html:d.html});if(i<2)f=new File([await exportDocx({...d.document,content})],'again.docx');}return states;},[...await zip.generateAsync({type:'nodebuffer'})]);
 for(const s of states){expect(s.boxes).toHaveLength(1);expect(JSON.parse(s.boxes[0].attrs.frame).fill).toBe('#205080');expect(s.html!.match(/VISIBLE BOX TEXT/g)).toHaveLength(1);expect(s.html).toContain('BODY AFTER');}
});

test('clicks in text wrapped beside a textbox and immediate document-end typing use the painted caret',async({page})=>{
 const zip=await JSZip.loadAsync(await fixture('mixed'));const words=('Wrap around the box while preserving the actual caret position. ').repeat(22)+'BODY END';let xml=await zip.file('word/document.xml')!.async('string');xml=xml.replace('BODY AFTER',words).replace('<wp:posOffset>571500</wp:posOffset>','<wp:posOffset>0</wp:posOffset>');zip.file('word/document.xml',xml);await page.goto('/');
 await page.locator('input[type=file]').first().setInputFiles({name:'wrapped-caret.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:await zip.generateAsync({type:'nodebuffer'})});const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true}),para=editor.locator('p').filter({hasText:'BODY END'}).first();await expect(para.locator('[data-flow-widget]').first()).toBeAttached();
 await para.click({position:{x:410,y:12}});
 const selection=await editor.evaluate((el:any)=>{const n=getSelection()!,v=el.editor.view;return{model:[v.state.selection.anchor,v.state.selection.head],native:[v.posAtDOM(n.anchorNode,n.anchorOffset),v.posAtDOM(n.focusNode,n.focusOffset)]};});expect(selection.model).toEqual(selection.native);
 for(let i=0;i<8;i++){await para.click({position:{x:410,y:12}});await page.keyboard.press('ControlOrMeta+End');await page.keyboard.insertText(' END-TYPING');await expect(para).toHaveText(words+' END-TYPING');await page.keyboard.press('ControlOrMeta+z');await expect(para).toHaveText(words);}
});
