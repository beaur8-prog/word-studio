import {test,expect} from '@playwright/test';

const END_STYLE='WordStudioTableEnd';
for(const trailing of ['table','empty','text']) test(`DOCX final ${trailing} keeps authored content and supplies only a structural table end`,async({page})=>{
 await page.goto('/');
 const states=await page.evaluate(async trailing=>{
  const {importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts'),{default:JSZip}=await import('/node_modules/.vite/deps/jszip.js');
  const html='<p>LEAD</p><table><tr><td><p>LAST CELL</p></td></tr></table>'+(trailing==='table'?'':trailing==='empty'?'<p style="font-size:24pt;line-height:40pt"></p>':'<p style="font-size:24pt;line-height:40pt">AUTHORED TAIL</p>');
  let input=new File([html],'tail.html');const states=[],W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  for(let save=0;save<3;save++){
   const loaded=await importFile(input),doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};
   const blob=await exportDocx(doc),zip=await JSZip.loadAsync(await blob.arrayBuffer()),xml=new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'),'application/xml');
   const body=xml.getElementsByTagNameNS(W,'body')[0],blocks=[...body.children].filter(c=>c.localName!=='sectPr'),end=blocks.at(-1)!,pr=[...end.children].find(c=>c.localName==='pPr'),flag=(name:string)=>pr&&[...pr.children].find(c=>c.localName===name),spacing=flag('spacing');
   states.push({blocks:doc.content.content!.map(n=>n.type),lastTag:end.localName,lastStyle:flag('pStyle')?.getAttributeNS(W,'val'),lastFont:pr?.getElementsByTagNameNS(W,'sz')[0]?.getAttributeNS(W,'val'),line:spacing?.getAttributeNS(W,'line'),after:spacing?.getAttributeNS(W,'after'),text:[...xml.getElementsByTagNameNS(W,'t')].map(t=>t.textContent).join(''),structural:blocks.filter(p=>p.getElementsByTagNameNS(W,'pStyle')[0]?.getAttributeNS(W,'val')==='WordStudioTableEnd').length});
   input=new File([blob],'again.docx');
  }
  return states;
 },trailing);
 for(const s of states){
  expect(s.blocks).toEqual(trailing==='table'?['paragraph','table']:['paragraph','table','paragraph']);
  expect(s.text).toBe('LEADLAST CELL'+(trailing==='text'?'AUTHORED TAIL':''));expect(s.lastTag).toBe('p');
  if(trailing==='table'){expect(s.lastStyle).toBe(END_STYLE);expect(s.structural).toBe(1);expect(s.lastFont).toBe('2');expect(s.line).toBe('20');expect(s.after).toBe('0');}
  else {expect(s.structural).toBe(0);expect(s.lastFont).toBe('48');expect(s.line).toBe('800');}
 }
});

for(const edit of ['text','spacing','bookmark','move','run-format']) test(`external ${edit} on a table-end paragraph stays editable`,async({page})=>{
 await page.goto('/');
 const result=await page.evaluate(async edit=>{
  const {importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts'),{default:JSZip}=await import('/node_modules/.vite/deps/jszip.js');
  const loaded=await importFile(new File(['<p>LEAD</p><table><tr><td><p>CELL</p></td></tr></table>'],'end.html')),doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};
  const blob=await exportDocx(doc),zip=await JSZip.loadAsync(await blob.arrayBuffer()),source=await zip.file('word/document.xml')!.async('string');
  const xml=new DOMParser().parseFromString(source,'application/xml'),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',body=xml.getElementsByTagNameNS(W,'body')[0],p=[...body.children].find(p=>p.getElementsByTagNameNS(W,'pStyle')[0]?.getAttributeNS(W,'val')==='WordStudioTableEnd');
  if(!p)throw Error('Missing structural end paragraph');
  const element=(name:string,attrs:Record<string,string>={})=>{const e=xml.createElementNS(W,'w:'+name);for(const[k,v]of Object.entries(attrs))e.setAttributeNS(W,'w:'+k,v);return e};
  if(edit==='run-format'){const r=element('r'),pr=element('rPr');pr.append(element('b'));r.append(pr);p.append(r);}
  if(edit==='text'){const r=element('r'),t=element('t');t.textContent='EXTERNAL TAIL';r.append(t);p.append(r);}
  if(edit==='spacing')p.getElementsByTagNameNS(W,'spacing')[0].setAttributeNS(W,'w:line','400');
  if(edit==='bookmark'){p.append(element('bookmarkStart',{id:'100',name:'USER-END'}),element('bookmarkEnd',{id:'100'}));}
  if(edit==='move'){const after=element('p'),r=element('r'),t=element('t');t.textContent='CONTINUATION';r.append(t);after.append(r);p.after(after);}
  zip.file('word/document.xml',new XMLSerializer().serializeToString(xml));const reopened=await importFile(new File([await zip.generateAsync({type:'blob'})],'external.docx'));const content=parseDocumentHtml(reopened.html!);
  return{types:content.content!.map(n=>n.type),json:JSON.stringify(content)};
 },edit);
 expect(result.types.slice(0,3)).toEqual(['paragraph','table','paragraph']);
 if(edit==='text')expect(result.json).toContain('EXTERNAL TAIL');
 if(edit==='spacing')expect(result.json).toContain('20pt');
 if(edit==='bookmark')expect(result.json).toContain('USER-END');
 if(edit==='move')expect(result.json).toContain('CONTINUATION');
});
