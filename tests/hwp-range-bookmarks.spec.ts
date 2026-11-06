import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
import {createRequire} from 'node:module';
const fixture=createRequire(import.meta.url)('./hwp-range-bookmark-fixture.cjs');
let samples:Record<string,any>;
test.beforeAll(async()=>{samples=await fixture();});
for(const id of ['hwp-ordinary','hwpx-ordinary','hwp-nested','hwpx-nested','unknown','invalid-name'])test(`${id}: source range bookmarks preserve text, names and print through two DOCX saves`,async({page},info)=>{
 await page.goto('/');const source=samples[id];
 const result=await page.evaluate(async({bytes,format,negative})=>{const{HwpReaderSession}=await import('/src/hwp/reader.ts'),{importEditableHwp}=await import('/src/hwp/editing.ts'),{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/@tiptap/core/dist/index.js'),{createPrintLayout}=await import('/src/print-layout.ts');const session=new HwpReaderSession();let doc;try{await session.open(new File([new Uint8Array(bytes)],'bookmarks.'+format));doc=importEditableHwp(await session.editable(),'bookmarks.'+format);}finally{session.close();}const states=[],files=[],logical=n=>n.type==='hardBreak'?'\n':n.type==='documentTab'?'\t':(n.text||'')+(n.content||[]).map(logical).join('');for(let save=0;save<(negative?1:3);save++){const names=[];const walk=n=>{if(n.type==='documentBookmark')names.push(n.attrs.name);n.content?.forEach(walk);};walk(doc.content);const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);try{states.push({save,names,text:logical(doc.content),pages:layout.pages,printText:[...layout.element.querySelectorAll('.print-content')].map(p=>p.textContent).join('')});}finally{layout.dispose();}if(save<2&&!negative){const blob=await exportDocx(doc),bytes=new Uint8Array(await blob.arrayBuffer());let binary='';for(let i=0;i<bytes.length;i+=8192)binary+=String.fromCharCode(...bytes.subarray(i,i+8192));files.push(btoa(binary));const opened=await importFile(new File([blob],'bookmarks.docx'));doc={...opened.document,content:parseDocumentHtml(opened.html)};}}return{states,files};},{bytes:[...source.bytes],format:source.format,negative:!!source.negative});
 for(const state of result.states){
  expect(state.pages).toBe(1);expect(state.printText).toBe(state.text);
  expect(state.text).toContain('BOOKMARK_BODY');expect(state.text).toContain('BOOKMARK_CELL');
  if(source.negative){expect(state.names).toEqual([]);expect(state.text).toContain('원본 보기에서 확인');}
  else {expect(state.names).toEqual(['정부 범위 BODY','정부 범위 CELL']);expect(state.text).toBe('BODY PRE BOOKMARK_BODY BODY TAILCELL PRE BOOKMARK_CELL CELL TAILEND');expect(state.text).toBe(result.states[0].text);}
 }
 for(const[i,file]of result.files.entries())await fs.writeFile(info.outputPath(`hwp-gov-range-bookmarks-${id}-save${i+1}.docx`),Buffer.from(file,'base64'));
});
