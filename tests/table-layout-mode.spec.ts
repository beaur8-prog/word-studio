import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-script-cells-fixture.cjs');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
let directory: string, files: Record<string,string>;
test.beforeAll(async()=>{directory=await fs.mkdtemp(path.join(os.tmpdir(),'word-table-mode-'));files=await fixture(directory);});
test.afterAll(async()=>{if(directory)await fs.rm(directory,{recursive:true,force:true});});

async function modes(page:any,bytes:number[],extension:string){
 await page.goto('/');
 return page.evaluate(async({bytes,extension,W})=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts'),{default:JSZip}=await import('/node_modules/.vite/deps/jszip.js');
  let input=new File([new Uint8Array(bytes)],'mode.'+extension),doc:any;
  if(['hwp','hwpx'].includes(extension)){
   const{HwpReaderSession}=await import('/src/hwp/reader.ts'),{importEditableHwp}=await import('/src/hwp/editing.ts');const session=new HwpReaderSession();
   try{await session.open(input);doc=importEditableHwp(await session.editable(),input.name);}finally{session.close();}
  }
  const states=[];
  for(let save=0;save<3;save++){
   if(!doc){const loaded=await importFile(input);doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};}
   const blob=await exportDocx(doc),zip=await JSZip.loadAsync(await blob.arrayBuffer()),xml=new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'),'application/xml');
   states.push([...xml.getElementsByTagNameNS(W,'tbl')].map(t=>{const pr=[...t.children].find(c=>c.localName==='tblPr');return{mode:pr?.getElementsByTagNameNS(W,'tblLayout')[0]?.getAttributeNS(W,'type'),text:[...t.getElementsByTagNameNS(W,'t')].map(t=>t.textContent).join(''),grid:[...([...t.children].find(c=>c.localName==='tblGrid')?.children||[])].map(c=>Number(c.getAttributeNS(W,'w')))};}));
   input=new File([blob],'again.docx');doc=undefined;
  }
  return states;
 },{bytes,extension,W});
}
for(const extension of ['hwp','hwpx'])test(`${extension} ordinary and nested tables export fixed grids through two saves`,async({page})=>{
 const states=await modes(page,[...await fs.readFile(files[extension])],extension);
 for(const state of states){expect(state.length).toBeGreaterThanOrEqual(2);expect(state.every((t:any)=>t.mode==='fixed')).toBe(true);expect(state.map((t:any)=>t.text).join('')).toContain('Nested Latin 789');}
});
test('authored wide and nested tables retain fixed column widths through two saves',async({page})=>{
 const html='<table><tr><td colwidth="300"><p>LONG LEFT</p><table><tr><td colwidth="180"><p>NESTED</p></td></tr></table></td><td colwidth="150"><p>RIGHT</p></td></tr></table>';
 const states=await modes(page,[...Buffer.from(html)],'html');
 for(const state of states){expect(state).toHaveLength(2);expect(state.every((t:any)=>t.mode==='fixed')).toBe(true);expect(state[0].text).toBe('LONG LEFTNESTEDRIGHT');expect(state.map((t:any)=>t.grid)).toEqual(states[0].map((t:any)=>t.grid));expect(state[0].grid[0]/state[0].grid[1]).toBeCloseTo(2,2);}
});
for(const defaultMode of ['fixed','autofit','fixed-true'])test(`external Word table direct/style/default ${defaultMode} policies stay independent through two saves`,async({page})=>{
 const zip=await JSZip.loadAsync(await Packer.toBuffer(new Document({sections:[{children:[new Paragraph('seed')]}]})));
 zip.file('word/styles.xml',`<w:styles xmlns:w="${W}"><w:style w:type="table" w:default="${defaultMode.endsWith('-true')?'true':'1'}" w:styleId="DefaultTable"><w:name w:val="Default"/><w:tblPr><w:tblLayout w:type="${defaultMode.split('-')[0]}"/></w:tblPr></w:style><w:style w:type="table" w:styleId="Base"><w:name w:val="Base"/><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr></w:style><w:style w:type="table" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/></w:style></w:styles>`);
 const props=['','<w:tblStyle w:val="Derived"/>','<w:tblStyle w:val="Derived"/><w:tblLayout w:type="autofit"/>','<w:tblLayout w:type="fixed"/>'];
 zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${props.map((pr,i)=>`<w:tbl><w:tblPr>${pr}<w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:noWrap/></w:tcPr><w:p><w:r><w:t>TABLE-${i}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>RIGHT-${i}</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`).join('')}<w:sectPr/></w:body></w:document>`);
 const states=await modes(page,[...await zip.generateAsync({type:'uint8array'})],'docx');
 for(const state of states){expect(state.map((t:any)=>t.mode)).toEqual([defaultMode.split('-')[0],'fixed','autofit','fixed']);expect(state.map((t:any)=>t.text)).toEqual(props.map((_,i)=>`TABLE-${i}RIGHT-${i}`));}
});
test('omitted Word table layout keeps the standard autofit default',async({page})=>{
 const zip=await JSZip.loadAsync(await Packer.toBuffer(new Document({sections:[{children:[new Paragraph('seed')]}]})));
 zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>AUTO DEFAULT</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`);
 const states=await modes(page,[...await zip.generateAsync({type:'uint8array'})],'docx');for(const state of states)expect(state[0].mode).toBe('autofit');
});

test('authored nowrap cells use autofit without changing widths or leaking the policy to a parent table',async({page})=>{
 const html='<table><tr><td colwidth="300"><p>OUTER</p><table><tr><td colwidth="34" style="white-space:nowrap;padding:6.8px"><p style="font-size:13pt">10.</p></td><td colwidth="180"><p>INNER CONTENT</p></td></tr></table></td></tr></table>';
 const states=await modes(page,[...Buffer.from(html)],'html');
 for(const state of states){expect(state.map((t:any)=>t.mode)).toEqual(['fixed','autofit']);expect(state[1].grid).toEqual([510,2700]);expect(state[1].text).toBe('10.INNER CONTENT');}
});
