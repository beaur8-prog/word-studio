import {test,expect} from '@playwright/test';
import {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,LineRuleType,WidthType,BorderStyle} from 'docx';
import JSZip from 'jszip';
import {writeFile} from 'node:fs/promises';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const modes=['default','direct-allow','direct-avoid','style-avoid','style-allow','conditional','default-style-avoid'] as const;
async function fixture(mode:typeof modes[number]){
 const p=(text:string,line=20)=>new Paragraph({children:[new TextRun({text,font:'Arial',size:18})],spacing:{before:0,after:0,line:line*15,lineRule:LineRuleType.EXACT}});
 const no={style:BorderStyle.NONE,size:0,color:'auto'};
 const table=new Table({width:{size:4500,type:WidthType.DXA},columnWidths:[2250,2250],style:(mode.startsWith('style')||mode.startsWith('direct'))?'SplitChild':mode==='conditional'?'Conditional':'',borders:{top:no,bottom:no,left:no,right:no,insideHorizontal:no,insideVertical:no},rows:[new TableRow({...(mode.startsWith('direct')?{cantSplit:mode==='direct-avoid'}:{}),children:['L','R'].map(prefix=>new TableCell({margins:{top:0,bottom:0,left:0,right:0},children:[0,1,2].map(i=>p(prefix+i))}))})]});
 const bytes=await Packer.toBuffer(new Document({sections:[{properties:{page:{size:{width:6000,height:6000},margin:{top:750,bottom:750,left:750,right:750}}},children:[p('LEAD',250),table]}]}));
 const zip=await JSZip.loadAsync(bytes),styles=await zip.file('word/styles.xml')!.async('string');
 zip.file('word/styles.xml',styles.replace('</w:styles>',`<w:style w:type="table" ${mode==='default-style-avoid'?'w:default="1"':''} w:styleId="SplitBase"><w:name w:val="Split base"/><w:trPr><w:cantSplit/></w:trPr></w:style><w:style w:type="table" w:styleId="SplitChild"><w:name w:val="Split child"/><w:basedOn w:val="SplitBase"/>${mode==='style-allow'?'<w:trPr><w:cantSplit w:val="false"/></w:trPr>':''}</w:style><w:style w:type="table" w:styleId="Conditional"><w:name w:val="Conditional"/><w:tblStylePr w:type="firstRow"><w:trPr><w:cantSplit/></w:trPr></w:tblStylePr></w:style></w:styles>`));
 return zip.generateAsync({type:'nodebuffer'});
}
for(const mode of modes)test(`Word row policy ${mode} survives two saves without losing cells or source defaults`,async({page})=>{
 const input=await fixture(mode);await writeFile(test.info().outputPath(`textbox-save-input-row-${mode}.docx`),input);await page.goto('/');
 const result=await page.evaluate(async({bytes,mode})=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts');
  const states:any[]=[],saves:number[][]=[];let file=new File([new Uint8Array(bytes)],'row.docx'),doc:any;
  for(let stage=0;stage<3;stage++){
   const loaded=await importFile(file);doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};const root=document.createElement('div');root.innerHTML=generateHTML(doc.content,editorExtensions);const layout=await createPrintLayout(root.innerHTML,doc.page);
   try{const row=doc.content.content.find((n:any)=>n.type==='table').content[0];states.push({attrs:row.attrs,cells:row.content.map((c:any)=>c.content.map((p:any)=>p.content?.map((t:any)=>t.text||'').join('')).join('|')),pages:[...layout.element.querySelectorAll('.print-content')].map(e=>e.textContent!.replace(/\s/g,'')),zoomed:[...layout.element.querySelectorAll<HTMLElement>('.print-content [style*="zoom"]')].map(e=>e.style.zoom)});}finally{layout.dispose();}
   if(stage<2){const b=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(b);file=new File([new Uint8Array(b)],'again.docx');}
  }return{doc,states,saves};
 },{bytes:[...input],mode});
 await writeFile(test.info().outputPath('row-policy-states.json'),JSON.stringify(result.states,null,2));
 const allowed=['default','direct-allow','style-allow'].includes(mode);
 for(const state of result.states){expect(state.cells).toEqual(['L0|L1|L2','R0|R1|R2']);expect(state.pages).toHaveLength(2);expect(state.zoomed).toEqual([]);expect(state.pages[0].includes('L0'),JSON.stringify(state)).toBe(allowed);expect(state.pages[0].includes('R0')).toBe(allowed);expect(state.pages.join('')).toContain('L2');expect(state.pages.join('')).toContain('R2');}
 for(const[i,bytes]of result.saves.entries()){
  const buffer=Buffer.from(bytes);await writeFile(test.info().outputPath(`textbox-save-row-${mode}-${i+1}.docx`),buffer);const zip=await JSZip.loadAsync(buffer),xml=await zip.file('word/document.xml')!.async('string');const flags=[...xml.matchAll(/<w:cantSplit(?:\s[^>]*?)?\/?>(?:<\/w:cantSplit>)?/g)].map(m=>!/(?:w:val="(?:false|0|off)")/.test(m[0]));expect(flags).toEqual(mode==='default'||mode==='style-allow'?[]:[mode!=='direct-allow']);
 }
 // Same semantic state in the actual editor, including different screen zoom.
 await page.locator('input[type=file]').first().setInputFiles({name:'row.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(result.doc))});const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor.locator('table')).toHaveCount(1);await expect(editor.locator('tr:not(.pagination-row-gap) > td')).toHaveCount(2);
 for(const zoom of [100,50,125,100]){const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');await expect.poll(()=>editor.locator('td p').first().evaluate((p,zoom)=>{const root=p.closest<HTMLElement>('.tiptap')!,sheet=root.closest<HTMLElement>('.document-stack')!,y=(p.getBoundingClientRect().top-root.getBoundingClientRect().top)/(zoom/100);return Math.floor((y+.5)/(Number(sheet.dataset.pageHeight)+24))+1;},zoom)).toBe(allowed?1:2);}
 await writeFile(test.info().outputPath('row-policy-states.json'),JSON.stringify(result.states,null,2));
});
