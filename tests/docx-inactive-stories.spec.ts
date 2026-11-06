import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
for(const mode of ['absent','inactive','mixed','blank-even'])test(`synthetic DOCX ${mode}: header/footer references preserve dormant values without changing active page furniture`,async({page},info)=>{
 await page.goto('/');
 const result=await page.evaluate(async mode=>{
  const{exportDocx,importFile}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts'),{defaultPage}=await import('/src/types.ts'),JSZip=(await import('/node_modules/.vite/deps/jszip.js')).default;
  const settings={...defaultPage,margin:20,pageNumbers:false,header:'DEFAULT HEADER',footer:'DEFAULT FOOTER',differentOddEven:mode==='blank-even'};
  if(mode==='inactive')Object.assign(settings,{evenHeaderHtml:'<p>DORMANT HEADER</p>',evenFooterHtml:'<p>DORMANT FOOTER</p>',evenPageNumbers:false});
  let html='<p>BODY A1</p><div data-page-break="true"></div><p>BODY A2</p>';
  if(mode==='mixed'){const second={...settings,differentOddEven:true,header:'ODD B',footer:'ODD FOOT B',evenHeader:'EVEN B',evenFooter:'EVEN FOOT B'};html+='<div data-section-page="'+JSON.stringify(second).replace(/"/g,'&quot;')+'"></div><p>BODY B1</p><div data-page-break="true"></div><p>BODY B2</p>';}
  let doc:any={format:'word-studio',version:1,id:crypto.randomUUID(),title:'Synthetic story metadata',comments:[],updatedAt:new Date().toISOString(),page:settings,content:parseDocumentHtml(html)};
  const text=(html:string|undefined)=>{const t=document.createElement('template');t.innerHTML=html||'';return t.content.textContent||'';},states=[],files=[],references=[];
  for(let save=0;save<3;save++){
   const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
   try{states.push({pages:layout.pages,header:[...layout.element.querySelectorAll('.print-header')].map(h=>(h.textContent||'').trim()),footer:[...layout.element.querySelectorAll('.print-footer')].map(h=>(h.textContent||'').trim()),differentOddEven:doc.page.differentOddEven,evenHeader:text(doc.page.evenHeaderHtml)||doc.page.evenHeader||'',evenFooter:text(doc.page.evenFooterHtml)||doc.page.evenFooter||''});}finally{layout.dispose();}
   if(save<2){const blob=await exportDocx(doc),zip=await JSZip.loadAsync(await blob.arrayBuffer()),xml=new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'),'application/xml'),ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main';references.push({headers:[...xml.getElementsByTagNameNS(ns,'headerReference')].filter(e=>e.getAttributeNS(ns,'type')==='even').length,footers:[...xml.getElementsByTagNameNS(ns,'footerReference')].filter(e=>e.getAttributeNS(ns,'type')==='even').length});files.push([...new Uint8Array(await blob.arrayBuffer())]);const loaded=await importFile(new File([blob],'stories.docx'));doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};}
  }
  return{states,files,references};
 },mode);
 for(const[i,bytes]of result.files.entries())await fs.writeFile(info.outputPath(`hwp-gov-inactive-stories-${mode}-save${i+1}.docx`),Buffer.from(bytes));
 for(const state of result.states){expect(state.header).toEqual(result.states[0].header);expect(state.footer).toEqual(result.states[0].footer);expect(state.pages).toBe(mode==='mixed'?4:2);if(mode==='inactive'){expect(state.evenHeader).toBe('DORMANT HEADER');expect(state.evenFooter).toBe('DORMANT FOOTER');}if(mode==='blank-even'){expect(state.header).toEqual(['DEFAULT HEADER','']);expect(state.footer).toEqual(['DEFAULT FOOTER','']);}}
 for(const refs of result.references){expect(refs.headers).toBe(mode==='absent'?0:mode==='mixed'?2:1);expect(refs.footers).toBe(mode==='absent'?0:mode==='mixed'?2:1);}
});
