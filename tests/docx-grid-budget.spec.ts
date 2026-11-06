import {test,expect} from '@playwright/test';
import {Document,Packer,Paragraph} from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
for(const merged of [false,true])test(`synthetic DOCX ${merged?'merged':'sixteen-column'} grid keeps its total pixel budget through two saves`,async({page},info)=>{
 const zip=await JSZip.loadAsync(await Packer.toBuffer(new Document({sections:[{children:[new Paragraph('fixture')]}]})));
 const border='<w:tcBorders>'+['top','left','bottom','right'].map(s=>`<w:${s} w:val="single" w:sz="6" w:color="000000"/>`).join('')+'</w:tcBorders>';
 const cell=(text:string,props='')=>`<w:tc><w:tcPr>${props}${border}<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="Studio Text" w:eastAsia="Studio Text"/><w:sz w:val="16"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Studio Text" w:eastAsia="Studio Text"/><w:sz w:val="16"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
 const words:string[]=[];
 const number=(n:number)=>{const s=String(100000+n);words.push(s);return s;};
 const rows=Array.from({length:merged?3:26},(_,r)=>{
  if(!merged)return '<w:tr>'+Array.from({length:16},(_,c)=>cell(number(r*16+c))).join('')+'</w:tr>';
  const label=r===2?'':number(r*16),props='<w:gridSpan w:val="2"/>'+(r>0?`<w:vMerge${r===1?' w:val="restart"':''}/>`:'');
  return '<w:tr>'+cell(label,props)+Array.from({length:14},(_,c)=>cell(number(r*16+c+2))).join('')+'</w:tr>';
 }).join('');
 zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr><w:tblLayout w:type="fixed"/><w:tblW w:w="9696" w:type="dxa"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="606"/>'.repeat(16)}</w:tblGrid>${rows}</w:tbl><w:sectPr><w:pgSz w:w="10772" w:h="10205"/><w:pgMar w:left="567" w:right="567" w:top="567" w:bottom="567" w:header="0" w:footer="0"/></w:sectPr></w:body></w:document>`);
 await page.goto('/');
 const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts');
  let loaded=await importFile(new File([new Uint8Array(bytes)],'grid.docx')),doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};const states=[],files=[];
  for(let save=0;save<3;save++){
   const table=doc.content.content!.find(n=>n.type==='table')!,grid=table.content![0].content!.flatMap(c=>c.attrs!.colwidth),layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
   try{states.push({grid,total:grid.reduce((s,n)=>s+n,0),pages:layout.pages,text:layout.element.querySelector('.print-content')!.textContent,allText:[...layout.element.querySelectorAll('.print-content')].map(p=>p.textContent).join(''),cells:table.content!.map(r=>r.content!.map(c=>({rowspan:c.attrs!.rowspan,colspan:c.attrs!.colspan,widths:c.attrs!.colwidth}))),maxLines:Math.max(...[...layout.element.querySelectorAll('p')].filter(p=>/^\d{6}$/.test(p.textContent||'')).map(p=>{const r=document.createRange();r.selectNodeContents(p);const centers:number[]=[];for(const b of r.getClientRects()){if(b.width<.1||b.height<.1)continue;const y=(b.top+b.bottom)/2;if(!centers.some(c=>Math.abs(c-y)<1))centers.push(y);}return centers.length;}))});}finally{layout.dispose();}
   if(save<2){const blob=await exportDocx(doc);files.push([...new Uint8Array(await blob.arrayBuffer())]);loaded=await importFile(new File([blob],'grid.docx'));doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};}
  }
  return{states,files};
 },[...await zip.generateAsync({type:'uint8array'})]);
 await fs.writeFile(info.outputPath('grid-states.json'),JSON.stringify(result.states,null,2));
 for(const[i,bytes]of result.files.entries())await fs.writeFile(info.outputPath(`hwp-gov-grid-budget-${merged?'merged':'columns'}-save${i+1}.docx`),Buffer.from(bytes));
 for(const state of result.states){expect(state.grid).toHaveLength(16);expect(state.total).toBe(640);expect(state.allText).toBe(words.join(''));expect(state.maxLines).toBe(1);expect(state.pages).toBe(1);}
 expect(result.states.map(s=>s.grid)).toEqual([result.states[0].grid,result.states[0].grid,result.states[0].grid]);
 if(merged)expect(result.states[0].cells[1][0]).toMatchObject({rowspan:2,colspan:2,widths:[40,40]});
});
