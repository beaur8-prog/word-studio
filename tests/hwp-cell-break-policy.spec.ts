import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
import {createRequire} from 'node:module';
const fixture=createRequire(import.meta.url)('./hwp-cell-break-policy-fixture.cjs');
for(const inline of [false,true])for(const nested of [false,true])for(const format of ['hwp','hwpx'])for(const policy of inline?[2]:[0,1,2])test(`${format} ${nested?'nested':'ordinary'} ${['NONE','TABLE','CELL'][policy]} ${inline?'inline':'floating'}: preceding edits and two DOCX saves keep cell-break permission, full text and grid`,async({page},info)=>{
 const sample=(await fixture())[`${format}-${policy}-${nested?'nested':'ordinary'}${inline?'-inline':''}`];await page.goto('/');
 const result=await page.evaluate(async({bytes,format})=>{
  const {HwpReaderSession}=await import('/src/hwp/reader.ts'),{importEditableHwp}=await import('/src/hwp/editing.ts'),{exportDocx,importFile}=await import('/src/document-io.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{editorExtensions,parseDocumentHtml}=await import('/src/extensions.ts'),{createPrintLayout}=await import('/src/print-layout.ts');
  const session=new HwpReaderSession();
  try{await session.open(new File([new Uint8Array(bytes)],'policy.'+format));let doc=importEditableHwp(await session.editable(),'policy.'+format);
   doc.page={...doc.page,size:'Custom',width:150,height:180,margin:10,margins:undefined,pageNumbers:false};
   const lead=doc.content.content!.find(n=>n.type==='paragraph'&&n.content?.some(t=>t.text==='LEAD'))!;lead.attrs={...lead.attrs,lineSpacing:'350px',spaceBefore:'0px',spaceAfter:'0px'};
   const first=structuredClone(doc),states=[],files=[];
   for(let save=0;save<3;save++){
    const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
    try{const sheets=[...layout.element.querySelectorAll('.print-content')],rows=[...layout.element.querySelectorAll<HTMLTableRowElement>('tr')];
     states.push({pages:sheets.map(p=>p.textContent),columns:[0,1].map(i=>rows.map(r=>r.closest('table')?.parentElement?.closest('table')?(r.cells[i]?.textContent||''):r.cells.length===2?(r.cells[i]?.textContent||''):'').join('')),widths:rows.filter(r=>r.cells.length===2).map(r=>[...r.cells].map(c=>c.getBoundingClientRect().width)),policies:(()=>{const a:unknown[]=[];const visit=(n:any)=>{if(n.type==='tableRow'&&n.content?.length===2)a.push(n.attrs?.rowSplit);n.content?.forEach(visit);};visit(doc.content);return a;})(),scaled:[...layout.element.querySelectorAll<HTMLElement>('[style]')].some(n=>!!n.style.zoom),maxOverflow:Math.max(0,...sheets.flatMap(p=>[...p.querySelectorAll('tr')].map(r=>r.getBoundingClientRect().bottom-p.getBoundingClientRect().bottom)))});
    }finally{layout.dispose();}
    if(save<2){const blob=await exportDocx(doc);files.push([...new Uint8Array(await blob.arrayBuffer())]);const loaded=await importFile(new File([blob],'again.docx'));doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};}
   }
   return{first,states,files};
  }finally{session.close();}
 },{bytes:[...sample.bytes],format});
 for(const [i,bytes]of result.files.entries())await fs.writeFile(info.outputPath(`hwp-gov-cellbreak-${format}-${policy}-${nested?'nested':'ordinary'}-${inline?'inline':'floating'}-save${i+1}.docx`),Buffer.from(bytes));
 for(const state of result.states){expect(state.pages).toHaveLength(2);expect(state.scaled).toBe(false);expect(state.maxOverflow).toBeLessThanOrEqual(.5);expect(state.pages[0]!.includes('LEFT 0')).toBe(policy===2&&!inline);expect(state.pages[1]).toContain('LEFT 23');expect(state.pages.join('')).toContain('END');
  for(const [i,label]of ['LEFT','RIGHT'].entries())expect(state.columns[i]).toBe(Array.from({length:24},(_,n)=>label+' '+n).join(''));
  for(const widths of state.widths){expect(widths).toHaveLength(2);expect(widths[0]).toBeCloseTo(widths[1],0);}
  expect(state.policies).toHaveLength(1);expect(state.policies[0]==='allow').toBe(policy===2&&!inline);
 }
 if(nested){
  await page.locator('input[type=file]').first().setInputFiles({name:'cell-policy.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(result.first))});
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText('LEFT 23');
  const position=(label:string)=>editor.evaluate((e,label)=>{const p=[...e.querySelectorAll('p')].find(p=>p.textContent===label)!;if(!p)throw Error('Missing cell paragraph');let scale=1;for(let n:HTMLElement|null=e;n;n=n.parentElement){const css=getComputedStyle(n),zoom=parseFloat(css.zoom);if(zoom>0)scale*=zoom;if(css.transform!=='none'){const m=new DOMMatrixReadOnly(css.transform);scale*=Math.hypot(m.a,m.b);}}const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);while(walker.nextNode()){if(!walker.currentNode.textContent)continue;const range=document.createRange();range.setStart(walker.currentNode,0);range.setEnd(walker.currentNode,1);const rect=range.getBoundingClientRect(),origin=e.getBoundingClientRect(),height=Number(e.closest<HTMLElement>('.document-stack')!.dataset.pageHeight);return 1+Math.floor(((rect.top-origin.top)/scale+.5)/(height+24));}throw Error('No cell glyph');},label);
  await expect.poll(()=>position('LEFT 0')).toBe(policy===2&&!inline?1:2);
  await expect.poll(()=>position('LEFT 23')).toBe(2);
 }
});
