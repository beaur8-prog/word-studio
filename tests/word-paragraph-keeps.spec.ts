import {test,expect} from '@playwright/test';
import {Document,Packer,Paragraph,TextRun,LineRuleType} from 'docx';
import {writeFile} from 'node:fs/promises';
const modes=['inherited-next','direct-off','inherited-off','keep-lines','lines-off','chain','explicit-break','prefix-next','whole-next','oversized-lines','oversized-chain'] as const;
async function fixture(mode:typeof modes[number]){
 const p=(text:string,opts:any={},line=20)=>new Paragraph({children:text.split('~').map((s,i)=>new TextRun({text:s,font:'Arial',size:18,break:i?1:0})),spacing:{before:0,after:0,line:line*15,lineRule:LineRuleType.EXACT},...opts});
 const styles=[{id:'KeepBase',name:'Keep base',paragraph:{keepNext:true,keepLines:true}},{id:'KeptHeading',name:'heading 2',basedOn:'KeepBase'},{id:'ResetHeading',name:'Reset heading',basedOn:'KeptHeading',paragraph:{keepNext:false}}];
 let children:any[];
 if(mode==='keep-lines'||mode==='lines-off')children=[p('LEAD',{},250),p('T0~T1~T2~T3',{keepLines:mode==='keep-lines',widowControl:false})];
 else if(mode==='oversized-lines')children=[p('LEAD',{},40),p(Array.from({length:25},(_,i)=>'T'+i).join('~'),{keepLines:true,widowControl:false})];
 else if(mode==='oversized-chain')children=[p('LEAD',{},80),...Array.from({length:18},(_,i)=>p('HEAD'+i,{style:'KeptHeading'})),p('END')];
 else if(mode==='chain')children=[p('LEAD',{},250),p('HEAD_A',{style:'KeptHeading'}),p('HEAD_B',{style:'KeptHeading'}),p('BODY')];
 else if(mode==='prefix-next'||mode==='whole-next')children=[p('LEAD',{},200),p('HEADING',{style:'KeptHeading'}),p(Array.from({length:10},(_,i)=>'T'+i).join('~'),{widowControl:true,keepLines:mode==='whole-next'})];
 else children=[p('LEAD',{},270),p('HEADING',{style:mode==='inherited-off'?'ResetHeading':'KeptHeading',...(mode==='direct-off'?{keepNext:false}:{})}),p('BODY',mode==='explicit-break'?{pageBreakBefore:true}:{})];
 return Packer.toBuffer(new Document({styles:{paragraphStyles:styles},sections:[{properties:{page:{size:{width:6000,height:6000},margin:{top:750,right:750,bottom:750,left:750}}},children}]}));
}
for(const mode of modes)test(`Word paragraph keeps ${mode} preserve text and authored pagination through two saves`,async({page})=>{
 const input=await fixture(mode);await writeFile(test.info().outputPath(`textbox-save-input-keeps-${mode}.docx`),input);await page.goto('/');
 const result=await page.evaluate(async bytes=>{
  const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts'),{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts');
  let file=new File([new Uint8Array(bytes)],'keeps.docx'),doc:any;const states:any[]=[],saves:number[][]=[];
  for(let stage=0;stage<3;stage++){const loaded=await importFile(file);doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);try{states.push({content:doc.content,pages:[...layout.element.querySelectorAll('.print-content')].map(p=>p.textContent!.replace(/\s/g,'')),zoomed:[...layout.element.querySelectorAll<HTMLElement>('.print-content [style*="zoom"]')].map(e=>e.style.zoom)});}finally{layout.dispose();}if(stage<2){const b=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(b);file=new File([new Uint8Array(b)],'again.docx');}}
  return{doc,states,saves};
 },[...input]);
 await writeFile(test.info().outputPath('keeps-states.json'),JSON.stringify(result.states,null,2));for(const[i,b]of result.saves.entries())await writeFile(test.info().outputPath(`textbox-save-keeps-${mode}-${i+1}.docx`),Buffer.from(b));
 const off=['direct-off','inherited-off','explicit-break'].includes(mode);
 const expected=mode==='chain'?['LEAD','HEAD_AHEAD_BBODY']:off?['LEADHEADING','BODY']:mode==='inherited-next'?['LEAD','HEADINGBODY']:mode==='keep-lines'?['LEAD','T0T1T2T3']:mode==='lines-off'?['LEADT0T1','T2T3']:mode==='prefix-next'?['LEADHEADINGT0T1T2T3','T4T5T6T7T8T9']:mode==='whole-next'?['LEAD','HEADING'+Array.from({length:10},(_,i)=>'T'+i).join('')]:null;
 for(const state of result.states){
  const paragraphs:any[]=[];const walk=(n:any)=>{if(['paragraph','heading'].includes(n.type))paragraphs.push(n);for(const c of n.content||[])walk(c);};walk(state.content);
  if(['inherited-next','direct-off','inherited-off','explicit-break','prefix-next','whole-next'].includes(mode)){expect(paragraphs[1].attrs.wordKeepNext).toBe(!['direct-off','inherited-off'].includes(mode));expect(paragraphs[1].attrs.wordKeepLines).toBe(true);}
  if(['keep-lines','lines-off','oversized-lines'].includes(mode))expect(paragraphs[1].attrs.wordKeepLines).toBe(mode!=='lines-off');
  expect(state.zoomed).toEqual([]);if(expected)expect(state.pages).toEqual(expected);if(mode==='oversized-lines'){expect(state.pages[0]).toBe('LEAD');expect(state.pages.join('')).toBe('LEAD'+Array.from({length:25},(_,i)=>'T'+i).join(''));expect(state.pages.length).toBeLessThanOrEqual(4);}if(mode==='oversized-chain'){expect(state.pages.join('')).toBe('LEAD'+Array.from({length:18},(_,i)=>'HEAD'+i).join('')+'END');expect(state.pages.length).toBeLessThanOrEqual(4);}}
 await page.locator('input[type=file]').first().setInputFiles({name:'keeps.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(result.doc))});const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText('LEAD');
 for(const zoom of [100,50,125,100]){const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');await expect.poll(()=>editor.evaluate((root,zoom)=>{const sheet=root.closest<HTMLElement>('.document-stack')!,h=Number(sheet.dataset.pageHeight),origin=root.getBoundingClientRect().top,pages:string[]=[];const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);while(walker.nextNode()){const n=walker.currentNode as Text;if(n.parentElement?.closest('.pagination-gap,.pagination-cell-gap,.page-break,.column-break,.section-break,[data-flow-widget]'))continue;for(let i=0;i<n.length;i++){if(/\s/.test(n.data[i]))continue;const range=document.createRange();range.setStart(n,i);range.setEnd(n,i+1);const r=range.getBoundingClientRect(),y=(r.top-origin)/(zoom/100),page=Math.max(0,Math.floor((y+.5)/(h+24)));pages[page]=(pages[page]||'')+n.data[i];}}return pages;},zoom)).toEqual(expected || result.states[0].pages);}
 const before=await editor.evaluate((e:any)=>e.editor.getJSON());await editor.locator('p,h1,h2,h3').last().click();await page.keyboard.press(process.platform==='darwin'?'Meta+ArrowDown':'Control+End');await page.keyboard.insertText(' EDIT');await expect(editor).toContainText(' EDIT');await page.keyboard.press('ControlOrMeta+z');expect(await editor.evaluate((e:any)=>e.editor.getJSON())).toEqual(before);await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor).toContainText(' EDIT');
});
