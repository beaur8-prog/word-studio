import {test,expect,type TestInfo} from '@playwright/test';
import fs from 'node:fs/promises';
import {Document,Packer,Paragraph,Table,TableRow,TableCell,LevelFormat,HeadingLevel,Header,Footer} from 'docx';
import JSZip from 'jszip';
async function writeDocx(testInfo: TestInfo, name: string, bytes: number[]) {
  const file = testInfo.outputPath(name);
  await fs.writeFile(file, Buffer.from(bytes));
}


async function fixture() {
  return Packer.toBuffer(new Document({
    styles:{paragraphStyles:[
      {id:'Context',name:'Context',paragraph:{contextualSpacing:true,spacing:{before:240,after:200}}},
      {id:'Derived',name:'Derived',basedOn:'Context'},
      {id:'Other',name:'Other',paragraph:{spacing:{before:240,after:200}}},
      {id:'Unflagged',name:'Unflagged',paragraph:{spacing:{before:240,after:200}}},
    ]},
    sections:[{children:[
      new Paragraph({text:'INHERITED',style:'Derived'}),
      new Paragraph({text:'DIRECT-FALSE',style:'Derived',contextualSpacing:false}),
      new Paragraph({text:'DIRECT-TRUE',style:'Other',contextualSpacing:true}),
      new Paragraph({text:'UNSPECIFIED',style:'Other'}),
      new Paragraph({text:'NORMAL-DIRECT',style:'Normal',contextualSpacing:true,spacing:{before:240,after:200}}),
      new Paragraph({text:'DIFFERENT-NO-POLICY',style:'Unflagged'}),
      new Table({rows:[new TableRow({children:[new TableCell({children:[
        new Paragraph({text:'CELL-ONE',style:'Context'}),new Paragraph({text:'CELL-TWO',style:'Context'}),
      ]})]})]}),
    ]}],
  }));
}

test('source style identity and inherited/direct contextual spacing survive two standard DOCX saves',async({page},testInfo)=>{
  await page.goto('/');
  const result=await page.evaluate(async input=>{
    const{importFile,exportDocx}=await import('/src/document-io.ts');
    const{parseDocumentHtml}=await import('/src/extensions.ts');
    const{default:JSZip}=await import('/node_modules/.vite/deps/jszip.js');
    let file=new File([new Uint8Array(input)],'context.docx');const stages=[],saves=[];
    for(let i=0;i<3;i++){
      const opened=await importFile(file),model={...opened.document,content:parseDocumentHtml(opened.html!)};
      const rows:any[]=[];const walk=(n:any)=>{if(['paragraph','heading'].includes(n.type))rows.push({text:(n.content||[]).map((x:any)=>x.text||'').join(''),style:n.attrs?.wordParagraphStyle,enabled:n.attrs?.wordContextualSpacing,before:n.attrs?.spaceBefore,after:n.attrs?.spaceAfter});for(const c of n.content||[])walk(c)};walk(model.content);
      stages.push(rows);
      if(i<2){const bytes=[...new Uint8Array(await(await exportDocx(model)).arrayBuffer())];saves.push(bytes);file=new File([new Uint8Array(bytes)],'context.docx')}
    }
    return{stages,saves};
  },[...await fixture()]);
  for(const rows of result.stages){
    expect(rows.filter(r=>r.text).map(r=>[r.text,r.enabled])).toEqual([
      ['INHERITED',true],['DIRECT-FALSE',false],['DIRECT-TRUE',true],['UNSPECIFIED',null],['NORMAL-DIRECT',true],['DIFFERENT-NO-POLICY',null],['CELL-ONE',true],['CELL-TWO',true],
    ]);
    expect(rows[0].style).toBeTruthy();expect(rows[0].style).toBe(rows[1].style);
    expect(rows[2].style).toBe(rows[3].style);expect(rows[0].style).not.toBe(rows[2].style);
    // An unflagged custom style must not merge into a flagged Normal style.
    expect(rows[4].style).not.toBe(rows[5].style);
    for(const r of rows.filter(r=>r.text)) {expect(r.before).toBe('16px');expect(r.after).toBe('13.3333px')}
  }
  for(let i=0;i<result.saves.length;i++)await writeDocx(testInfo,`contextual-save${i+1}.docx`,result.saves[i]);
});

test('same-style one-sided spacing subtracts from the collapsed gap and does not zero the other paragraph',async({page})=>{
  const input=await Packer.toBuffer(new Document({styles:{paragraphStyles:[{id:'Same',name:'Same'}]},sections:[{children:[
    new Paragraph({text:'FIRST',style:'Same',contextualSpacing:true,spacing:{before:0,after:200,line:240}}),
    new Paragraph({text:'SECOND',style:'Same',contextualSpacing:false,spacing:{before:240,after:0,line:240}}),
  ]}]}));
  await page.goto('/');
  const result=await page.evaluate(async input=>{
    const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js');
    const{importFile}=await import('/src/document-io.ts');const{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts');
    const{createPrintLayout}=await import('/src/print-layout.ts');
    const opened=await importFile(new File([new Uint8Array(input)],'gaps.docx'));
    const mount=document.createElement('div');mount.style.cssText='position:absolute;left:100px;top:100px;width:600px';document.body.append(mount);
    const editor=new Editor({element:mount,extensions:editorExtensions,content:parseDocumentHtml(opened.html!)});
    await document.fonts.ready;await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
    const gap=(root:HTMLElement)=>{const p=[...root.querySelectorAll('p')].filter(p=>['FIRST','SECOND'].includes(p.textContent||''));return p[1].getBoundingClientRect().top-p[0].getBoundingClientRect().bottom};
    const edited=gap(editor.view.dom);const model=editor.getJSON();
    const printed=await createPrintLayout(editor.getHTML(),opened.document.page);
    try{return{edited,printed:gap(printed.element),model}}
    finally{printed.dispose();editor.destroy();mount.remove()}
  },[...input]);
  expect(result.edited).toBeCloseTo(2*4/3,1);expect(result.printed).toBeCloseTo(2*4/3,1);
  expect(result.model.content?.map((p:any)=>p.attrs.spaceAfter)).toEqual(['13.3333px','0px']);
});


test('conditional gaps work across list items and retain different-style and container boundaries',async({page})=>{
  await page.goto('/');
  const result=await page.evaluate(async()=>{
    const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js');
    const{editorExtensions}=await import('/src/extensions.ts');
    const{createPrintLayout}=await import('/src/print-layout.ts');const{defaultPage}=await import('/src/types.ts');
    const p=(text:string,style:string,enabled:boolean|null,before:number,after:number)=>({type:'paragraph',attrs:{wordParagraphStyle:style,wordContextualSpacing:enabled,spaceBefore:before+'px',spaceAfter:after+'px',lineSpacing:'1'},content:[{type:'text',text}]});
    const content={type:'doc',content:[
      p('BODY','Same',true,0,10),
      {type:'bulletList',content:[
        {type:'listItem',content:[p('LIST-ONE','Same',false,12,3)]},
        {type:'listItem',content:[p('LIST-TWO','Same',true,8,4),{type:'bulletList',content:[{type:'listItem',content:[p('NESTED-OTHER','Other',true,5,6)]}]}]},
      ]},
      {type:'table',content:[{type:'tableRow',content:[
        {type:'tableCell',content:[p('CELL-ONE','Same',true,12,10)]},
        {type:'tableCell',content:[p('CELL-TWO','Same',true,12,10)]},
      ]}]},
      p('AFTER-TABLE','Same',true,12,10),
    ]};
    const mount=document.createElement('div');mount.style.cssText='position:absolute;left:100px;top:100px;width:600px';document.body.append(mount);
    const editor=new Editor({element:mount,extensions:editorExtensions,content});
    const snapshot=(root:HTMLElement)=>Object.fromEntries([...root.querySelectorAll('p')].map(p=>[p.textContent,{before:parseFloat(getComputedStyle(p).marginTop),after:parseFloat(getComputedStyle(p).marginBottom)}]));
    const editing=snapshot(editor.view.dom),model=editor.getJSON();
    const printed=await createPrintLayout(editor.getHTML(),defaultPage);
    try{return{editing,printing:snapshot(printed.element),model}}finally{printed.dispose();editor.destroy();mount.remove()}
  });
  for(const rows of [result.editing,result.printing]){
    expect(rows['BODY'].after).toBe(0);expect(rows['LIST-ONE']).toEqual({before:2,after:0});
    expect(rows['LIST-TWO']).toEqual({before:0,after:4});
    expect(rows['NESTED-OTHER']).toEqual({before:5,after:6});
    expect(rows['CELL-ONE']).toEqual({before:12,after:10});expect(rows['CELL-TWO']).toEqual({before:12,after:10});
    expect(rows['AFTER-TABLE']).toEqual({before:12,after:10});
  }
  expect(result.model.content?.[0].attrs?.spaceAfter).toBe('10px');
});

test('changing the neighboring style, splitting and joining update view-only gaps through undo and redo',async({page})=>{
  await page.goto('/');
  const result=await page.evaluate(async()=>{
    const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js');const{editorExtensions}=await import('/src/extensions.ts');
    const{closeHistory}=await import('/node_modules/.vite/deps/@tiptap_pm_history.js');
    const p=(text:string)=>({type:'paragraph',attrs:{wordParagraphStyle:'Same',wordContextualSpacing:true,spaceBefore:'8px',spaceAfter:'10px'},content:[{type:'text',text}]});
    const mount=document.createElement('div');document.body.append(mount);
    const editor=new Editor({element:mount,extensions:editorExtensions,content:{type:'doc',content:[p('FIRST'),p('SECOND')]}});
    const snapshot=()=>({text:editor.getText(),gaps:[...editor.view.dom.querySelectorAll('p')].map(p=>[getComputedStyle(p).marginTop,getComputedStyle(p).marginBottom]),attrs:editor.getJSON().content!.map(p=>p.attrs)});
    const initial=snapshot();const second=editor.state.doc.child(0).nodeSize;
    editor.view.dispatch(closeHistory(editor.state.tr));editor.commands.setTextSelection(second+1);editor.commands.updateAttributes('paragraph',{wordParagraphStyle:'Other'});
    const different=snapshot();editor.commands.undo();const undone=snapshot();editor.commands.redo();const redone=snapshot();editor.commands.undo();
    editor.view.dispatch(closeHistory(editor.state.tr));editor.commands.setTextSelection(second+4);editor.commands.splitBlock();const split=snapshot();
    editor.view.dispatch(closeHistory(editor.state.tr));editor.commands.joinBackward();const joined=snapshot();
    editor.commands.undo();const undoJoin=snapshot();
    const html=editor.getHTML();editor.destroy();mount.remove();return{initial,different,undone,redone,split,joined,undoJoin,html};
  });
  expect(result.initial.gaps).toEqual([['8px','0px'],['0px','10px']]);
  expect(result.different.gaps).toEqual([['8px','10px'],['8px','10px']]);
  expect(result.undone).toEqual(result.initial);expect(result.redone).toEqual(result.different);
  expect(result.split.gaps).toEqual([['8px','0px'],['0px','0px'],['0px','10px']]);
  expect(result.joined.gaps).toEqual(result.initial.gaps);expect(result.undoJoin.gaps).toEqual(result.split.gaps);
  for(const stage of [result.initial,result.different,result.split,result.joined])for(const attrs of stage.attrs)expect([attrs?.spaceBefore,attrs?.spaceAfter]).toEqual(['8px','10px']);
  expect(result.html).toContain('margin-bottom: 10px');
});

test('standard heading semantics and a single pStyle survive conditional spacing exports',async({page},testInfo)=>{
  const input=await Packer.toBuffer(new Document({styles:{paragraphStyles:[{id:'Subtitle',name:'Subtitle'}]},sections:[{children:[
    new Paragraph({text:'HEADING-ONE',heading:HeadingLevel.HEADING_1,contextualSpacing:true}),
    new Paragraph({text:'HEADING-TWO',heading:HeadingLevel.HEADING_1,contextualSpacing:false}),
    new Paragraph({text:'SUBTITLE-NO-POLICY',style:'Subtitle'}),
    new Paragraph({text:'BODY-UNCHANGED'}),
  ]}]}));
  await page.goto('/');
  const result=await page.evaluate(async input=>{
    const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts'),{default:JSZip}=await import('/node_modules/.vite/deps/jszip.js');
    const states=[],saves=[];let file=new File([new Uint8Array(input)],'headings.docx');
    for(let i=0;i<3;i++){const opened=await importFile(file),doc={...opened.document,content:parseDocumentHtml(opened.html!)};states.push(doc.content);
      if(i<2){const b=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(b);const zip=await JSZip.loadAsync(new Uint8Array(b));const xml=new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'),'application/xml');
        const styles=new DOMParser().parseFromString(await zip.file('word/styles.xml')!.async('string'),'application/xml'),defined=new Set([...styles.getElementsByTagNameNS('*','style')].map(s=>s.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','styleId')));
        for(const p of [...xml.getElementsByTagNameNS('*','p')]){const refs=[...p.getElementsByTagNameNS('*','pStyle')];if(refs.length>1)throw Error('duplicate pStyle');for(const ref of refs)if(!defined.has(ref.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','val')))throw Error('missing paragraph style definition')}file=new File([new Uint8Array(b)],'again.docx')}}
    return{states,saves};
  },[...input]);
  for(const doc of result.states){expect(doc.content!.map((p:any)=>[p.type,p.attrs.level||null,p.attrs.wordContextualSpacing])).toEqual([['heading',1,true],['heading',1,false],['heading',2,null],['paragraph',null,null]])}
  for(let i=0;i<result.saves.length;i++)await writeDocx(testInfo,`contextual-heading-save${i+1}.docx`,result.saves[i]);
});

test('header and footer paragraph policies survive standard DOCX and use their own print flows',async({page},testInfo)=>{
  const source=await Packer.toBuffer(new Document({styles:{paragraphStyles:[{id:'Story',name:'Story',paragraph:{contextualSpacing:true,spacing:{before:240,after:200}}}]},sections:[{
    headers:{default:new Header({children:[new Paragraph({text:'HEADER-ONE',style:'Story'}),new Paragraph({text:'HEADER-TWO',style:'Story'})]})},
    footers:{default:new Footer({children:[new Paragraph({text:'FOOTER-ONE',style:'Story'}),new Paragraph({text:'FOOTER-TWO',style:'Story'})]})},
    children:[new Paragraph({text:'BODY-STORY-ISOLATION',style:'Story'})],
  }]}));
  await page.goto('/');
  const result=await page.evaluate(async bytes=>{
    const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts'),{createPrintLayout}=await import('/src/print-layout.ts');
    const opened=await importFile(new File([new Uint8Array(bytes)],'stories.docx')),doc={...opened.document,content:parseDocumentHtml(opened.html!)};
    const states=[],saves=[];let model=doc;
    for(let i=0;i<3;i++){const layout=await createPrintLayout('<p>PRINT-BODY</p>',model.page);try{states.push(['.print-header','.print-footer .document-furniture'].map(selector=>[...layout.element.querySelector(selector)!.querySelectorAll('p')].map(p=>({text:p.textContent,enabled:(p as HTMLElement).dataset.wordContextualSpacing,before:parseFloat(getComputedStyle(p).marginTop),after:parseFloat(getComputedStyle(p).marginBottom)}))))}finally{layout.dispose()}
      if(i<2){const b=[...new Uint8Array(await(await exportDocx(model)).arrayBuffer())];saves.push(b);const next=await importFile(new File([new Uint8Array(b)],'again.docx'));model={...next.document,content:parseDocumentHtml(next.html!)}}}
    return{states,saves};
  },[...source]);
  for(const stage of result.states){for(const story of stage){expect(story).toHaveLength(2);expect(story[0].after).toBe(0);expect(story[1].before).toBe(0);for(const p of story)expect(p.enabled).toBe('true')}}
  for(let i=0;i<result.saves.length;i++)await writeDocx(testInfo,`contextual-story-save${i+1}.docx`,result.saves[i]);
});


test('document defaults, default style true, numbering and explicit false resolve independently',async({page},testInfo)=>{
  const source=await Packer.toBuffer(new Document({
    styles:{default:{document:{paragraph:{contextualSpacing:true,spacing:{before:60,after:100}}}},paragraphStyles:[
      {id:'Disabled',name:'Disabled',paragraph:{contextualSpacing:false}},
    ]},
    numbering:{config:[{reference:'One',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',style:{paragraph:{contextualSpacing:false}}}]}]},
    sections:[{children:[new Paragraph({text:'DEFAULT-TRUE'}),new Paragraph({text:'STYLE-FALSE',style:'Disabled'}),
      new Paragraph({text:'NUMBERING-FALSE',numbering:{reference:'One',level:0}}),
      new Paragraph({text:'DIRECT-ON',numbering:{reference:'One',level:0},contextualSpacing:true}),
    ]}],
  }));
  const zip=await JSZip.loadAsync(source);let styles=await zip.file('word/styles.xml')!.async('string');
  styles=styles.replace(/<w:style[^>]*w:styleId="Normal"[^>]*>/,s=>s.replace('w:styleId="Normal"','w:styleId="ActualDefault"').replace('w:default="1"','w:default="true"'));
  zip.file('word/styles.xml',styles);
  const bytes=await zip.generateAsync({type:'uint8array'});
  await page.goto('/');
  const result=await page.evaluate(async bytes=>{
    const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');let file=new File([new Uint8Array(bytes)],'defaults.docx');const stages=[],saves=[];
    for(let i=0;i<3;i++){const opened=await importFile(file),doc={...opened.document,content:parseDocumentHtml(opened.html!)};const rows:any[]=[];const walk=(n:any)=>{if(n.type==='paragraph')rows.push({text:(n.content||[]).map((c:any)=>c.text||'').join(''),enabled:n.attrs.wordContextualSpacing,style:n.attrs.wordParagraphStyle});for(const c of n.content||[])walk(c)};walk(doc.content);stages.push(rows);
      if(i<2){const b=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(b);file=new File([new Uint8Array(b)],'again.docx')}}return{stages,saves};
  },[...bytes]);
  expect(result.stages[0][0].style).toBe('ActualDefault');
  for(const rows of result.stages)expect(rows.map(p=>[p.text,p.enabled])).toEqual([['DEFAULT-TRUE',true],['STYLE-FALSE',false],['NUMBERING-FALSE',false],['DIRECT-ON',true]]);
  for(let i=0;i<result.saves.length;i++)await writeDocx(testInfo,`contextual-numbering-save${i+1}.docx`,result.saves[i]);
});

test('invalid contextual metadata is rejected and large valid paragraph spaces remain conditional',async({page})=>{
  await page.goto('/');
  const result=await page.evaluate(async()=>{
    const{validateDocument}=await import('/src/document-io.ts'),{defaultPage}=await import('/src/types.ts');
    const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{editorExtensions}=await import('/src/extensions.ts');
    const make=(attrs:any)=>({format:'word-studio',version:1,title:'validation',updatedAt:new Date().toISOString(),page:defaultPage,comments:[],content:{type:'doc',content:[{type:'paragraph',attrs}]}});
    const invalid=[{wordContextualSpacing:'true'},{wordParagraphStyle:42},{wordParagraphStyle:'Bad\u0000Style'}].map(attrs=>{try{validateDocument(make(attrs));return false}catch{return true}});
    const mount=document.createElement('div');document.body.append(mount);
    const p=(text:string)=>({type:'paragraph',attrs:{wordParagraphStyle:'Same',wordContextualSpacing:true,spaceBefore:'3000px',spaceAfter:'4000px'},content:[{type:'text',text}]});
    const editor=new Editor({element:mount,extensions:editorExtensions,content:{type:'doc',content:[p('LONG-SPACE-ONE'),p('LONG-SPACE-TWO')]}});
    const gaps=[...editor.view.dom.querySelectorAll('p')].map(p=>[getComputedStyle(p).marginTop,getComputedStyle(p).marginBottom]);editor.destroy();mount.remove();return{invalid,gaps};
  });
  expect(result.invalid).toEqual([true,true,true]);expect(result.gaps).toEqual([['3000px','0px'],['0px','4000px']]);
});


test('spacing on the following paragraph subtracts only that side and repeated print measurement is stable',async({page})=>{
  await page.goto('/');
  const result=await page.evaluate(async()=>{
    const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{editorExtensions}=await import('/src/extensions.ts');
    const{layoutWordContextualSpacing}=await import('/src/word-contextual-spacing.ts');
    const p=(text:string,enabled:boolean,before:number,after:number)=>({type:'paragraph',attrs:{wordParagraphStyle:'Same',wordContextualSpacing:enabled,spaceBefore:before+'pt',spaceAfter:after+'pt'},content:[{type:'text',text}]});
    const mount=document.createElement('div');document.body.append(mount);
    const editor=new Editor({element:mount,extensions:editorExtensions,content:{type:'doc',content:[p('FIRST',false,0,12),p('SECOND',true,10,0)]}});
    const snapshot=(root:HTMLElement)=>[...root.querySelectorAll('p')].map(p=>[parseFloat(getComputedStyle(p).marginTop),parseFloat(getComputedStyle(p).marginBottom)]);
    const editing=snapshot(editor.view.dom);const printed=document.createElement('div');printed.className='tiptap';printed.innerHTML=editor.getHTML();document.body.append(printed);
    layoutWordContextualSpacing(printed);const once=snapshot(printed);layoutWordContextualSpacing(printed);const twice=snapshot(printed);
    const original=editor.getJSON();editor.destroy();mount.remove();printed.remove();return{editing,once,twice,original};
  });
  for(const rows of [result.editing,result.once,result.twice]){expect(rows[0][1]).toBeCloseTo(2*4/3,3);expect(rows[1][0]).toBe(0)}
  expect(result.original.content!.map((p:any)=>[p.attrs.spaceBefore,p.attrs.spaceAfter])).toEqual([['0pt','12pt'],['10pt','0pt']]);
});


test('heading-to-body edits retire the source policy, while undo and whole-document opens preserve it',async({page})=>{
  await page.goto('/');
  const result=await page.evaluate(async()=>{
    const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{editorExtensions}=await import('/src/extensions.ts');
    const mount=document.createElement('div');document.body.append(mount);
    const heading={type:'heading',attrs:{level:1,wordParagraphStyle:'SourceHeading',wordParagraphStyleKind:'heading1',wordContextualSpacing:true,spaceAfter:'10px'},content:[{type:'text',text:'TITLE'}]};
    const editor=new Editor({element:mount,extensions:editorExtensions,content:{type:'doc',content:[heading,{type:'paragraph',content:[{type:'text',text:'BODY'}]}]}});
    const initial=editor.getJSON();editor.commands.setTextSelection(6);editor.commands.splitBlock();const split=editor.getJSON();editor.commands.undo();const undone=editor.getJSON();editor.commands.redo();const redone=editor.getJSON();
    editor.commands.setContent({type:'doc',content:[{...heading,type:'paragraph',attrs:{...heading.attrs,wordParagraphStyleKind:'paragraph'}}]});const reopened=editor.getJSON();
    editor.commands.setTextSelection(2);editor.commands.setHeading({level:2});const changed=editor.getJSON();
    editor.destroy();mount.remove();return{initial,split,undone,redone,reopened,changed};
  });
  expect(result.split.content!.map((n:any)=>[n.type,n.attrs.wordParagraphStyle,n.attrs.wordContextualSpacing])).toEqual([['heading','SourceHeading',true],['paragraph',null,null],['paragraph',null,null]]);
  expect(result.undone).toEqual(result.initial);expect(result.redone).toEqual(result.split);
  expect(result.reopened.content![0].attrs?.wordParagraphStyle).toBe('SourceHeading');
  expect(result.changed.content![0].attrs?.wordParagraphStyle).toBeNull();expect(result.changed.content![0].attrs?.wordContextualSpacing).toBeNull();
});


test('the same Title style remains a body heading and an ordinary textbox paragraph through two saves',async({page},testInfo)=>{
  const source=await Packer.toBuffer(new Document({sections:[{children:[new Paragraph({text:'BODY TITLE',heading:HeadingLevel.TITLE,contextualSpacing:true}),new Paragraph({text:'BODY END'})]}]}));
  const zip=await JSZip.loadAsync(source);let xml=await zip.file('word/document.xml')!.async('string');
  const box=`<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="TitleBox" style="position:absolute;margin-left:20px;margin-top:100px;width:240px;height:80px;mso-position-horizontal-relative:margin;mso-position-vertical-relative:margin" fillcolor="#EEEEEE" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr><w:pStyle w:val="Title"/><w:contextualSpacing/></w:pPr><w:r><w:t>BOX TITLE</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
  xml=xml.replace('<w:sectPr>',box+'<w:p><w:r><w:t>AFTER BOX</w:t></w:r></w:p><w:sectPr>');zip.file('word/document.xml',xml);const bytes=await zip.generateAsync({type:'uint8array'});
  await page.goto('/');
  const result=await page.evaluate(async bytes=>{
    const{importFile,exportDocx}=await import('/src/document-io.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');const{Editor}=await import('/node_modules/.vite/deps/@tiptap_react.js');const{editorExtensions}=await import('/src/extensions.ts');const states=[],saves=[];let editing:any=null;let file=new File([new Uint8Array(bytes)],'mixed-title.docx');
    const text=(n:any):string=>(n.text||'')+(n.content||[]).map(text).join('');
    for(let i=0;i<3;i++){const opened=await importFile(file),doc={...opened.document,content:parseDocumentHtml(opened.html!)};const box=doc.content.content!.find(n=>n.type==='documentTextBox')!;states.push({body:doc.content.content!.find(n=>text(n)==='BODY TITLE'),box:box.content![0]});if(i===0){const mount=document.createElement('div');document.body.append(mount);const editor=new Editor({element:mount,extensions:editorExtensions,content:doc.content});editor.commands.setTextSelection(2);editor.commands.setParagraph();const changed=editor.getJSON();editor.commands.undo();editing={changedBody:changed.content!.find(n=>text(n)==='BODY TITLE'),unchangedBox:changed.content!.find(n=>n.type==='documentTextBox')!.content![0],undoExact:JSON.stringify(editor.getJSON())===JSON.stringify(doc.content)};editor.destroy();mount.remove()}if(i<2){const b=[...new Uint8Array(await(await exportDocx(doc)).arrayBuffer())];saves.push(b);file=new File([new Uint8Array(b)],'again.docx')}}return{states,saves,editing};
  },[...bytes]);
  for(const stage of result.states){expect(stage.body!.type).toBe('heading');expect(stage.body!.attrs?.level).toBe(1);expect(stage.box.type).toBe('paragraph');expect(stage.body!.attrs?.wordParagraphStyle).toBe(stage.box.attrs?.wordParagraphStyle);expect(stage.body!.attrs?.wordParagraphStyleName).toBe('Title');expect(stage.box.attrs?.wordParagraphStyleName).toBe('Title')}
  expect(result.editing.changedBody.type).toBe('paragraph');expect(result.editing.changedBody.attrs.wordParagraphStyle).toBeNull();expect(result.editing.changedBody.attrs.wordContextualSpacing).toBeNull();expect(result.editing.unchangedBox.attrs.wordParagraphStyle).toBe('Title');expect(result.editing.undoExact).toBe(true);
  for(let i=0;i<result.saves.length;i++)await writeDocx(testInfo,`contextual-title-save${i+1}.docx`,result.saves[i]);
});
