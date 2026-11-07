import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const { fixture, formEmfBytes } = createRequire(import.meta.url)('./hwp-metafile-fixture.cjs');
let directory: string, files: Record<string,string>;
test.beforeAll(async () => { directory=await fs.mkdtemp(path.join(os.tmpdir(),'word-studio-hwp-gdi-tests-'));files=await fixture(directory,formEmfBytes()); });
test.afterAll(async () => { await fs.rm(directory,{recursive:true}); });
for(const format of ['hwp','hwpx']) test(`${format}: real EMF form keeps its grey cells and top-aligned text through crop and two DOCX saves`,async({page})=>{
 await page.goto('/');
 const states=await page.evaluate(async({bytes,format})=>{
  const{HwpReaderSession}=await import('/src/hwp/reader.ts'),{importEditableHwp}=await import('/src/hwp/editing.ts');
  const{exportDocx,importFile}=await import('/src/document-io.ts'),{parseDocumentHtml,editorExtensions}=await import('/src/extensions.ts');
  const{generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{createPrintLayout}=await import('/src/print-layout.ts');
  const reader=new HwpReaderSession(),states=[];
  try{
   await reader.open(new File([new Uint8Array(bytes)],'form.'+format));let doc=importEditableHwp(await reader.editable(),'form.'+format);
   for(let save=0;save<3;save++){
    const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
    try{
     const pictures=[];
     for(const img of layout.element.querySelectorAll<HTMLImageElement>('.print-content img')){
      const c=document.createElement('canvas');c.width=img.naturalWidth;c.height=img.naturalHeight;const ctx=c.getContext('2d')!;ctx.fillStyle='white';ctx.fillRect(0,0,c.width,c.height);ctx.drawImage(img,0,0);
      const pixel=(x:number,y:number)=>[...ctx.getImageData(Math.floor(x*c.width),Math.floor(y*c.height),1,1).data];
      const raw=ctx.getImageData(0,0,c.width,Math.floor(c.height*.45)).data;let top=Infinity,black=0;
      for(let i=0;i<raw.length;i+=4)if(raw[i]<30&&raw[i+1]<30&&raw[i+2]<30){top=Math.min(top,Math.floor(i/4/c.width));black++;}
      pictures.push({grey:pixel(.9,.4),white:pixel(.9,.8),inkTop:top/c.height,black});
     }
     states.push({save,pages:layout.pages,text:layout.element.textContent,pictures});
    }finally{layout.dispose();}
    if(save<2){const loaded=await importFile(new File([await exportDocx(doc)],'again.docx'));doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};}
   }
   return states;
  }finally{reader.close();}
 },{bytes:[...await fs.readFile(files[format])],format});
 for(const state of states){
  expect(state.pages).toBe(1);expect(state.text).toContain('END');expect(state.pictures).toHaveLength(2);
  for(const p of state.pictures){for(const n of p.grey.slice(0,3))expect(n).toBeCloseTo(192,0);expect(p.grey[3]).toBe(255);expect(p.white).toEqual([255,255,255,255]);expect(p.black).toBeGreaterThan(50);expect(p.inkTop).toBeGreaterThan(.02);expect(p.inkTop).toBeLessThan(.2);}
 }
});

test('metafile rasterization keeps verified local clipping and patterns while rejecting external references',async({page})=>{
 const requests:string[]=[];page.on('request',r=>{if(r.url().includes('untrusted-emf.invalid'))requests.push(r.url())});await page.goto('/');
 const pixels=await page.evaluate(async()=>{
  const{prepareEditableHwp}=await import('/src/hwp/svg-images.ts');
  const svg='<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><defs><clipPath id="cell"><rect width="20" height="20"/></clipPath><pattern id="brush" patternUnits="userSpaceOnUse" width="10" height="10"><rect width="5" height="10" fill="red"/><rect x="5" width="5" height="10" fill="blue"/></pattern></defs><g clip-path="url(#cell)"><rect width="40" height="20" fill="url(#brush)"/></g><image href="https://untrusted-emf.invalid/picture.png"/><rect x="30" y="15" width="1" height="1" fill="url(https://untrusted-emf.invalid/paint)"/></svg>';
  const holder=document.createElement('div'),el=document.createElement('img');el.width=40;el.height=20;el.setAttribute('data-hwp-svg',svg);el.setAttribute('data-hwp-metafile','true');holder.append(el);
  const source=await prepareEditableHwp({html:holder.innerHTML,warnings:[],page:{size:'A4',orientation:'portrait',margin:20,header:'',footer:'',pageNumbers:false}});holder.innerHTML=source.html;const img=new Image();img.src=holder.querySelector('img')!.src;await img.decode();const c=document.createElement('canvas');c.width=40;c.height=20;const ctx=c.getContext('2d')!;ctx.fillStyle='white';ctx.fillRect(0,0,40,20);ctx.drawImage(img,0,0,40,20);return[2,7,22,35].map(x=>[...ctx.getImageData(x,5,1,1).data]);
 });
 expect(pixels).toEqual([[255,0,0,255],[0,0,255,255],[255,255,255,255],[255,255,255,255]]);expect(requests).toEqual([]);
});
