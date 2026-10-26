import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-crop-fixture.cjs');
let directory: string, files: Record<string,string>;
test.beforeAll(async () => { directory=await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-hwp-crop-tests-')); files=await fixture(directory); });
test.afterAll(async () => { await fs.rm(directory,{recursive:true}); });
for (const format of ['hwp','hwpx']) test(`${format}: cropped body/cell pictures retain visible pixels and image-only paragraph alignment through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession(), states = [];
    try {
      await session.open(new File([new Uint8Array(bytes)],'crop.'+format)); let doc=importEditableHwp(await session.editable(),'crop.'+format);
      for(let save=0;save<3;save++) {
        const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
        try {
          const images=[...layout.element.querySelectorAll<HTMLImageElement>('.print-content img')], colors=[];
          for(const img of images){const canvas=document.createElement('canvas');canvas.width=3;canvas.height=1;canvas.getContext('2d')!.drawImage(img,0,0,3,1);colors.push([...canvas.getContext('2d')!.getImageData(0,0,3,1).data]);}
          const cell=layout.element.querySelector('td')!, img=cell.querySelector('img')!, a=cell.getBoundingClientRect(),b=img.getBoundingClientRect();
          states.push({save,paragraph:img.closest('p')?.style.cssText,paragraphs:cell.querySelectorAll('p').length,colors,centerError:Math.abs((a.left+a.right-b.left-b.right)/2),pages:layout.pages,text:layout.element.textContent});
        }finally{layout.dispose()}
        if(save<2){const loaded=await importFile(new File([await exportDocx(doc)],'again.docx'));doc={...loaded.document,content:parseDocumentHtml(loaded.html!)};}
      }
      return states;
    }finally{session.close()}
  },{bytes:[...await fs.readFile(files[format])],format});
  for(const state of states){
    expect(state.pages).toBe(1);expect(state.paragraphs).toBe(3);expect(state.text).toContain('END');expect(state.colors).toHaveLength(2);expect(state.centerError, JSON.stringify({save:state.save,paragraph:state.paragraph})).toBeLessThan(1);
    for(const pixels of state.colors)for(let i=0;i<pixels.length;i+=4){expect(pixels[i]).toBeGreaterThan(245);expect(pixels[i+2]).toBeLessThan(10);expect(pixels[i+3]).toBe(255);}
  }
});
