import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-header-fixture.cjs');
let directory: string, files: Record<string,string>;
test.beforeAll(async () => { directory=await fs.mkdtemp(path.resolve('.cache/hwp-header-tests-')); files=await fixture(directory); });
test.afterAll(async () => { await fs.rm(directory,{recursive:true}); });
for(const format of ['hwp','hwpx']) test(`${format}: rich odd/even headers, hidden first page and live page fields survive two DOCX saves`,async({page})=>{
  await page.goto('/');
  const result=await page.evaluate(async({bytes,format})=>{
    const {HwpReaderSession}=await import('/src/hwp/reader.ts'),{importEditableHwp}=await import('/src/hwp/editing.ts');
    const {exportDocx,importFile}=await import('/src/document-io.ts'),{createPrintLayout}=await import('/src/print-layout.ts');
    const {generateHTML}=await import('/node_modules/.vite/deps/@tiptap_react.js'),{editorExtensions,parseDocumentHtml}=await import('/src/extensions.ts');
    const session=new HwpReaderSession();
    try{
      await session.open(new File([new Uint8Array(bytes)],'headers.'+format));
      let doc=importEditableHwp(await session.editable(),'headers.'+format);const first=doc,states=[];
      for(let save=0;save<3;save++){
        const layout=await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
        try{states.push({pages:layout.pages,headers:[...layout.element.querySelectorAll('.print-header')].map(h=>({text:h.textContent,alts:[...h.querySelectorAll('img')].map(i=>i.alt),raw:/data-hwp-object/.test(h.innerHTML)}))});}finally{layout.dispose()}
        if(save<2){const again=await importFile(new File([await exportDocx(doc)],'again.docx'));doc={...again.document,content:parseDocumentHtml(again.html!)};}
      }
      return {first,states};
    }finally{session.close()}
  },{bytes:[...await fs.readFile(files[format])],format});
  for(const state of result.states){
    expect(state.pages).toBe(3);expect(state.headers[0].alts).toEqual([]);expect(state.headers[0].text).toBe('');
    expect(state.headers[1].alts.join('')).toContain('EVEN HEADER');expect(state.headers[1].text?.trim()).toBe('2');
    expect(state.headers[2].alts.join('')).toContain('ODD HEADER');expect(state.headers[2].text?.trim()).toBe('3');
    expect(state.headers.some(h=>h.raw)).toBe(false);
  }
  await page.locator('input[type=file]').first().setInputFiles({name:'headers.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(result.first))});
  await expect.poll(()=>page.locator('.paper').count()).toBe(3);
  await expect(page.locator('.paper').nth(1).locator('.paper-header img')).toHaveAttribute('alt',/EVEN HEADER/);
  await expect(page.locator('.paper').nth(2).locator('.paper-header [data-page-number]')).toHaveText('3');
});

test('a new-page section changes stories on the following sheet; continuous sections update automatic numbering', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { parseDocumentHtml } = await import('/src/extensions.ts');
    const { importFile } = await import('/src/document-io.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const original = { ...defaultPage, header: 'FIRST STORY', pageNumbers: true };
    const next = { ...original, header: 'NEXT STORY', breakType: 'nextPage', pageNumbers: false };
    const marker = (settings: object) => `<div data-section-page="${JSON.stringify(settings).replace(/"/g, '&quot;')}"></div>`;
    const html = '<p>BEFORE</p>' + marker(next) + '<p>AFTER</p>';
    const loaded = await importFile(new File([html], 'stories.html'));
    const layout = await createPrintLayout(html, original);
    const headers = [...layout.element.querySelectorAll('.print-header')].map(h => h.textContent);
    layout.dispose();
    const continuous = await createPrintLayout('<p>BODY</p>' + marker({ ...next, breakType: 'continuous' }), original);
    const numbers = continuous.element.querySelectorAll('.print-automatic-number').length;
    continuous.dispose();
    return { doc: { ...loaded.document, page: original, content: parseDocumentHtml(html) }, headers, numbers };
  });
  expect(result.headers).toEqual(['FIRST STORY', 'NEXT STORY']); expect(result.numbers).toBe(0);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'stories.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  await expect(page.locator('.paper-header')).toHaveText(['FIRST STORY', 'NEXT STORY']);
});
