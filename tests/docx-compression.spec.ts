import { test, expect } from '@playwright/test';

test('DOCX repack preserves sections, ordered body and image bytes while retaining ZIP compression', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const expected = Array.from({length:300},(_,i)=>`Paragraph ${i}: ` + 'The whole document and its image must survive unchanged. '.repeat(3));
    const marker = `<div data-section-page="${JSON.stringify({...defaultPage,orientation:'landscape'}).replaceAll('"','&quot;')}"></div>`;
    const html = expected.map((text,i)=>(i===150?marker:'')+`<p>${text}</p>`).join('') + `<p><img src="data:image/png;base64,${png}" width="1" height="1"></p>`;
    const loaded = await importFile(new File([html],'compression.html'));
    let doc = {...loaded.document, content:parseDocumentHtml(loaded.html!)}, states=[];
    for(let save=0;save<2;save++) {
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const source = await zip.file('word/document.xml')!.async('string');
      const xml = new DOMParser().parseFromString(source,'application/xml'),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const body = [...xml.getElementsByTagNameNS(W,'p')].map(p=>[...p.getElementsByTagNameNS(W,'t')].map(t=>t.textContent).join('')).filter(Boolean);
      const media = Object.keys(zip.files).filter(name=>name.startsWith('word/media/')&&!zip.files[name].dir);
      states.push({bytes:blob.size,xmlBytes:new TextEncoder().encode(source).length,body,sections:xml.getElementsByTagNameNS(W,'sectPr').length,images:await Promise.all(media.map(name=>zip.file(name)!.async('base64')))});
      const reopened=await importFile(new File([blob],'again.docx'));doc={...reopened.document,content:parseDocumentHtml(reopened.html!)};
    }
    return {expected,png,states};
  });
  for(const state of result.states) {
    expect(state.body).toEqual(result.expected);expect(state.sections).toBe(2);expect(state.images).toEqual([result.png]);
    expect(state.bytes).toBeLessThan(state.xmlBytes / 4);
  }
});
