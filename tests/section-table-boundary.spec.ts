import { test, expect } from '@playwright/test';

for (const ending of ['table', 'blank'] as const) test(`DOCX ${ending} section ending preserves body paragraphs through two saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ending => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const paper = { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false };
    const section = `<div data-section-page="${JSON.stringify(paper).replaceAll('"', '&quot;')}"></div>`;
    const html = `<table style="margin:0"><tr style="height:${ending === 'table' ? 600 : 520}px"><td style="padding:0"><p style="line-height:20px;margin:0">TABLE BODY</p></td></tr></table>` +
      (ending === 'blank' ? '<p style="line-height:20px;margin:0"></p>' : '') + section + '<p style="line-height:20px;margin:0">AFTER SECTION</p>';
    const loaded = await importFile(new File([html], 'boundary.html'));
    let doc = { ...loaded.document, page: paper, content: parseDocumentHtml(html) };
    const states = [];
    for (let save = 0; save < 3; save++) {
      const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
      const layout = await createPrintLayout(root.innerHTML, doc.page);
      states.push({ paragraphs: [...root.querySelectorAll('p')].map(p => p.textContent), sections: root.querySelectorAll('[data-section-page]').length, pages: layout.pages }); layout.dispose();
      if (save === 2) break;
      const again = await importFile(new File([await exportDocx(doc)], 'boundary.docx'));
      doc = { ...again.document, content: parseDocumentHtml(again.html!) };
    }
    return { states, doc };
  }, ending);
  for (const state of result.states) {
    expect(state.paragraphs).toEqual(ending === 'blank' ? ['TABLE BODY', '', 'AFTER SECTION'] : ['TABLE BODY', 'AFTER SECTION']);
    expect(state.sections).toBe(1); expect(state.pages).toBe(2);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'boundary.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  await expect(page.getByRole('textbox', { name: '문서 편집 영역' })).toContainText('AFTER SECTION');
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('2');
  }
});

test('text added externally to a generated DOCX section boundary remains editable', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const html = '<table><tr><td><p>TABLE</p></td></tr></table>' + `<div data-section-page="${JSON.stringify(defaultPage).replaceAll('"', '&quot;')}"></div><p>AFTER</p>`;
    const loaded = await importFile(new File([html], 'boundary.html'));
    const zip = await JSZip.loadAsync(await (await exportDocx({ ...loaded.document, content: parseDocumentHtml(html) })).arrayBuffer());
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const carrier = [...xml.getElementsByTagNameNS(W, 'p')].find(p => p.getElementsByTagNameNS(W, 'sectPr').length)!;
    const run = xml.createElementNS(W, 'w:r'), text = xml.createElementNS(W, 'w:t'); text.textContent = 'EXTERNAL TEXT'; run.append(text); carrier.append(run);
    zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
    const again = await importFile(new File([await zip.generateAsync({ type: 'blob' })], 'edited.docx'));
    const template = document.createElement('template'); template.innerHTML = again.html!;
    return { text: template.content.textContent, sections: template.content.querySelectorAll('[data-section-page]').length };
  });
  expect(result.text).toContain('EXTERNAL TEXT'); expect(result.sections).toBe(1);
});
