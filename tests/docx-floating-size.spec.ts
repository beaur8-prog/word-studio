import { test, expect } from '@playwright/test';

for (const location of ['paper', 'cell']) test(`DOCX ${location} floating picture keeps its explicit size through two saves while inline pictures still fit`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async location => {
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const geometry = { horizontalPosition: { relative: 'page', offset: 37 * 9525 }, verticalPosition: { relative: 'page', offset: 10 * 9525 }, wrap: { type: 0 }, behindDocument: true, layoutInCell: true };
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const width = location === 'paper' ? 4.013 : 93.467, height = location === 'paper' ? 1137.853 : 88.44;
    const floating = `<img data-inline-image="true" data-image-floating="${JSON.stringify(geometry).replaceAll('"', '&quot;')}" src="data:image/png;base64,${png}" width="${width}" height="${height}" alt="FLOAT">`;
    const inline = `<img data-inline-image="true" src="data:image/png;base64,${png}" width="1200" height="1200" alt="INLINE">`;
    const html = location === 'paper' ? `<p>BEFORE${floating}</p><p>${inline}</p><p>AFTER</p>` : `<p>BEFORE</p><table style="width:43px"><colgroup><col style="width:43px"></colgroup><tr><td colwidth="43" style="width:43px;padding:0"><p>${floating}</p><p>${inline}</p></td></tr></table><p>AFTER</p>`;
    const loaded = await importFile(new File([html], 'float.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(html) };
    const states = [];
    for (let save = 0; save < 3; save++) {
      const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
      const img = root.querySelector<HTMLImageElement>('img[alt=FLOAT]')!, flow = root.querySelector<HTMLImageElement>('img[alt=INLINE]')!;
      const layout = await createPrintLayout(root.innerHTML, doc.page);
      await Promise.all([...layout.element.querySelectorAll<HTMLImageElement>('img')].map(img => img.decode()));
      const painted = layout.element.querySelector<HTMLImageElement>('img[alt=FLOAT]')!.getBoundingClientRect();
      states.push({ width: Number(img.getAttribute('width')), height: Number(img.getAttribute('height')), paintedWidth: painted.width, paintedHeight: painted.height,
        floating: JSON.parse(img.dataset.imageFloating!), inlineWidth: Number(flow.getAttribute('width')), inlineHeight: Number(flow.getAttribute('height')) }); layout.dispose();
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return states;
  }, location);
  for (const state of states) {
    for (const field of ['width', 'height', 'paintedWidth', 'paintedHeight'] as const) expect(Math.abs(state[field] - states[0][field]), field).toBeLessThan(.6);
    expect(state.floating).toMatchObject(states[0].floating);
  }
  expect(states[1].inlineWidth).toBeLessThan(states[0].inlineWidth); expect(states[1].inlineHeight).toBeLessThan(states[0].inlineHeight);
});
