import { test, expect } from '@playwright/test';

for (const y of [50, 130]) test(`full-width page table at ${y}px preserves its anchor and following text through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async y => {
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const anchor = JSON.stringify({ x: 151.2, y, bottom: 12 }).replace(/"/g, '&quot;');
    const html = `<table data-table-anchor="${anchor}"><colgroup><col style="width:490px"></colgroup><tr style="height:60px"><td style="padding:0;border:0"><p style="line-height:20px;margin:0">TITLE</p></td></tr></table><p style="line-height:20px;margin:0">BODY</p>`;
    const loaded = await importFile(new File([html], 'title.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, margin: 40, pageNumbers: false }, content: parseDocumentHtml(html) }; const states = [], first = doc;
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const paper = layout.element.querySelector('.print-page')!, table = paper.querySelector('table')!, body = paper.querySelector('.print-content > p')!;
        states.push({ pages: layout.pages, tableY: table.getBoundingClientRect().top - paper.getBoundingClientRect().top, bodyY: body.getBoundingClientRect().top - paper.getBoundingClientRect().top, text: paper.querySelector('.print-content')!.textContent });
      } finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return { first, states };
  }, y);
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.tableY).toBeCloseTo(y, 0);
    expect(state.bodyY).toBeCloseTo(Math.max(40 * 96 / 25.4, y + 60 + 12), 0);
    expect(state.text).toBe('TITLEBODY');
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'title.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [100, 50, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.locator('table').evaluate((table, zoom) => (table.getBoundingClientRect().top - table.closest('.document-stack')!.getBoundingClientRect().top) / (zoom / 100), zoom)).toBeCloseTo(y, 0);
  }
  await editor.locator('td p').fill('EDIT TITLE'); await page.keyboard.press('ControlOrMeta+z'); await expect(editor.locator('td p')).toHaveText('TITLE');
});
