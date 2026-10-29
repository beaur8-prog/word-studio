import { test, expect } from '@playwright/test';

for (const kind of ['trailing', 'followed', 'blank-only']) test(`${kind}: real blank lines keep their page slots in print, editor and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async kind => {
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const body = kind === 'blank-only' ? '<br>'.repeat(16) : Array.from({ length: 8 }, (_, i) => 'LINE' + i).join('<br>') + '<br><br>';
    const style = 'font-family:Arial;font-size:10pt;line-height:20px;margin:0';
    const html = `<p style="${style};line-height:130px">LEAD</p><p style="${style}">${body}</p>` + (kind === 'trailing' ? '' : `<p style="${style}">AFTER</p>`);
    const loaded = await importFile(new File([html], 'blank-lines.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 100, margin: 10, pageNumbers: false }, content: parseDocumentHtml(loaded.html!) };
    const original = doc, states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const after = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].find(p => p.textContent === 'AFTER');
        let position = null;
        if (after) { const r = document.createRange(); r.selectNodeContents(after); const paper = after.closest<HTMLElement>('.print-page')!;
          position = { page: Number(paper.dataset.page) - 1, y: r.getBoundingClientRect().top - paper.getBoundingClientRect().top };
        }
        states.push({ pages: layout.pages, breaks: layout.element.querySelectorAll('.print-content br').length, position });
      } finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return { original, states, breaks: kind === 'blank-only' ? 16 : 9 };
  }, kind);
  for (const state of result.states) {
    expect(state.pages).toBe(2); expect(state.breaks).toBe(result.breaks);
    expect(state.position?.page).toBe(result.states[0].position?.page);
    if (state.position) expect(Math.abs(state.position.y - result.states[0].position!.y)).toBeLessThan(1.5);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'blank-lines.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect(page.locator('.document-stack')).toHaveAttribute('data-pages', '2');
    if (kind !== 'trailing') await expect.poll(() => editor.locator('p').filter({ hasText: /^AFTER$/ }).evaluate((p, { zoom, expected }) => {
      const stack = p.closest<HTMLElement>('.document-stack')!, root = p.closest<HTMLElement>('.tiptap')!, stride = Number(stack.dataset.pageHeight) + 24;
      const r = document.createRange(); r.selectNodeContents(p); const y = (r.getBoundingClientRect().top - root.getBoundingClientRect().top) / (zoom / 100);
      const sheet = Math.max(0, Math.floor((y + .5) / stride));
      return sheet === expected.page && Math.abs(y - sheet * stride - expected.y) < 1.5;
    }, { zoom, expected: result.states[0].position! })).toBe(true);
  }
});
