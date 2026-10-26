import { test, expect } from '@playwright/test';

test('explicit page breaks in unbalanced columns retain all physical pages through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const html = '<div data-column-count="2" data-column-space="24" data-column-balance="false"><p>ONE</p><div data-page-break="true"></div><p>TWO</p><div data-page-break="true"></div><p>THREE</p></div>';
    const loaded = await importFile(new File([html], 'columns.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; const states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ pages: layout.pages, text: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent), content: doc.content }); }
      finally { layout.dispose(); }
      if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
    }
    return { states, doc };
  });
  for (const state of result.states) { expect(state.pages).toBe(3); expect(state.text).toEqual(['ONE', 'TWO', 'THREE']); }
  for (const state of result.states.slice(1)) {
    expect(state.content.content.filter(n => n.type === 'pageBreak')).toHaveLength(2);
    expect(state.content.content.filter(n => n.type === 'columnGroup')).toHaveLength(3);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'columns.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('3');
});

test('a final column group does not acquire an extra blank page from an outside trailing paragraph', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const html = '<div data-column-count="2" data-column-space="24" data-column-balance="false">' + Array.from({ length: 110 }, (_, i) => `<p style="line-height:20px;margin:0;font-size:10pt">LINE ${i}</p>`).join('') + '</div>';
    const loaded = await importFile(new File([html], 'end.html'));
    const doc = { ...loaded.document, page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false }, content: parseDocumentHtml(html) };
    const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try { return { doc, pages: layout.pages }; } finally { layout.dispose(); }
  });
  expect(result.pages).toBe(2);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'end.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('2');
  await expect(editor.locator(':scope > p')).toHaveCount(0);
  await editor.locator('p').last().fill('EDIT LAST LINE');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor.locator('p').last()).toHaveText('LINE 109');
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('2');
});
