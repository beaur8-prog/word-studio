import { test, expect } from '@playwright/test';

for (const alignment of ['center', 'right'] as const) test(`${alignment} body and footer tables retain alignment through two DOCX saves and the editor`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async alignment => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const table = `<table style="width:120px;margin-left:auto;margin-right:${alignment === 'center' ? 'auto' : '0'}"><tr><td colwidth="120" style="width:120px;padding:0"><p>TABLE</p></td></tr></table>`;
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Table alignment', updatedAt: new Date().toISOString(), comments: [],
      page: { size: 'A4', orientation: 'portrait', margin: 25, header: '', footer: '', footerHtml: table, pageNumbers: false }, content: parseDocumentHtml('<p>BODY</p>' + table) };
    const states = [];
    for (let round = 0; round < 3; round++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const difference = (container: Element) => {
          const box = container.getBoundingClientRect(), table = container.querySelector('table')!.getBoundingClientRect();
          return alignment === 'center' ? table.left + table.width / 2 - (box.left + box.width / 2) : table.right - box.right;
        };
        states.push({ body: difference(layout.element.querySelector('.print-content')!), footer: difference(layout.element.querySelector('.print-footer')!) });
      } finally { layout.dispose(); }
      if (round < 2) { const imported = await importFile(new File([await exportDocx(doc)], 'aligned.docx')); doc = { ...imported.document, content: parseDocumentHtml(imported.html!) }; }
    }
    return { doc, states };
  }, alignment);
  for (const state of result.states) { expect(Math.abs(state.body)).toBeLessThan(.2); expect(Math.abs(state.footer)).toBeLessThan(.2); }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'alignment.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  await expect.poll(async () => page.locator('.tiptap').evaluate((editor, alignment) => {
    const box = editor.querySelector('p')!.getBoundingClientRect(), table = editor.querySelector('table')!.getBoundingClientRect();
    return Math.abs(alignment === 'center' ? table.left + table.width / 2 - (box.left + box.width / 2) : table.right - box.right);
  }, alignment)).toBeLessThan(.3);
});
