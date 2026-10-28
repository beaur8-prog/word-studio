import { test, expect } from '@playwright/test';

test('a repeated header reserves its collapsed border at a fractional line boundary through editing, zoom and two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const models = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts'), { defaultPage } = await import('/src/types.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const line = (text: string) => `<p style="font-family:Arial;font-size:12pt;line-height:22.4px;margin:0">${text}</p>`;
    const cell = (column: string) => `<td colwidth="200" style="padding:1.88px;border-left:1px solid black;border-right:1px solid black;border-top:none;border-bottom:none">${Array.from({ length: 36 }, (_, i) => line(column + ' ' + i)).join('')}</td>`;
    const html = `<table style="width:400px;margin:0"><colgroup><col style="width:200px"><col style="width:200px"></colgroup><tr style="height:45.76px"><th colwidth="200" style="padding:1.88px;border:1px solid black">${line('HEADER LEFT')}</th><th colwidth="200" style="padding:1.88px;border:1px solid black">${line('HEADER RIGHT')}</th></tr><tr data-row-split="allow" style="height:772.33px">${cell('LEFT')}${cell('RIGHT')}</tr></table>`;
    const loaded = await importFile(new File([html], 'border-header.html'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!), page: { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false } };
    const probe = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
    try {
      const table = probe.element.querySelector('table')!, row = table.rows[1], paragraph = row.cells[0].querySelector('p')!;
      const header = row.getBoundingClientRect().top - table.getBoundingClientRect().top;
      const padding = parseFloat(getComputedStyle(row.cells[0]).paddingTop) + parseFloat(getComputedStyle(row.cells[0]).paddingBottom);
      const usable = header + padding + 12 * paragraph.getBoundingClientRect().height - .7;
      doc.page.height = usable * 25.4 / 96 + 20;
    } finally { probe.dispose(); }
    const result = [];
    for (let save = 0; save < 3; save++) {
      doc.title = 'REPEATING HEADER ' + save;
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        result.push({ model: doc, pages: layout.pages, starts: [...layout.element.querySelectorAll('p')].filter(p => /^(?:LEFT|RIGHT) \d+$/.test(p.textContent || '')).map(p => ({ text: p.textContent!, page: Number(p.closest<HTMLElement>('.print-page')!.dataset.page) })) });
      } finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return result;
  });
  for (const state of models) {
    expect(state.pages).toBeGreaterThanOrEqual(3); expect(state.starts).toHaveLength(72);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'headers.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.model)) });
    await expect(page.getByRole('textbox', { name: '문서 제목' })).toHaveValue(state.model.title);
    for (const zoom of [100, 50, 125, 200, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(async () => page.evaluate(({ zoom, expected }) => {
        const root = document.querySelector<HTMLElement>('[role="textbox"][contenteditable]')!, sheet = root.closest<HTMLElement>('.document-stack')!;
        const origin = root.getBoundingClientRect().top, stride = Number(sheet.dataset.pageHeight) + 24;
        const actual = [...root.querySelectorAll('p')].filter(p => /^(?:LEFT|RIGHT) \d+$/.test(p.textContent || '')).map(p => {
          const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT); let text: Text | null = null;
          while (walker.nextNode()) if (walker.currentNode.textContent?.trim() && !walker.currentNode.parentElement?.closest('.ProseMirror-widget')) { text = walker.currentNode as Text; break; }
          const range = document.createRange(); range.setStart(text!, 0); range.setEnd(text!, 1);
          return { text: p.textContent!, page: Math.floor(((range.getBoundingClientRect().top - origin) / (zoom / 100) + .5) / stride) + 1 };
        });
        return { pages: Number(sheet.dataset.pages), mismatches: actual.filter(p => expected.find(e => e.text === p.text)?.page !== p.page) };
      }, { zoom, expected: state.starts })).toEqual({ pages: state.pages, mismatches: [] });
    }
    const first = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('td').first().locator('p').first();
    await first.fill('LEFT 0 EDIT'); await expect(first).toHaveText('LEFT 0 EDIT');
    await page.keyboard.press('ControlOrMeta+z'); await expect(first).toHaveText('LEFT 0');
  }
});
