import { test, expect } from '@playwright/test';

for (const border of ['border:none', 'border-top:none;border-right:none;border-bottom:none;border-left:none', 'border:hidden']) {
  test(`${border}: borderless layout cells survive editing and two DOCX saves`, async ({ page }) => {
    await page.goto('/');
    const result = await page.evaluate(async border => {
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      let doc = { format: 'word-studio', version: 1, title: 'Borderless layout', updatedAt: new Date().toISOString(), comments: [],
        page: { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', footer: '', pageNumbers: false },
        content: parseDocumentHtml(`<table><tr><td colwidth="240" style="padding:0;${border}"><p>LAYOUT LABEL</p><table><tr><td style="border:2px solid #c00000"><p>INNER CONTENT</p></td></tr></table><p></p></td><td colwidth="240"><p>DEFAULT CELL</p></td></tr></table>`) };
      const first = doc, states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const outer = layout.element.querySelector('table')!, cell = outer.rows[0].cells[0], inner = cell.querySelector('td')!, normal = outer.rows[0].cells[1];
          const edges = (node: HTMLElement) => ['Top', 'Right', 'Bottom', 'Left'].map(side => ({ style: getComputedStyle(node)[`border${side}Style` as 'borderTopStyle'], width: getComputedStyle(node)[`border${side}Width` as 'borderTopWidth'] }));
          states.push({ pages: layout.pages, outer: edges(cell), inner: edges(inner), normal: edges(normal), text: outer.textContent });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { first, states };
    }, border);
    for (const state of result.states) {
      expect(state.pages).toBe(1); expect(state.text).toBe('LAYOUT LABELINNER CONTENTDEFAULT CELL');
      for (const edge of state.outer) { expect(['none', 'hidden']).toContain(edge.style); expect(edge.width).toBe('0px'); }
      for (const edge of state.inner) { expect(edge.style).toBe('solid'); expect(edge.width).toBe('2px'); }
      for (const edge of state.normal) { expect(edge.style).toBe('solid'); expect(edge.width).toBe('1px'); }
    }
    await page.locator('input[type=file]').first().setInputFiles({ name: 'layout.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
    const cell = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('table').first().locator(':scope > tbody > tr > td').first();
    for (const zoom of [100, 50, 125, 200, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect(cell).toHaveCSS('border-top-width', '0px');
    }
    const paragraph = cell.locator(':scope > p').first(); await paragraph.fill('EDITED LABEL'); await expect(paragraph).toHaveText('EDITED LABEL');
    await page.keyboard.press('ControlOrMeta+z'); await expect(paragraph).toHaveText('LAYOUT LABEL'); await expect(cell).toHaveCSS('border-top-width', '0px');
  });
}
