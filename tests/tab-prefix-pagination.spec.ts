import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { Document, Packer, Paragraph, TextRun, Tab, TabStopType, LineRuleType } from 'docx';

for (const stop of [4500, 3600]) test('right tab at ' + stop + ' twips uses the same legal page cut in editor and print', async ({ page }) => {
  const input = await Packer.toBuffer(new Document({ sections: [{
    properties: { page: { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } } },
    children: [
      new Paragraph({ children: [new TextRun({ text: 'LEAD', font: 'Arial', size: 18 })],
        spacing: { before: 0, after: 0, line: 4000, lineRule: LineRuleType.EXACT }, widowControl: false }),
      new Paragraph({ children: [new TextRun({ children: ['FIELD', new Tab(), '123'], font: 'Arial', size: 18 })],
        tabStops: [{ type: TabStopType.RIGHT, position: stop }], indent: { left: 150, right: 300 },
        spacing: { before: 0, after: 0, line: 450, lineRule: LineRuleType.EXACT }, widowControl: false }),
      new Paragraph({ children: [new TextRun({ text: 'AFTER', font: 'Arial', size: 18 })],
        spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.EXACT }, widowControl: false }),
    ],
  }] }));
  await writeFile(test.info().outputPath('textbox-save-tab-prefix-source.docx'), input);
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const states = [], saves = []; let file = new File([new Uint8Array(bytes)], 'tab-prefix.docx');
    for (let pass = 0; pass < 3; pass++) {
      const loaded = await importFile(file); const doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      const p = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].find(p => p.textContent?.startsWith('FIELD'))!;
      const paper = p.closest<HTMLElement>('.print-page')!, origin = paper.getBoundingClientRect(), glyphs = [];
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const text = walker.currentNode as Text;
        for (let i = 0; i < text.length; i++) if (/\S/.test(text.data[i])) {
          const range = document.createRange(); range.setStart(text, i); range.setEnd(text, i + 1);
          const box = range.getBoundingClientRect();
          glyphs.push({ text: text.data[i], page: Number(paper.dataset.page), x: box.x - origin.x, y: box.y - origin.y });
        }
      }
      states.push({ doc, pages: layout.pages, glyphs, text: layout.element.querySelector('.print-content')?.textContent }); layout.dispose();
      if (pass < 2) { const output = await exportDocx(doc); saves.push([...new Uint8Array(await output.arrayBuffer())]); file = new File([output], 'again.docx'); }
    }
    return { states, saves };
  }, [...input]);
  for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath('textbox-save-tab-prefix-' + (index + 1) + '.docx'), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.pages).toBe(2);
    expect(state.glyphs.map(g => g.text).join('')).toBe('FIELD123');
    await page.locator('input[type=file]').first().setInputFiles({ name: 'tabs.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor).toContainText('FIELD'); await expect(page.locator('.busy-indicator')).toBeHidden();
    for (const zoom of [100, 50, 125, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' });
      await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(async () => { const measured = await editor.evaluate((root, printed) => {
        const sheet = root.closest<HTMLElement>('.document-stack')!, height = Number(sheet.dataset.pageHeight), origin = root.getBoundingClientRect();
        let scale = 1;
        for (let a: Element | null = root; a; a = a.parentElement) {
          const css = getComputedStyle(a); scale *= parseFloat(css.zoom) || 1;
          if (css.transform !== 'none') { const m = new DOMMatrixReadOnly(css.transform); scale *= Math.hypot(m.a,m.b); }
        }
        const p = [...root.querySelectorAll('p')].find(p => p.textContent?.startsWith('FIELD'))!, walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), actual = [];
        while (walker.nextNode()) {
          const text = walker.currentNode as Text;
          if (text.parentElement?.closest('.pagination-cell-gap,.pagination-gap')) continue;
          for (let i = 0; i < text.length; i++) if (/\S/.test(text.data[i])) {
            const range = document.createRange(); range.setStart(text,i); range.setEnd(text,i+1);
            const box = range.getBoundingClientRect(), y = (box.y-origin.y)/scale, page = Math.floor((y+.5)/(height+24));
            actual.push({ text:text.data[i], page:page+1, x:(box.x-origin.x)/scale, y:y-page*(height+24) });
          }
        }
        return { actual, printed, pages:Number(sheet.dataset.pages), same:actual.length===printed.length && actual.every((g,i) => g.text===printed[i].text && g.page===printed[i].page && Math.abs(g.x-printed[i].x)<1 && Math.abs(g.y-printed[i].y)<1) };
      }, state.glyphs); await writeFile(test.info().outputPath('tab-prefix-last-inspection.json'),JSON.stringify(measured,null,2)); return {pages:measured.pages,same:measured.same}; }, { timeout:8000 }).toEqual({ pages:2, same:true });
    }
  }
});
