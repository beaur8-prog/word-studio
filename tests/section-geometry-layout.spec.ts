import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, HeightRule, LineRuleType, BorderStyle, SectionType } from 'docx';

const none = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const paragraph = (text: string) => new Paragraph({
  children: [new TextRun({ text, font: 'Arial', size: 18 })],
  spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.EXACT },
  widowControl: false,
});
const row = () => new Table({
  width: { size: 4500, type: WidthType.DXA }, columnWidths: [4500],
  borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
  rows: [new TableRow({ cantSplit: true, height: { value: 4500, rule: HeightRule.ATLEAST },
    children: [new TableCell({ margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [paragraph('CELL')] })] })],
});

for (const change of ['margins', 'custom-height'] as const) {
  test('same-kind Word paper with different ' + change + ' keeps editor and print pages through two saves', async ({ page }) => {
    const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } };
    const second = change === 'margins'
      ? { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300 } }
      : { ...first, size: { width: 6000, height: 6900 } };
    const bytes = await Packer.toBuffer(new Document({ sections: [
      { properties: { page: first }, children: [paragraph('FIRST')] },
      { properties: { type: SectionType.NEXT_PAGE, page: second }, children: [paragraph('LEAD'), row(), paragraph('AFTER')] },
    ] }));
    await writeFile(test.info().outputPath('textbox-save-section-source.docx'), bytes);
    await page.goto('/');
    const result = await page.evaluate(async bytes => {
      const { importFile, exportDocx } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      let input = new File([new Uint8Array(bytes)], 'sections.docx');
      const states = [], saves = [];
      for (let pass = 0; pass < 3; pass++) {
        const loaded = await importFile(input);
        const doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        const printed = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].filter(p => p.textContent?.trim()).map(p => {
          const paper = p.closest<HTMLElement>('.print-page')!, box = p.getBoundingClientRect(), root = paper.getBoundingClientRect();
          return { text: p.textContent, page: Number(paper.dataset.page), x: box.left - root.left, y: box.top - root.top, height: box.height };
        });
        states.push({ doc, pages: layout.pages, printed, warnings: layout.warnings }); layout.dispose();
        if (pass < 2) { const output = await exportDocx(doc); saves.push([...new Uint8Array(await output.arrayBuffer())]); input = new File([output], 'again.docx'); }
      }
      return { states, saves };
    }, [...bytes]);
    for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath('textbox-save-section-' + (index + 1) + '.docx'), Buffer.from(bytes));
    for (const state of result.states) {
      expect(state.pages).toBe(2); expect(state.warnings).toEqual([]);
      expect(state.printed.map(p => p.text)).toEqual(['FIRST', 'LEAD', 'CELL', 'AFTER']);
      await page.locator('input[type=file]').first().setInputFiles({ name: 'sections.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
      const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
      await expect(editor).toContainText('AFTER');
      await expect(page.locator('.busy-indicator')).toBeHidden();
      for (const zoom of [100, 50, 125, 100]) {
        const slider = page.getByRole('slider', { name: '확대 축소' });
        await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
        await expect.poll(async () => editor.evaluate((root, printed) => {
          const stack = root.closest<HTMLElement>('.document-stack')!, sections = [...root.querySelectorAll<HTMLElement>('.document-section-content')];
          let scale = 1;
          for (let p: Element | null = root; p; p = p.parentElement) {
            const css = getComputedStyle(p); scale *= parseFloat(css.zoom) || 1;
            if (css.transform !== 'none') { const m = new DOMMatrixReadOnly(css.transform); scale *= Math.hypot(m.a, m.b); }
          }
          const paragraphs = [...root.querySelectorAll('p')].filter(p => p.textContent?.trim());
          const actual = paragraphs.map(p => {
            const section = p.closest<HTMLElement>('.document-section-content'), local = section?.querySelector<HTMLElement>('.section-editable') || root;
            const h = Number(section?.dataset.pageHeight ?? stack.dataset.pageHeight), box = p.getBoundingClientRect(), origin = local.getBoundingClientRect();
            const y = (box.top - origin.top) / scale, localPage = Math.max(0, Math.floor((y + .5) / (h + 24)));
            const before = section ? sections.slice(0, sections.indexOf(section)).reduce((sum, s) => sum + Number(s.dataset.pages || 1), 0) : 0;
            return { text: p.textContent, page: before + localPage + 1, x: (box.left - origin.left) / scale, y: y - localPage * (h + 24), height: box.height / scale };
          });
          return { pages: Number(stack.dataset.pages),
            same: actual.length === printed.length && actual.every((p,i) => p.text === printed[i].text && p.page === printed[i].page && Math.abs(p.x - printed[i].x) <= 1 && Math.abs(p.y - printed[i].y) <= 1 && Math.abs(p.height - printed[i].height) <= .6) };
        }, state.printed), { timeout: 8000 }).toEqual({ pages: 2, same: true });
      }
    }
  });
}

test('framed section headers and PAGE/NUMPAGES use document-wide order and totals', async ({ page }) => {
  const { Header, Footer, PageNumber, PageBreak } = await import('docx');
  const header = (label: string) => new Header({ children: [paragraph(label)] });
  const footer = () => new Footer({ children: [new Paragraph({ children: [
    new TextRun('PAGE '), new TextRun({ children: [PageNumber.CURRENT] }),
    new TextRun(' OF '), new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
  ] })] });
  const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750, header: 100, footer: 100 } };
  const second = { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300, header: 100, footer: 100 } };
  const stories = () => ({ headers: { default: header('ODD'), even: header('EVEN') }, footers: { default: footer(), even: footer() } });
  const bytes = await Packer.toBuffer(new Document({ evenAndOddHeaderAndFooters: true, sections: [
    { ...stories(), properties: { page: first }, children: [paragraph('FIRST')] },
    { ...stories(), properties: { type: SectionType.NEXT_PAGE, page: second }, children: [
      paragraph('SECOND'), new Paragraph({ children: [new PageBreak()] }), paragraph('THIRD'),
    ] },
  ] }));
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'section-fields.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor).toContainText('THIRD');
  await expect(page.locator('.busy-indicator')).toBeHidden();
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('3');
  await expect(page.locator('.section-paper .section-header')).toHaveText(['ODD', 'EVEN', 'ODD']);
  await expect(page.locator('.section-paper .section-footer')).toHaveText(['PAGE 1 OF 3', 'PAGE 2 OF 3', 'PAGE 3 OF 3']);
  const original = await editor.evaluate((el: any) => el.editor.getJSON());
  await editor.locator('.section-editable p').filter({ hasText: /^FIRST$/ }).click();
  await page.keyboard.press('End'); await page.keyboard.insertText(' EDIT');
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor.evaluate((el: any) => el.editor.getJSON())).toEqual(original);
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('FIRST EDIT');
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor.evaluate((el: any) => el.editor.getJSON())).toEqual(original);
  await expect(page.locator('.section-paper .section-footer')).toHaveText(['PAGE 1 OF 3', 'PAGE 2 OF 3', 'PAGE 3 OF 3']);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page')).toHaveCount(3);
  await expect(page.locator('.print-header')).toHaveText(['ODD', 'EVEN', 'ODD']);
  await expect(page.locator('.print-footer')).toHaveText(['PAGE 1 OF 3', 'PAGE 2 OF 3', 'PAGE 3 OF 3']);
});

for (const [kind, pages] of [['continuous', 1], ['oddPage', 3], ['evenPage', 2]] as const) {
  test('different-margin ' + kind + ' sections retain their shared or parity pages', async ({ page }) => {
    const type = kind === 'continuous' ? SectionType.CONTINUOUS : kind === 'oddPage' ? SectionType.ODD_PAGE : SectionType.EVEN_PAGE;
    const first = { size: { width: 6000, height: 6000 }, margin: { top: 750, bottom: 750, left: 750, right: 750 } };
    const second = { ...first, margin: { top: 300, bottom: 300, left: 300, right: 300 } };
    const input = await Packer.toBuffer(new Document({ sections: [
      { properties: { page: first }, children: [paragraph('FIRST')] },
      { properties: { type, page: second }, children: [paragraph('SECOND')] },
    ] }));
    await page.goto('/');
    const result = await page.evaluate(async bytes => {
      const { importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const opened = await importFile(new File([new Uint8Array(bytes)], 'control.docx'));
      const doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      const result = { doc, pages: layout.pages, text: [...layout.element.querySelectorAll('.print-content p')].map(p => p.textContent).filter(Boolean) };
      layout.dispose(); return result;
    }, [...input]);
    expect(result.pages).toBe(pages); expect(result.text).toEqual(['FIRST', 'SECOND']);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'control.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
    await expect(page.getByRole('textbox', { name: '문서 편집 영역', exact: true })).toContainText('SECOND');
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(String(pages));
  });
}
