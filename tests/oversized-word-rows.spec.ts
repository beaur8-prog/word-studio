import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, LineRuleType, BorderStyle } from 'docx';

const none = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const paragraph = (text: string) => new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 18 })],
  spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.EXACT }, widowControl: false });
const table = (rows: TableRow[], columns: number, edge = none) => new Table({ width: { size: 4500, type: WidthType.DXA },
  columnWidths: Array(columns).fill(4500 / columns), borders: { top: edge, bottom: edge, left: edge, right: edge, insideHorizontal: edge, insideVertical: edge }, rows });
const cell = (children: any[], padding = 0) => new TableCell({ children, margins: { top: padding * 15, bottom: padding * 15, left: padding * 15, right: padding * 15 } });

const cases = [
  { columns: 1, nested: false, padding: 0, header: false },
  { columns: 2, nested: false, padding: 0, header: false },
  { columns: 1, nested: true, padding: 0, header: false },
  { columns: 2, nested: true, padding: 0, header: false },
  { columns: 2, nested: false, padding: 7, header: true },
  { columns: 1, nested: true, padding: 5, header: false },
  { columns: 2, nested: true, padding: 3, header: true },
];
for (const { columns, nested, padding, header } of cases) {
  test(`oversized Word cantSplit ${columns}-column ${nested ? 'nested' : 'body'} padding ${padding} header ${header} continues at its original size`, async ({ page }) => {
    const lineCount = padding ? 20 : 24, lastLabel = `C0LINE${String(lineCount - 1).padStart(2, '0')}`;
    const edge = padding ? { style: BorderStyle.SINGLE, size: 12, color: '334455' } : none;
    const rows = [new TableRow({ cantSplit: true, children: Array.from({ length: columns }, (_, col) =>
      cell(Array.from({ length: lineCount }, (_, line) => paragraph(`C${col}LINE${String(line).padStart(2, '0')}`)), padding)) })];
    if (header) rows.unshift(new TableRow({ tableHeader: true, children: Array.from({ length: columns }, (_, col) => cell([paragraph(`HEADER ${col}`)], padding)) }));
    const content = table(rows, columns, edge);
    const body = nested ? table([new TableRow({ cantSplit: false, children: [cell([paragraph('OUTER'), content, paragraph('OUTER END')], padding)] })], 1, edge) : content;
    const input = await Packer.toBuffer(new Document({ sections: [{ properties: { page: { size: { width: 6000, height: 6000 },
      margin: { top: 750, right: 750, bottom: 750, left: 750 } } }, children: [paragraph('LEAD'), body, paragraph('AFTER')] }] }));
    await writeFile(test.info().outputPath('textbox-save-input-oversized.docx'), input);
    await page.goto('/');
    const result = await page.evaluate(async bytes => {
      const { importFile, exportDocx } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      let file = new File([new Uint8Array(bytes)], 'oversized.docx'), doc: any;
      const states: any[] = [], saves: number[][] = [];
      for (let save = 0; save < 3; save++) {
        const opened = await importFile(file); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) };
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const paragraphs = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].filter(p => p.textContent?.trim()).map(p => {
            const box = p.getBoundingClientRect(), root = p.closest('.print-content')!.getBoundingClientRect();
            return { text: p.textContent, page: Number(p.closest<HTMLElement>('.print-page')!.dataset.page),
              font: getComputedStyle(p.querySelector('span') || p).fontSize, height: box.height, top: box.top - root.top, bottom: box.bottom - root.top, body: root.height };
          });
          states.push({ pages: layout.pages, paragraphs, warnings: layout.warnings,
            zooms: [...layout.element.querySelectorAll<HTMLElement>('.print-content [style*="zoom"]')].map(el => el.style.zoom),
            text: paragraphs.map(p => p.text).join('|') });
        } finally { layout.dispose(); }
        if (save < 2) { const output = await exportDocx(doc); saves.push([...new Uint8Array(await output.arrayBuffer())]); file = new File([output], 'again.docx'); }
      }
      return { states, doc, saves };
    }, [...input]);
    for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-oversized-${index + 1}.docx`), Buffer.from(bytes));
    for (const state of result.states) {
      expect(state.warnings).toEqual([]);
      expect(state.zooms).toEqual([]);
      expect(state.pages).toBe(3);
      expect(state.paragraphs.find(p => p.text === 'LEAD').page).toBe(1);
      for (let col = 0; col < columns; col++) {
        expect(state.paragraphs.find(p => p.text === `C${col}LINE00`).page).toBe(2);
        expect(state.paragraphs.find(p => p.text === `C${col}LINE${String(lineCount - 1).padStart(2, '0')}`).page).toBe(3);
        expect(state.paragraphs.filter(p => p.text.startsWith(`C${col}LINE`))).toHaveLength(lineCount);
      }
      for (const p of state.paragraphs.filter(p => p.text.includes('LINE'))) {
        expect(parseFloat(p.font)).toBeCloseTo(12, 1);
        expect(p.height).toBeCloseTo(20, 1);
        expect(p.top).toBeGreaterThanOrEqual(-.6);
        expect(p.bottom).toBeLessThanOrEqual(p.body + .6);
      }
    }
    expect(result.states.map(s => s.text)).toEqual(Array(3).fill(result.states[0].text));
    await page.locator('input[type=file]').first().setInputFiles({ name: 'oversized.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor).toContainText(lastLabel);
    const inspect = () => editor.evaluate(root => {
      const sheet = root.closest<HTMLElement>('.document-stack')!, h = Number(sheet.dataset.pageHeight), origin = root.getBoundingClientRect().top;
      let scale = 1;
      for (let p: Element | null = root; p; p = p.parentElement) {
        const css = getComputedStyle(p); scale *= parseFloat(css.zoom) || 1;
        if (css.transform !== 'none') { const m = new DOMMatrixReadOnly(css.transform); scale *= Math.hypot(m.a, m.b); }
      }
      return [...root.querySelectorAll('p')].filter(p => /^C[01]LINE/.test(p.textContent || '')).map(p => {
        const r = p.getBoundingClientRect(), y = (r.top - origin) / scale, page = Math.floor((y + .5) / (h + 24));
        return { text: p.textContent, page: page + 1, y: y - page * (h + 24), height: r.height / scale,
          top: Number(sheet.dataset.margin), bottom: h - Number(sheet.dataset.bottomMargin) };
      });
    });
    for (const zoom of [100, 50, 125, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' });
      await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(async () => {
        const rows = await inspect();
        return { first: rows.find(r => r.text === 'C0LINE00')?.page, last: rows.find(r => r.text === lastLabel)?.page,
          inBody: rows.every(r => r.y >= r.top - .6 && r.y + r.height <= r.bottom + .6),
          samePrintedLines: rows.every(r => {
            const printed = result.states[0].paragraphs.find(p => p.text === r.text);
            return printed?.page === r.page && Math.abs(r.y - r.top - printed.top) <= 1.25;
          }) };
      }).toEqual({ first: 2, last: 3, inBody: true, samePrintedLines: true });
    }
    if (header) {
      const headers = await editor.evaluate(root => {
        const sheet = root.closest<HTMLElement>('.document-stack')!, h = Number(sheet.dataset.pageHeight), origin = root.getBoundingClientRect().top;
        return [...sheet.querySelectorAll('p')].filter(p => (p.textContent || '').startsWith('HEADER')).map(p => ({
          text: p.textContent, page: Math.floor((p.getBoundingClientRect().top - origin + .5) / (h + 24)) + 1,
        })).sort((a, b) => a.page - b.page || a.text!.localeCompare(b.text!));
      });
      const printed = result.states[0].paragraphs.filter(p => p.text.startsWith('HEADER')).map(p => ({ text: p.text, page: p.page })).sort((a, b) => a.page - b.page || a.text.localeCompare(b.text));
      expect(headers).toEqual(printed);
    }
    const before = await editor.evaluate((el: any) => el.editor.getJSON());
    const target = editor.locator('p').filter({ hasText: /^C0LINE00$/ });
    await target.fill('EDIT C0LINE00'); await page.keyboard.press('ControlOrMeta+z');
    await expect(target).toHaveText('C0LINE00');
    expect(await editor.evaluate((el: any) => el.editor.getJSON())).toEqual(before);
    await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('EDIT C0LINE00');
  });
}
