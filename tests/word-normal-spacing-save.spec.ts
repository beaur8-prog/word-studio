import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, LineRuleType, WidthType } from 'docx';

for (const single of [true, false]) test(`Word ${single ? 'implicit single' : 'explicit 1.15'} leading preserves page and table coordinates through two saves`, async ({ page }) => {
  const spacing = { before: 0, after: 0, ...(single ? {} : { line: 276, lineRule: LineRuleType.AUTO }) };
  const para = (text: string) => new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 22 })], spacing });
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 22 }, paragraph: { spacing } } } },
    sections: [{ properties: { page: { size: { width: 6000, height: 6000 }, margin: { top: 450, bottom: 450, left: 450, right: 450 } } }, children: [
      para('FIRST TITLE'),
      new Table({ width: { size: 4500, type: WidthType.DXA }, columnWidths: [2250, 2250], rows: Array.from({ length: 20 }, (_, i) => new TableRow({ children: [
        new TableCell({ children: [para(`ROW ${i} LEFT\nsecond line`)] }),
        new TableCell({ children: [para(`ROW ${i} RIGHT`)] }),
      ] })) }), para('LAST PARAGRAPH'),
    ] }],
  })));
  // This independent input has no inherited line multiplier: Word's implicit
  // single spacing must survive the exporter's otherwise different defaults.
  if (single) zip.file('word/styles.xml', (await zip.file('word/styles.xml')!.async('string')).replace(/ w:line="[^"]*"/g, '').replace(/ w:lineRule="[^"]*"/g, ''));
  const input = await zip.generateAsync({ type: 'uint8array' });
  await writeFile(test.info().outputPath(`textbox-save-normal-${single}-source.docx`), input);
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes }) => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    let file = new File([new Uint8Array(bytes)], 'normal.docx'); const states = [], saves = [];
    for (let pass = 0; pass < 3; pass++) {
      const loaded = await importFile(file), doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      const paragraphs = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].map(p => {
        const box = p.getBoundingClientRect(), paper = p.closest<HTMLElement>('.print-page')!;
        return { text: p.textContent, page: Number(paper.dataset.page), y: box.top - paper.getBoundingClientRect().top, height: box.height };
      });
      states.push({ pages: layout.pages, paragraphs, doc }); layout.dispose();
      if (pass < 2) { const output = await exportDocx(doc); saves.push([...new Uint8Array(await output.arrayBuffer())]); file = new File([output], 'saved.docx'); }
    }
    return { states, saves };
  }, { bytes: [...input] });
  for (const [i, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-normal-${single}-${i+1}.docx`), Buffer.from(bytes));
  await writeFile(test.info().outputPath('normal-spacing-coordinates.json'), JSON.stringify(result.states.map(({doc,...rest})=>rest),null,2));
  const before = result.states[0];
  expect(before.pages).toBeGreaterThan(1);
  for (const state of result.states.slice(1)) {
    expect(state.pages).toBe(before.pages);
    expect(state.paragraphs.map(p=>p.text)).toEqual(before.paragraphs.map(p=>p.text));
    for (const [i,p] of state.paragraphs.entries()) {
      expect(p.page, p.text || '').toBe(before.paragraphs[i].page);
      expect(Math.abs(p.y-before.paragraphs[i].y), p.text || '').toBeLessThan(1);
      expect(Math.abs(p.height-before.paragraphs[i].height), p.text || '').toBeLessThan(1);
    }
  }
});
