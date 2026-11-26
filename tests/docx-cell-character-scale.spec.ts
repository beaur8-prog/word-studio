import { expect, test } from '@playwright/test';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx';
import { readFile } from 'node:fs/promises';

const sample = 'WWWWWWWWW WWWWWWWWW';
const width = { size: 2000, type: WidthType.DXA } as const;

for (const type of ['fixed', 'autofit', 'nested fixed', 'multiple paragraphs'] as const)
test(`uniform character scales in an explicit ${type} table wrap within their physical cells`, async ({ page }, info) => {
  const nested = type === 'nested fixed', multi = type === 'multiple paragraphs';
  const layout = nested || multi ? 'fixed' : type;
  const text = (value: string, scale: number) => new TextRun({ text: value, font: 'Courier New', size: 22, scale });
  const paragraphs = [
    new Paragraph({ children: [text(sample, 65)] }),
    new Paragraph({ children: [text(sample, 100)] }),
    new Paragraph({ children: [text(sample.slice(0, 10), 65), text(sample.slice(10), 100)] }),
    new Paragraph({ indent: { left: 45 }, children: [text(sample, 65)] }),
  ];
  if (type === 'fixed') paragraphs.push(
    new Paragraph({ indent: { left: 209, hanging: 209 }, children: [text(sample, 90)] }),
  );
  const mergedWidth = { size: 650, type: WidthType.DXA } as const;
  const mergedNeighbor = layout === 'autofit';
  const table = new Table({
    width: mergedNeighbor ? { size: 2650, type: WidthType.DXA } : width,
    layout, columnWidths: nested ? [950, 1050] : mergedNeighbor ? [650, 2000] : [2000],
    rows: paragraphs.map((paragraph, index) => {
      const first = new Paragraph({ children: [text('Lead', index === 1 ? 100 : 65)] });
      const children = [new TableCell({ width, ...(nested ? { columnSpan: 2 } : {}),
        children: multi ? [first, paragraph] : [paragraph] })];
      if (mergedNeighbor && index !== 1) children.unshift(new TableCell({
        width: mergedWidth, rowSpan: index === 0 ? 2 : 1, children: [new Paragraph('A')],
      }));
      return new TableRow({ children });
    }),
  });
  const hostWidth = { size: 3600, type: WidthType.DXA } as const;
  const body = nested ? new Table({ width: hostWidth, layout: 'fixed', columnWidths: [3600], rows: [
    new TableRow({ children: [new TableCell({ width: hostWidth, children: [table, new Paragraph('')] })] }),
  ] }) : table;
  let bytes = await Packer.toBuffer(new Document({ sections: [{ children: [body] }] }));
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
  const input = page.locator('input[type=file]').first();

  for (let stage = 0; stage <= 2; stage++) {
    const name = `scaled-cells-${stage}`;
    await input.setInputFiles({ name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
    await expect(title).toHaveValue(name);
    const samples = editor.locator('td > p').filter({ hasText: sample });
    await expect(samples).toHaveCount(paragraphs.length);
    await expect.poll(() => samples.first().evaluate(p => getComputedStyle(p).transform)).toBe('matrix(0.65, 0, 0, 1, 0, 0)');
    const geometry = await page.evaluate(async ({ data, sample }) => {
      const { importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { wordLayoutOptions } = await import('/src/word-trailing-spaces.ts');
      const imported = await importFile(new File([new Uint8Array(data)], 'scaled-cells.docx'));
      const { document: model, html } = imported;
      const print = await createPrintLayout(generateHTML(parseDocumentHtml(html!), editorExtensions), model.page, undefined, undefined, wordLayoutOptions(model));
      function bounds(root: ParentNode) {
        return [...root.querySelectorAll('td > p')].filter(p => p.textContent === sample).map(p => {
          const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          const cell = p.parentElement!.getBoundingClientRect();
          const tops = new Set<number>();
          let outside = 0;
          const lineStarts = new Map<number, number>();
          while (walker.nextNode()) {
            const range = document.createRange();
            range.selectNodeContents(walker.currentNode);
            for (const rect of range.getClientRects()) if (rect.width > .1) {
              tops.add(rect.top);
              lineStarts.set(rect.top, Math.min(lineStarts.get(rect.top) ?? Infinity, rect.left));
              outside = Math.max(outside, cell.left - rect.left, rect.right - cell.right);
            }
          }
          const starts = [...lineStarts.entries()].sort(([a], [b]) => a - b).map(([, x]) => x);
          return { text: p.textContent, scale: getComputedStyle(p).transform, lines: tops.size, outside,
            firstOffset: starts.length === 2 ? starts[1] - starts[0] : null };
        });
      }
      try { return { editor: bounds(document.querySelector('.tiptap.ProseMirror')!), print: bounds(print.element) }; }
      finally { print.dispose(); }
    }, { data: [...bytes], sample });
    await info.attach(`stage-${stage}-geometry`, { body: JSON.stringify(geometry), contentType: 'application/json' });
    for (const scope of [geometry.editor, geometry.print]) {
      expect(scope.map(item => item.text)).toEqual(Array(paragraphs.length).fill(sample));
      expect(scope.slice(0, 4).map(item => item.scale)).toEqual(['matrix(0.65, 0, 0, 1, 0, 0)', 'none', 'none', 'none']);
      if (type === 'fixed') {
        expect(scope[4].scale).toBe('matrix(0.9, 0, 0, 1, 0, 0)');
        expect(scope[4].lines).toBe(2);
        expect(scope[4].firstOffset).toBeCloseTo(209 / 15, 1);
        expect(scope[4].outside).toBeLessThan(.25);
      }
      expect(scope[0].lines).toBe(1);
      expect(scope[1].lines).toBe(2);
      expect(scope[0].outside).toBeLessThan(.25);
    }
    if (stage === 2) break;
    const targetTable = editor.locator(nested ? 'table table' : 'table').first();
    const cell = targetTable.locator(':scope > tbody > tr').first().locator(':scope > td > p').last();
    await cell.click({ position: { x: 15, y: 8 } });
    await page.keyboard.type('Z');
    await expect(cell).toContainText('Z');
    await page.keyboard.press('ControlOrMeta+z');
    await expect(cell).toHaveText(sample);
    await page.keyboard.press('ControlOrMeta+Shift+z');
    await expect(cell).toContainText('Z');
    await page.keyboard.press('ControlOrMeta+z');
    await expect(cell).toHaveText(sample);
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    bytes = await readFile((await (await pending).path())!);
  }
});
