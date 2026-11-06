import { test, expect } from '@playwright/test';
import { createRequire } from 'node:module';
import fs from 'node:fs/promises';
const fixture = createRequire(import.meta.url)('./hwp-empty-cell-terminal-fixture.cjs');
let samples: Record<string, any>;
test.beforeAll(async () => { samples = await fixture(); });

for (const format of ['hwp', 'hwpx']) for (const mode of ['empty', 'space', 'two', 'break'])
test(`${format} ${mode}: empty terminal cell lines keep the same height at every table depth through editing and two DOCX saves`, async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const groups = [];
    for (const [index, source] of bytes.entries()) {
      const reader = new HwpReaderSession(); let doc;
      try { await reader.open(new File([new Uint8Array(source)], 'blank.' + format)); doc = importEditableHwp(await reader.editable(), 'blank.' + format); }
      finally { reader.close(); }
      const original = structuredClone(doc), states = [], files = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const table = [...layout.element.querySelectorAll<HTMLTableElement>('table')].at(-1)!, cell = table.rows[0].cells[1];
          states.push({ pages: layout.pages, height: table.getBoundingClientRect().height, cellText: cell.textContent,
            paragraphs: cell.querySelectorAll('p').length, breaks: cell.querySelectorAll('br').length,
            lastHeight: cell.querySelector('p:last-child')!.getBoundingClientRect().height, text: layout.element.querySelector('.print-content')!.textContent });
        } finally { layout.dispose(); }
        if (save < 2) { const blob = await exportDocx(doc); files.push([...new Uint8Array(await blob.arrayBuffer())]); const loaded = await importFile(new File([blob], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      groups.push({ index, original, states, files });
    }
    return groups;
  }, { bytes: ['ordinary', 'nested'].map(depth => [...samples[`${mode}-${depth}-${format}`].bytes]), format });
  for (const group of result) for (const [i, bytes] of group.files.entries()) await fs.writeFile(info.outputPath(`hwp-gov-empty-${format}-${mode}-${group.index}-save${i+1}.docx`), Buffer.from(bytes));
  for (let save = 0; save < 3; save++) {
    const ordinary = result[0].states[save], nested = result[1].states[save];
    expect(nested.pages).toBe(ordinary.pages); expect(nested.height).toBeCloseTo(ordinary.height, 0);
    expect(nested.cellText).toBe(ordinary.cellText); expect(nested.paragraphs).toBe(mode === 'two' ? 2 : 1); expect(nested.breaks).toBe(ordinary.breaks);
    for (const state of [ordinary, nested]) { expect(state.cellText).toBe(result[0].states[0].cellText); expect(state.breaks).toBe(result[0].states[0].breaks); }
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'blank.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result[1].original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' }), paragraph = editor.locator('table').last().locator('td').last().locator('p').last();
  const before = await paragraph.textContent();
  await paragraph.fill('EDIT GROW '.repeat(30)); await expect(paragraph).toContainText('EDIT GROW');
  await expect.poll(() => paragraph.evaluate(p => p.getBoundingClientRect().height)).toBeGreaterThan(80);
  await page.keyboard.press('ControlOrMeta+z'); await expect(paragraph).toHaveText(before || '');
  await expect.poll(() => paragraph.evaluate(p => p.getBoundingClientRect().height)).toBeCloseTo(result[1].states[0].lastHeight, 0);
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(paragraph).toContainText('EDIT GROW');
});
