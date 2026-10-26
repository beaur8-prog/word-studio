import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-cell-paragraph-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-cell-paragraphs-')); files = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });
for (const format of ['hwp', 'hwpx', 'hwp-nested', 'hwpx-nested']) test(`${format}: blank cell paragraph keeps its height and following alignment through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'cells.' + format));
      let doc = importEditableHwp(await session.editable(), 'cells.' + format);
      const states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const cell = layout.element.querySelector('.print-content td')!;
          const paragraphs = [...cell.querySelectorAll('p')];
          states.push({ pages: layout.pages, text: paragraphs.map(p => p.textContent),
            heights: paragraphs.map(p => p.getBoundingClientRect().height),
            alignment: paragraphs.map(p => getComputedStyle(p).textAlign), cellHeight: cell.getBoundingClientRect().height });
        } finally { layout.dispose(); }
        if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format: format.split('-')[0] });
  for (const state of states) {
    expect(state.pages).toBe(1);
    expect(state.text).toEqual(format.endsWith('-nested') ? ['BEFORE', 'NESTED', '', 'AFTER'] : ['BEFORE', '', 'AFTER']);
    expect(state.heights.at(-2)).toBeCloseTo(40, 1);
    expect(state.alignment.at(-1)).toBe('right');
    expect(state.cellHeight).toBeCloseTo(states[0].cellHeight, 0);
  }
});
