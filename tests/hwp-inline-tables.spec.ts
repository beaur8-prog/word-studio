import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-inline-tables-fixture.cjs');
let directory: string, files: { hwp: string; hwpx: string; x: number };
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-hwp-inline-tables-')); files = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });
for (const format of ['hwp', 'hwpx'] as const) test(`${format}: inline table prefix and suffix spaces add no lines, retain indentation and preserve a separate blank paragraph through DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'inline.' + format));
      let doc = importEditableHwp(await session.editable(), 'inline.' + format);
      const states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const table = layout.element.querySelector('table')!, paper = table.closest('.print-page')!;
          const before = table.previousElementSibling!, after = table.nextElementSibling!;
          const paragraphs = [...layout.element.querySelectorAll('.print-content > p')];
          const tail = paragraphs.find(p => p.textContent === 'TAIL')!;
          states.push({ pages: layout.pages, x: table.getBoundingClientRect().left - paper.getBoundingClientRect().left,
            prefix: before.textContent, suffix: after.textContent, prefixHeight: before.getBoundingClientRect().height, suffixHeight: after.getBoundingClientRect().height,
            separateBlank: tail.nextElementSibling!.textContent, blankHeight: tail.nextElementSibling!.getBoundingClientRect().height, text: layout.element.textContent });
        } finally { layout.dispose(); }
        if (save < 2) { const result = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...result.document, content: parseDocumentHtml(result.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of states) {
    expect(state.pages).toBe(1); expect(state.prefix).toBe('    '); expect(state.suffix).toBe('  ');
    expect(state.prefixHeight).toBeLessThan(1); expect(state.suffixHeight).toBeLessThan(1);
    expect(state.separateBlank).toBe(' '); expect(state.blankHeight).toBeCloseTo(30, 0);
    expect(state.x).toBeCloseTo(files.x, 0); expect(state.text).toContain('TABLE TEXT');
  }
});
