import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-overlay-fixture.cjs');
let directory: string, files: Record<string, string>, centeredFiles: Record<string, string>;
test.beforeAll(async () => {
  directory = await fs.mkdtemp(path.resolve('.cache/hwp-overlay-'));
  files = await fixture(directory);
  centeredFiles = await fixture(path.join(directory, 'centered'), 1);
});
test.afterAll(async () => { await fs.rm(directory, { recursive: true, force: true }); });
for (const format of ['hwp', 'hwpx']) test(`${format}: a text table and its overlapping picture stay together through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
    const { importFile, exportDocx } = await import('/src/document-io.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'overlay.' + format));
      let doc = importEditableHwp(await session.editable(), 'overlay.' + format); const states = [];
      for (let save = 0; save < 3; save++) {
        const html = generateHTML(doc.content, editorExtensions), layout = await createPrintLayout(html, doc.page);
        try {
          const table = layout.element.querySelector('table')!, img = table.querySelector('img')!;
          const t = table.getBoundingClientRect(), b = img?.getBoundingClientRect();
          states.push({ text: table.textContent, images: layout.element.querySelectorAll('img').length, tableImages: table.querySelectorAll('img').length,
            height: t.height, x: b && b.x - t.x, y: b && b.y - t.y, geometry: img?.dataset.imageFloating, pages: layout.pages });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return states;
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of states) {
    expect(state.text).toContain('EDITABLE OVERLAY 한글'); expect(state.text).toContain('SECOND LINE');
    expect(state.images).toBe(1); expect(state.tableImages).toBe(1); expect(state.height).toBeLessThan(350); expect(state.pages).toBe(1);
    expect(state.x).toBeGreaterThan(15); expect(state.x).toBeLessThan(25); expect(Math.abs(state.y!)).toBeLessThan(3);
    expect(JSON.parse(state.geometry!)).toMatchObject({ wrap: { type: 0 }, behindDocument: true });
  }
});

test('a sparse vertically centered cell keeps its background separate instead of shifting it with the text', async ({ page }) => {
  await page.goto('/');
  const state = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'centered.hwp'));
      const converted = await session.editable(), root = document.createElement('div'); root.innerHTML = converted.html;
      return { text: root.textContent, pictures: root.querySelectorAll('img').length, tablePictures: root.querySelectorAll('table img').length };
    } finally { session.close(); }
  }, [...await fs.readFile(centeredFiles.hwp)]);
  expect(state.text).toContain('EDITABLE OVERLAY 한글'); expect(state.text).toContain('SECOND LINE');
  expect(state.pictures).toBe(1); expect(state.tablePictures).toBe(0);
});
