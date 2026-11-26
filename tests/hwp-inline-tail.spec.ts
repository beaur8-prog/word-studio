import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-inline-tail-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-studio-hwp-inline-tail-')); files = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });
for (const format of ['hwp', 'hwpx']) test(`${format}: text on an inline table's final line stays beside it through two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp, exportEditableHwp } = await import('/src/hwp/editing.ts');
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
          const paragraphs = [...layout.element.querySelectorAll('.print-content p')];
          const suffix = paragraphs.find(p => p.textContent === 'SUFFIX')!;
          const body = paragraphs.find(p => p.textContent === 'TABLE BODY')!.closest('table')!;
          const after = paragraphs.find(p => p.textContent === 'AFTER')!;
          const s = suffix.getBoundingClientRect(), t = body.getBoundingClientRect();
          const reserve = Number(suffix.closest('table')?.getAttribute('data-hwp-inline-table-bottom'));
          states.push({ pages: layout.pages, text: layout.element.textContent, reserve,
            save, x: s.left, tableRight: t.right, top: s.top, tableTop: t.top, bottom: s.bottom, tableBottom: t.bottom, beside: s.left >= t.right - 1, inside: s.top >= t.top && s.bottom <= t.bottom + 1,
            after: after.getBoundingClientRect().top >= t.bottom - 1, height: t.height });
        } finally { layout.dispose(); }
        if (save < 2) { const result = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...result.document, content: parseDocumentHtml(result.html!) }; }
      }
      const { blob } = await exportEditableHwp(doc, format);
      return { states, native: [...new Uint8Array(await blob.arrayBuffer())] };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of states.states) {
    expect(state.reserve).toBe(140);
    expect(state.pages).toBe(1); expect(state.text).toMatch(/BEFORE.*TABLE BODY.*SUFFIX.*AFTER/s);
    expect(state.beside, JSON.stringify(states)).toBe(true); expect(state.inside).toBe(true); expect(state.after).toBe(true);
    expect(state.height).toBeCloseTo(states.states[0].height, 0);
  }
  const { default:init, HwpDocument } = await import('../node_modules/@rhwp/core/rhwp.js');
  await init({ module_or_path: await fs.readFile('./node_modules/@rhwp/core/rhwp_bg.wasm') });
  const native = new HwpDocument(new Uint8Array(states.native));
  try {
    const wrapper = JSON.parse(native.getPageControlLayout(0)).controls.find(control =>
      control.type === 'table' && control.rowCount === 1 && control.colCount === 2 && !control.cellPath?.length);
    expect(wrapper).toBeTruthy();
    const { outerBottom } = JSON.parse(native.getTableProperties(wrapper.secIdx, wrapper.paraIdx, wrapper.controlIdx));
    expect(outerBottom).toBe(140);
  } finally { native.free(); }
});
