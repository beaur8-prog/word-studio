import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-paragraph-table-fixture.cjs');
let directory: string, files: Record<string, any>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-paragraph-table-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) for (const kind of ['zero', 'offset', 'long']) test(`${format} ${kind}: paragraph table position follows preceding edits and margins, and offsets apply only once through two DOCX saves`, async ({ page }) => {
  await page.goto('/'); const file = files[`${format}-${kind}`];
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts'), { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'float.' + format)); const source = await session.editable();
      let doc = importEditableHwp(source, 'float.' + format); const states = [];
      const layoutState = async () => {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const tables = [...layout.element.querySelectorAll('.print-content > table')];
          const p = layout.element.querySelector('.print-content p')!;
          const normalized = document.createElement('div'); normalized.innerHTML = generateHTML(doc.content, editorExtensions);
          const tail = [...layout.element.querySelectorAll('.print-content > p')].find(p => p.textContent === 'AFTER')!;
          return { pages: layout.pages, tables: tables.map(t => ({ y: t.getBoundingClientRect().top - t.closest('.print-page')!.getBoundingClientRect().top, x: t.getBoundingClientRect().left - t.closest('.print-page')!.getBoundingClientRect().left })),
            gap: tables[0].getBoundingClientRect().top - p.getBoundingClientRect().bottom,
            bottom: tail.getBoundingClientRect().top - tables.at(-1)!.getBoundingClientRect().bottom,
            paragraphs: [...normalized.querySelectorAll(':scope > p')].map(p => p.textContent), text: normalized.textContent, anchor: normalized.querySelector('table')!.dataset.tableAnchor };
        } finally { layout.dispose(); }
      };
      const original = await layoutState();
      const margins = doc.page.margins || { top: doc.page.margin, right: doc.page.margin, bottom: doc.page.margin, left: doc.page.margin };
      doc.page = { ...doc.page, margins: { ...margins, left: margins.left + 5 } }; const first = doc;
      for (let save = 0; save < 3; save++) {
        states.push(await layoutState());
        if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
      }
      return { first, original, states, warnings: source.warnings };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(file.path)], format });
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify({ file, ...result, first: undefined }));
  expect(result.original.tables[0].x).toBeCloseTo(file.x, 0); expect(result.original.tables[0].y).toBeCloseTo(file.y, 0);
  expect(JSON.parse(result.original.anchor!)).toMatchObject({ reference: 'paragraph', horizontal: 'margin' });
  expect(result.warnings.some((w: string) => w.includes('떠 있는 표'))).toBe(false);
  for (const state of result.states) {
    expect(state.pages).toBe(result.original.pages); expect(state.gap).toBeCloseTo(file.offset, 0); expect(state.bottom).toBeCloseTo(file.bottom, 0);
    expect(state.tables[0].x).toBeCloseTo(file.x + 5 * 96 / 25.4, 0); expect(state.tables[0].y).toBeCloseTo(file.y, 0);
    for (const continuation of state.tables.slice(1)) expect(continuation.y).toBeCloseTo(result.first.page.margins!.top * 96 / 25.4, 0);
    expect(state.paragraphs).toEqual(result.original.paragraphs); expect(state.text).toBe(result.original.text);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'float.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' }), table = editor.locator(':scope > .tableWrapper');
  const before = editor.locator(':scope > p').filter({ hasText: /^BEFORE$/ });
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => table.evaluate((t, zoom) => (t.getBoundingClientRect().top - t.previousElementSibling!.getBoundingClientRect().bottom) / (zoom / 100), zoom)).toBeCloseTo(file.offset, 0);
  }
  const oldTop = await table.evaluate(t => t.getBoundingClientRect().top);
  await before.fill('BEFORE\nMORE');
  await expect.poll(() => table.evaluate(t => t.getBoundingClientRect().top - t.previousElementSibling!.getBoundingClientRect().bottom)).toBeCloseTo(file.offset, 0);
  await expect.poll(() => table.evaluate(t => t.getBoundingClientRect().top)).toBeCloseTo(oldTop + 20, 0);
  await page.keyboard.press('ControlOrMeta+z'); await expect(before).toHaveText('BEFORE');
  await expect.poll(() => table.evaluate(t => t.getBoundingClientRect().top)).toBeCloseTo(oldTop, 0);
});

test('paragraph table in a smaller second section uses its own body width through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const second = { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false };
    const anchor = JSON.stringify({ x: 8, y: 20, bottom: 4, reference: 'paragraph', horizontal: 'margin' }).replaceAll('"', '&quot;');
    const html = `<p>FIRST SECTION</p><div data-section-page="${JSON.stringify(second).replaceAll('"', '&quot;')}"></div><p style="line-height:20px">BEFORE</p><table data-table-anchor="${anchor}" style="width:450px"><colgroup><col style="width:450px"></colgroup><tr><td style="height:75px;padding:0"><p>SECOND TABLE</p></td></tr></table><p style="line-height:20px">AFTER</p>`;
    const loaded = await importFile(new File([html], 'sections.html'));
    let doc = { ...loaded.document, page: { ...defaultPage, margin: 20, pageNumbers: false }, content: parseDocumentHtml(html) };
    const states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const table = layout.element.querySelector('table')!, paper = table.closest('.print-page')!, previous = table.previousElementSibling!;
        states.push({ pages: layout.pages, x: table.getBoundingClientRect().left - paper.getBoundingClientRect().left,
          gap: table.getBoundingClientRect().top - previous.getBoundingClientRect().bottom, anchor: table.dataset.tableAnchor,
          text: [...layout.element.querySelectorAll('.print-content')].map(c => c.textContent).join('') });
      } finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    return states;
  });
  for (const state of states) {
    expect(state.pages).toBe(2); expect(state.gap).toBeCloseTo(20, 0); expect(state.x).toBeCloseTo(10 * 96 / 25.4 + 8, 0);
    expect(JSON.parse(state.anchor!)).toMatchObject({ reference: 'paragraph', horizontal: 'margin' });
    expect(state.text).toBe(states[0].text);
  }
});
