import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-nested-grid-fixture.cjs');
let directory: string, sample: { files: Record<string, string>; geometry: Record<string, number>[][] };
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-nested-grid-')); sample = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });
for (const format of ['hwp', 'hwpx']) test(`${format}: merged and deeply nested table columns retain native widths through two DOCX saves and editing`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'grid.' + format));
      const converted = await session.editable(); let doc = importEditableHwp(converted, 'grid.' + format);
      const original = doc, states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const tables = [...layout.element.querySelectorAll<HTMLTableElement>('.print-content table')].filter(t => t.parentElement!.closest('table'));
          states.push({ pages: layout.pages, tables: tables.map(t => ({
            text: t.textContent, width: t.getBoundingClientRect().width, height: t.getBoundingClientRect().height,
            cells: [...t.rows].flatMap(r => [...r.cells]).map(c => ({ width: c.getBoundingClientRect().width, span: c.colSpan, alignment: getComputedStyle(c).verticalAlign })),
          })) });
        } finally { layout.dispose(); }
        if (save < 2) { const again = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...again.document, content: parseDocumentHtml(again.html!) }; }
      }
      return { original, states, warnings: converted.warnings };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(sample.files[format])], format });
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify(result.states));
  expect(result.warnings.join('\n')).not.toMatch(/셀.*복원|변환하지 못한 개체/);
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.tables).toHaveLength(2);
    state.tables.forEach((t, i) => {
      expect(t.cells).toHaveLength(sample.geometry[i].length);
      t.cells.forEach((c, j) => {
        const source = sample.geometry[i][j];
        expect(c.width).toBeCloseTo(source.width / 75, 0);
        expect(c.span).toBe(source.colSpan);
        expect(c.alignment).toBe(['top', 'middle', 'bottom'][source.verticalAlign]);
      });
      expect(t.text).toBe(result.states[0].tables[i].text);
      expect(t.height).toBeCloseTo(result.states[0].tables[i].height, 0);
    });
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'grid.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  await expect(editor.locator('table')).toHaveCount(3);
  for (const zoom of [50, 100, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => editor.locator('p').filter({ hasText: /^DEEP VALUE$/ }).evaluate((p, zoom) => Math.round(p.closest('td')!.getBoundingClientRect().width / (zoom / 100)), zoom)).toBe(240);
    await expect(page.locator('.document-stack')).toHaveAttribute('data-pages', '1');
  }
  const deep = editor.locator('p').filter({ hasText: /^DEEP VALUE$/ });
  await deep.fill('EDITED VALUE'); await page.keyboard.press('ControlOrMeta+z'); await expect(deep).toHaveText('DEEP VALUE');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor.locator('p').filter({ hasText: /^EDITED VALUE$/ })).toHaveText('EDITED VALUE');
  await page.keyboard.press('ControlOrMeta+z');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(result.original.content);
});
