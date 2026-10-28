import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-cell-colors-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-cell-colors-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: colored repeat headers retain their fills and text through page continuations, editing and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts'), { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts'), { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'colors.' + format));
      let doc = importEditableHwp(await session.editable(), 'colors.' + format); const first = doc, states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          states.push({ pages: layout.pages, headers: [...layout.element.querySelectorAll<HTMLTableRowElement>('tr')].filter(r => r.cells[0]?.tagName === 'TH').map(row => [...row.cells].map(cell => {
            const run = cell.querySelector('span') || cell, css = getComputedStyle(run);
            return { text: cell.textContent, background: getComputedStyle(cell).backgroundColor, color: css.color, weight: css.fontWeight };
          })), body: [0, 1, 2].map(i => [...layout.element.querySelectorAll<HTMLTableRowElement>('tr')].filter(r => r.cells[0]?.tagName === 'TD').map(r => r.cells[i]?.textContent || '').join('')) });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { first, states };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(files[format])], format });
  for (const state of result.states) {
    expect(state.pages).toBeGreaterThan(1); expect(state.pages).toBe(result.states[0].pages); expect(state.headers).toHaveLength(state.pages);
    for (const row of state.headers) {
      expect(row.map(c => c.text)).toEqual(['BLUE HEADER', 'YELLOW HEADER', 'CLEAR HEADER']);
      expect(row.map(c => c.background)).toEqual(['rgb(31, 78, 120)', 'rgb(255, 242, 204)', 'rgba(0, 0, 0, 0)']);
      expect(row.map(c => c.color)).toEqual(['rgb(255, 255, 255)', 'rgb(192, 0, 0)', 'rgb(0, 0, 0)']); expect(Number(row[0].weight)).toBeGreaterThanOrEqual(600);
    }
    expect(state.body).toEqual([3, 4, 5].map(cell => Array.from({ length: 85 }, (_, p) => `BODY ${cell} ${p}`).join('')));
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'colors.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const header = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('th').first();
  for (const zoom of [100, 50, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect(header).toHaveCSS('background-color', 'rgb(31, 78, 120)');
  }
  await header.locator('p').click(); await page.keyboard.press('End'); await page.keyboard.type(' EDIT'); await expect(header).toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(header).toHaveText('BLUE HEADER'); await expect(header).toHaveCSS('background-color', 'rgb(31, 78, 120)');
});
