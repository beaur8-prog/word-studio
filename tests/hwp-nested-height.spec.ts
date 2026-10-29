import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const createFixture = createRequire(import.meta.url)('./hwp-nested-height-fixture.cjs');
let directory: string, samples: Record<string, { file: string; native: any[]; spacing: number }>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-nested-height-')); samples = await createFixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });
for (const mode of ['rows', 'merged', 'deep', 'grow', 'gap-zero', 'gap-large']) for (const format of ['hwp', 'hwpx']) test(`${format}: ${mode} nested table keeps minimum row heights and source spacing through two DOCX saves`, async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'height.' + format));
      let doc = importEditableHwp(await session.editable(), 'height.' + format); const original = doc, states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const table = [...layout.element.querySelectorAll<HTMLTableElement>('.print-content table')].filter(t => t.textContent!.includes('TALL B')).at(-1)!;
          states.push({ pages: layout.pages, spacing: parseFloat(getComputedStyle(table).marginBottom), rows: [...table.rows].map(r => r.getBoundingClientRect().height),
            cells: [...table.rows].flatMap(r => [...r.cells]).map(c => ({ text: c.textContent, height: c.getBoundingClientRect().height, span: c.rowSpan })), text: table.textContent });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { original, states };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(samples[mode + '-' + format].file)], format });
  await fs.writeFile(info.outputPath('native-and-converted.json'), JSON.stringify({ native: samples[mode + '-' + format].native, states: result.states }, null, 2));
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.rows).toHaveLength(2);
    expect(Math.abs(state.spacing - samples[mode + '-' + format].spacing)).toBeLessThan(.08);
    expect(state.rows[0]).toBeGreaterThanOrEqual(95); expect(state.rows[1]).toBeGreaterThanOrEqual(71);
    expect(state.text).toBe(result.states[0].text);
    if (mode === 'grow') expect(Math.abs(state.rows[0] - samples[mode + '-' + format].native[0].h)).toBeLessThan(.6);
    else { expect(state.rows[0]).toBeLessThan(98); expect(state.rows[1]).toBeLessThan(74); }
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'height.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }), table = editor.locator('table').filter({ hasText: 'TALL B' }).last();
  for (const zoom of [50, 100, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    const layout = await table.evaluate((t, zoom) => ({ rows: [...(t as HTMLTableElement).rows].map(r => r.getBoundingClientRect().height / (zoom / 100)),
      spacing: parseFloat(getComputedStyle(t.parentElement!).marginBottom) }), zoom);
    expect(layout.rows[0]).toBeGreaterThanOrEqual(95); expect(layout.rows[1]).toBeGreaterThanOrEqual(71);
    if (mode === 'grow') expect(Math.abs(layout.rows[0] - samples[mode + '-' + format].native[0].h)).toBeLessThan(.6);
    expect(Math.abs(layout.spacing - samples[mode + '-' + format].spacing)).toBeLessThan(.08);
  }
  const p = table.locator('p').filter({ hasText: /^TALL B$/ }); await p.fill('EDITED'); await page.keyboard.press('ControlOrMeta+z'); await expect(p).toHaveText('TALL B');
});
