import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-nested-borders-fixture.cjs');
let directory: string, samples: Record<string, { file: string; cells: Record<string, any>[] }>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-hwp-nested-borders-')); samples = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

for (const format of ['hwp', 'hwpx']) test(`${format}: nested table borders use native edge styles through two DOCX saves and editing`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ bytes, format }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts'), { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'), { createPrintLayout } = await import('/src/print-layout.ts');
    const session = new HwpReaderSession();
    try {
      await session.open(new File([new Uint8Array(bytes)], 'borders.' + format)); let doc = importEditableHwp(await session.editable(), 'borders.' + format);
      const first = doc, states = [];
      for (let save = 0; save < 3; save++) {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const nested = layout.element.querySelector('table table')!;
          states.push({ pages: layout.pages, text: layout.element.querySelector('.print-content')!.textContent, cells: [...nested.querySelectorAll<HTMLElement>('td')].map(cell => ['Top', 'Right', 'Bottom', 'Left'].map(side => {
            const css = getComputedStyle(cell); return { style: css[`border${side}Style` as 'borderTopStyle'], width: Number.parseFloat(css[`border${side}Width` as 'borderTopWidth']), color: css[`border${side}Color` as 'borderTopColor'] };
          })) });
        } finally { layout.dispose(); }
        if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
      }
      return { first, states };
    } finally { session.close(); }
  }, { bytes: [...await fs.readFile(samples[format].file)], format });
  const expected = [
    [{ style: 'solid', mm: .2, color: 'rgb(192, 0, 0)' }, { style: 'dashed', mm: .5, color: 'rgb(0, 112, 192)' }, { style: 'double', mm: .3, color: 'rgb(0, 128, 0)' }, { style: 'none', mm: 0 }],
    Array.from({ length: 4 }, () => ({ style: 'solid', mm: .2, color: 'rgb(0, 0, 0)' })),
  ];
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.text).toBe('OUTERNESTED LABELNESTED VALUETAIL'); expect(state.cells).toHaveLength(2);
    state.cells.forEach((edges, cell) => edges.forEach((edge, side) => {
      const source = expected[cell][side];
      expect(edge.style).toBe(source.style);
      // Browser layout quantizes physical border widths to device pixels.
      expect(Math.abs(edge.width - source.mm * 96 / 25.4)).toBeLessThanOrEqual(1);
      if (source.color) expect(edge.color).toBe(source.color);
    }));
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'borders.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.first)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' }), label = editor.locator('table table p').filter({ hasText: /^NESTED LABEL$/ });
  for (const zoom of [100, 50, 125, 200, 100]) { const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change'); await expect(label.locator('..')).toHaveCSS('border-left-style', result.states[0].cells[0][3].style); }
  await label.fill('EDITED LABEL'); await page.keyboard.press('ControlOrMeta+z'); await expect(label).toHaveText('NESTED LABEL');
});
