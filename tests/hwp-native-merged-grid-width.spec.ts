import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, test } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) {
  test(`${format}: a complete merged grid keeps every cell inside its physical paper`, async ({ page }) => {
    await page.goto('/');
    const outputs = await page.evaluate(async format => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { defaultPage } = await import('/src/types.ts');
      const sized = (text, widths, columns = 1, rows = 1) => `<td colspan="${columns}" rowspan="${rows}" colwidth="${widths.join(',')}" style="width:${widths.reduce((a, b) => a + b, 0)}px;height:${rows * 32}px"><p>${text}</p></td>`;
      const row = html => `<tr style="height:32px">${html}</tr>`;
      const body = inconsistent => '<p>Before</p><table>' +
        row(sized('First', [80, 60], 2) + sized('Second', inconsistent ? [18, 21] : [18, 22], 2)) +
        row(sized('Third', [80]) + sized('Fourth', [60]) + sized('Fifth', [18, 22], 2, 2)) +
        row(sized('Sixth', [80]) + sized('Last', [60])) +
        '</table><p>After</p>';
      const files = [];
      const reader = new HwpReaderSession();
      try {
        for (const inconsistent of [false, true]) {
          const saved = await reader.export(body(inconsistent),
            { ...defaultPage, size: 'Custom', width: 110, height: 160, margin: 15 }, format);
          files.push([...saved.bytes]);
        }
      } finally { reader.close(); }
      return files;
    }, format);
    const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
    const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
    await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
    for (const [index, bytes] of outputs.entries()) {
      const document = new HwpDocument(Uint8Array.from(bytes));
      try {
        expect(document.pageCount()).toBe(1);
        const paper = JSON.parse(document.getPageInfo(0));
        const controls = JSON.parse(document.getPageControlLayout(0)).controls;
        const table = controls.find(c => c.type === 'table');
        expect(table.rowCount).toBe(3);
        expect(table.colCount).toBe(4);
        const cells = JSON.parse(document.getTableCellBboxes(table.secIdx, table.paraIdx, table.controlIdx));
        expect(cells).toHaveLength(7);
        const text = new Map<string, string>();
        for (const run of JSON.parse(document.getPageTextLayout(0)).runs) {
          const key = `${run.paraIdx}/${run.parentParaIdx}/${run.cellIdx}`;
          text.set(key, (text.get(key) ?? '') + run.text);
        }
        const words = [...text.values()];
        for (const label of ['Before', 'First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth', 'Last', 'After'])
          expect(words.filter(value => value === label), label + ' in ' + JSON.stringify(words)).toHaveLength(1);
        if (index === 0) {
          expect(table.w).toBeCloseTo(180, 0);
          expect(table.x + table.w).toBeLessThan(paper.width);
          expect(cells.map(c => c.w)).toEqual([140, 40, 80, 60, 40, 80, 60]);
        } else {
          // An inconsistent authored column cannot establish a verified grid.
          expect(table.x + table.w).toBeGreaterThan(paper.width);
        }
      } finally { document.free(); }
    }
  });
}
