import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

async function fixture(page: any, long = false, constrained = false) {
  return page.evaluate(async ({ long, constrained }: { long: boolean; constrained: boolean }) => {
    const { defaultPage } = await import('/src/types.ts');
    const { exportDocx } = await import('/src/document-io.ts');
    const run = (text: string) => ({ type: 'paragraph', attrs: { fontFamily: 'Arial', fontSize: '12pt', lineSpacing: '1', wordLineRule: 'auto' }, content: [{ type: 'text', text, marks: [{ type: 'textStyle', attrs: { fontFamily: 'Arial', fontSize: '12pt' } }] }] });
    const cell = (text: string, width: number) => ({ type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: [width], wordCellWidthAuto: constrained ? null : true, paddingLeft: '6px', paddingRight: '6px', paddingTop: '0px', paddingBottom: '0px' }, content: [run(text)] });
    const table = (layout: string) => ({ type: 'table', attrs: { wordTableLayout: layout, wordTableWidthAuto: true, tableSpacing: '8px' }, content: [
      { type: 'tableRow', content: [cell(long ? 'This text is deliberately much longer than the available document width and therefore needs real authored line wrapping. '.repeat(5) : 'Content width', 70), cell('Second cell', 200)] },
      { type: 'tableRow', content: [cell('A', 70), cell('B', 200)] },
    ] });
    const doc: any = { format: 'word-studio', version: 1, title: 'AutoFit regression', updatedAt: new Date().toISOString(), page: defaultPage, comments: [], content: { type: 'doc', content: [table('autofit'), run('Between tables'), table('fixed')] } };
    return [...new Uint8Array(await (await exportDocx(doc)).arrayBuffer())];
  }, { long, constrained });
}

async function open(page: any, bytes: number[]) {
  await page.locator('input[type=file]').first().setInputFiles({ name: 'autofit.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor.locator('table')).toHaveCount(2);
  await page.evaluate(() => document.fonts.ready);
  return editor;
}

test('imported autofit grows a deficient preferred column, keeps fixed grids and source model unchanged across zoom and print', async ({ page }) => {
  await page.goto('/'); const editor = await open(page, await fixture(page));
  const model = await editor.evaluate((el: any) => el.editor.getJSON());
  const measure = () => editor.evaluate((root: any) => {
    return [...root.querySelectorAll('table')].map((t: any) => {
      const p = t.rows[0].cells[0].querySelector('p')!, range = document.createRange(); range.selectNodeContents(p);
      const rects = [...range.getClientRects()].filter(r => r.width > 0);
      return { layout: t.dataset.wordTableLayout, width: parseFloat(t.querySelector('col').style.width), lines: [...new Set(rects.map(r => Math.round(r.y * 10) / 10))].length, modelWidth: t.rows[0].cells[0].getAttribute('colwidth') };
    });
  });
  await expect.poll(async () => (await measure())[0].lines).toBe(1);
  const baseline = await measure();
  expect(baseline[0].layout).toBe('autofit'); expect(baseline[0].width).toBeGreaterThan(70);
  expect(baseline[1].layout).toBe('fixed'); expect(baseline[1].width).toBe(70); expect(baseline[1].lines).toBeGreaterThan(1);
  for (const zoom of [50, 125, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(measure).toEqual(baseline);
    expect(await editor.evaluate((el: any) => el.editor.getJSON())).toEqual(model);
  }
  await page.getByRole('button', { name: '보기', exact: true }).click(); await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-content table')).toHaveCount(2);
  const widths = await page.locator('.print-content table').evaluateAll(tables => tables.map(t => parseFloat((t.querySelector('col') as HTMLElement).style.width)));
  expect(widths).toEqual(baseline.map((t: any) => t.width));
});

test('autofit leaves oversized wrapping content inside its existing bounded grid', async ({ page }) => {
  await page.goto('/'); const editor = await open(page, await fixture(page, true));
  await expect.poll(() => editor.locator('table').first().locator('col').first().evaluate(el => parseFloat((el as HTMLElement).style.width))).toBe(70);
  const widths = await editor.locator('table').evaluateAll(tables => tables.map(t => t.getBoundingClientRect().width));
  for (const width of widths) expect(width).toBeLessThan(300);
});

test('measured autofit view does not write expanded columns into either DOCX save', async ({ page }, info) => {
  await page.goto('/'); const editor = await open(page, await fixture(page));
  await expect.poll(() => editor.locator('table').first().locator('col').first().evaluate(el => parseFloat((el as HTMLElement).style.width))).toBeGreaterThan(70);
  const result = await editor.evaluate(async (el: any) => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let doc: any = { format: 'word-studio', version: 1, title: 'Autofit preserved', updatedAt: new Date().toISOString(), page: defaultPage, comments: [], content: el.editor.getJSON() };
    const states = [], saves = [], widthTypes = [];
    const read = (n: any): any => ({ type: n.type, ...(n.type === 'table' ? { layout: n.attrs.wordTableLayout, autoWidth: n.attrs.wordTableWidthAuto } : {}), ...(n.type === 'tableCell' ? { columns: n.attrs.colwidth, autoWidth: n.attrs.wordCellWidthAuto, colspan: n.attrs.colspan, rowspan: n.attrs.rowspan } : {}), ...(n.text ? { text: n.text } : {}), ...(n.content ? { content: n.content.map(read) } : {}) });
    states.push(read(doc.content));
    for (let i = 0; i < 2; i++) {
      const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]);
      const zip = await JSZip.loadAsync(blob), xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      widthTypes.push(['tblW', 'tcW'].map(name => [...xml.getElementsByTagNameNS(W, name)].map(el => [el.getAttributeNS(W, 'type'), el.getAttributeNS(W, 'w')])));
      const imported = await importFile(new File([blob], 'again.docx')); doc = { ...imported.document, content: parseDocumentHtml(imported.html!) }; states.push(read(doc.content));
    }
    return { states, saves, widthTypes };
  });
  for (const [tables, cells] of result.widthTypes) { expect(tables).toEqual(Array(2).fill(['auto', '0'])); expect(cells).toEqual(Array(8).fill(['auto', '0'])); }
  expect(result.states[1]).toEqual(result.states[0]); expect(result.states[2]).toEqual(result.states[0]);
  for (const [i, bytes] of result.saves.entries()) await writeFile(info.outputPath(`textbox-save-autofit-source-grid-${i + 1}.docx`), Buffer.from(bytes));
});

test('a multi-page autofit table uses the complete table widths on every continuation', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const cell = (text: string, width: number) => `<td colwidth="${width}" data-word-cell-width-auto="true" style="padding:2px 6px"><p style="font:12pt Arial;line-height:20px;margin:0">${text}</p></td>`;
    const rows = Array.from({ length: 70 }, (_, i) => `<tr>${cell(i === 69 ? 'ROW_69_THE_LATE_WIDEST_CELL_END' : `ROW_${i}_END`, 70)}${cell('B', 200)}</tr>`).join('');
    const layout = await createPrintLayout(`<table data-word-table-layout="autofit" data-word-table-width-auto="true"><colgroup><col style="width:70px"><col style="width:200px"></colgroup><tbody>${rows}</tbody></table>`, defaultPage);
    try {
      return { pages: layout.pages, widths: [...layout.element.querySelectorAll('.print-content table')].map(t => [...t.querySelectorAll<HTMLElement>(':scope > colgroup > col')].map(c => parseFloat(c.style.width))), rows: [...layout.element.querySelectorAll('.print-content td:first-child p')].map(p => p.textContent) };
    } finally { layout.dispose(); }
  });
  expect(result.pages).toBeGreaterThan(1);
  expect(result.widths.length).toBeGreaterThan(1);
  expect(result.widths[0][0]).toBeGreaterThan(70);
  for (const widths of result.widths) expect(widths).toEqual(result.widths[0]);
  expect(result.rows).toEqual(Array.from({ length: 70 }, (_, i) => i === 69 ? 'ROW_69_THE_LATE_WIDEST_CELL_END' : `ROW_${i}_END`));
});

test('autofit does not use the whole paper width for a table in one of two columns', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const layout = await createPrintLayout('<div data-column-count="2" data-column-space="20" style="column-gap:20px"><table data-word-table-layout="autofit" data-word-table-width-auto="true"><colgroup><col style="width:70px"><col style="width:200px"></colgroup><tbody><tr><td colwidth="70" data-word-cell-width-auto="true" style="padding:0 6px"><p style="font:12pt Arial;line-height:20px;margin:0">A fairly wide content phrase</p></td><td colwidth="200" data-word-cell-width-auto="true" style="padding:0 6px"><p style="font:12pt Arial;line-height:20px;margin:0">B</p></td></tr></tbody></table></div>', defaultPage);
    try { const t = layout.element.querySelector('.print-content table')!; return { widths: [...t.getClientRects()].map(r => r.width), first: parseFloat((t.querySelector('col') as HTMLElement).style.width), pages: layout.pages, text: t.textContent }; }
    finally { layout.dispose(); }
  });
  expect(result.first).toBe(70);
  for (const width of result.widths) expect(width).toBeLessThan(302);
  expect(result.pages).toBe(1);
  expect(result.text).toBe('A fairly wide content phraseB');
});


test('autofit retains explicit cell width preferences instead of treating them as an unconstrained grid', async ({ page }) => {
  await page.goto('/'); const editor = await open(page, await fixture(page, false, true));
  await expect.poll(() => editor.locator('table').first().locator('col').first().evaluate(el => parseFloat((el as HTMLElement).style.width))).toBe(70);
  await expect(editor.locator('[data-word-cell-width-auto]')).toHaveCount(0);
  await page.getByRole('button', { name: '보기', exact: true }).click(); await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  expect(await page.locator('.print-content table').first().locator('col').first().evaluate(el => parseFloat((el as HTMLElement).style.width))).toBe(70);
});
