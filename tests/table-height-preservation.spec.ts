import { test, expect, type Locator } from '@playwright/test';

// Slide Studio 72409b8: a 72pt frame must not squeeze four 32pt rows.
// Word uses flowing tables; exercise its HTML/clipboard bridge, DOCX writer
// and print layout with the same contradictory outer and row dimensions.
const cases: { name: string; heights: number[]; merged?: boolean; multiline?: boolean }[] = [
  { name: 'four 32pt rows', heights: [32, 32, 32, 32] },
  { name: 'vertically merged cells', heights: [32, 32, 32, 32], merged: true },
  { name: 'unequal rows', heights: [24, 40, 28, 36] },
  { name: 'content taller than the requested minimum', heights: [12, 12, 12, 12], multiline: true },
];

async function geometry(root: Locator) {
  return root.evaluate(async element => {
    await document.fonts.ready;
    // OOXML line metrics are applied by the editor on its next animation frame.
    await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    const table = element.querySelector('table')!;
    const scale = element.getBoundingClientRect().width / (element as HTMLElement).offsetWidth;
    let overflow = 0;
    const walker = document.createTreeWalker(table, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode;
      if (!text.textContent?.trim()) continue;
      const cell = text.parentElement!.closest('td,th')!;
      const box = cell.getBoundingClientRect(), range = document.createRange();
      range.selectNodeContents(text);
      for (const rect of range.getClientRects()) overflow = Math.max(overflow,
        box.top - rect.top, rect.bottom - box.bottom, box.left - rect.left, rect.right - box.right);
    }
    return { height: table.getBoundingClientRect().height / scale,
      rows: [...table.rows].map(row => row.getBoundingClientRect().height / scale),
      overflow, text: table.textContent, zooms: [...element.querySelectorAll<HTMLElement>('[style]')].map(el => el.style.zoom).filter(Boolean) };
  });
}

for (const fixture of cases) {
  test(`${fixture.name} keep their height through paste, DOCX roundtrips and print`, async ({ page }, testInfo) => {
    await page.goto('/');
    const cell = (row: number, col: number, rowspan = 1) => `<td rowspan="${rowspan}" colwidth="180" style="padding:4px;font-family:Arial;font-size:12pt"><p style="margin:0;line-height:1.1">표 ${row + 1}-${col + 1}${fixture.multiline ? '<br>둘째 줄<br>마지막 줄' : ''}</p></td>`;
    const html = `<table style="height:72pt"><tbody>${fixture.heights.map((height, row) => `<tr style="height:${height}pt">${fixture.merged && row === 1 ? '' : cell(row, 0, fixture.merged && row === 0 ? 2 : 1)}${cell(row, 1)}</tr>`).join('')}</tbody></table>`;
    await page.locator('input[type=file]').first().setInputFiles({ name: 'empty.html', mimeType: 'text/html', buffer: Buffer.from('<p></p>') });
    const editor = page.locator('.tiptap');
    await editor.click();
    await page.evaluate(html => {
      const data = new DataTransfer(); data.setData('text/html', html);
      document.querySelector('.tiptap')!.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }));
    }, html);
    await expect(editor.locator('tr')).toHaveCount(4);
    const pasted = await geometry(editor);
    const expected = fixture.heights.map(height => height * 4 / 3);
    for (const [row, height] of pasted.rows.entries()) expect(height).toBeGreaterThanOrEqual(expected[row] - .5);
    expect(pasted.overflow).toBeLessThanOrEqual(1);
    if (!fixture.multiline) expect(pasted.height).toBeGreaterThanOrEqual(128 * 4 / 3 - .5);
    else expect(pasted.rows.every(height => height > 40)).toBe(true);

    const versions = await page.evaluate(async html => {
      const { importFile, exportDocx } = await import('/src/document-io.ts');
      const { parseDocumentHtml } = await import('/src/extensions.ts');
      const loaded = await importFile(new File([html], 'height.html', { type: 'text/html' }));
      let document = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      const versions = [document];
      for (let i = 0; i < 2; i++) {
        const reopened = await importFile(new File([await exportDocx(document)], 'height.docx'));
        document = { ...reopened.document, content: parseDocumentHtml(reopened.html!) };
        versions.push(document);
      }
      return versions;
    }, html);
    for (const doc of versions) {
      const table = doc.content.content!.find(node => node.type === 'table')!;
      for (const [row, node] of table.content!.entries()) expect(node.attrs!.rowHeight).toBeCloseTo(expected[row], 1);
    }
    await page.locator('input[type=file]').first().setInputFiles({ name: 'roundtrip.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(versions.at(-1))) });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(versions.at(-1)!.title);
    await expect(editor.locator('tr')).toHaveCount(4);
    // Loading keeps four rows on screen throughout the transition. Wait for
    // the new document's OOXML font measurements, rather than sampling its
    // temporary CSS multiplier before the layout plugins finish.
    await expect(editor.locator('p[data-word-line-rule=auto]').first()).toHaveAttribute('style', /line-height:\s*[\d.]+px/);
    const reopened = await geometry(editor);
    expect(reopened.text).toBe(pasted.text);
    expect(reopened.overflow).toBeLessThanOrEqual(1);
    for (const [row, height] of reopened.rows.entries()) {
      // HTML's em-based line spacing becomes OOXML's font-line multiplier;
      // content can grow after conversion, but must never be squeezed to fit.
      if (fixture.multiline) expect(height).toBeGreaterThanOrEqual(pasted.rows[row] - .5);
      else expect(height).toBeCloseTo(pasted.rows[row], 0);
    }
    if (fixture.merged) await expect(editor.locator('td').first()).toHaveAttribute('rowspan', '2');
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    await expect(page.locator('.print-page')).toHaveCount(1);
    const print = await geometry(page.locator('.print-content'));
    expect(print.text).toBe(pasted.text);
    expect(print.overflow).toBeLessThanOrEqual(1);
    expect(print.zooms).toEqual([]);
    for (const [row, height] of print.rows.entries()) expect(height).toBeCloseTo(reopened.rows[row], 0);
    await testInfo.attach('height-measurements', { body: JSON.stringify({ expected, pasted, reopened, print }, null, 2), contentType: 'application/json' });
  });
}
