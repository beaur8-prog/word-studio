import { test, expect } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) for (const firstPage of [false, true]) test(`${format}, first=${firstPage}: formatted odd/even stories and live PAGE/NUMPAGES remain editable after two saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async ({ format, firstPage }) => {
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Furniture', comments: [], updatedAt: new Date().toISOString(),
      page: { size: 'A4', orientation: 'portrait', margin: 30, header: '', footer: '', headerDistance: 5, footerDistance: 5, pageNumbers: false, differentOddEven: true,
        headerHtml: `<p style="font-size:9pt;line-height:12px;margin:0"><strong>ODD</strong> <span data-page-number="PAGE">9</span>/<span data-page-number="NUMPAGES">9</span></p><table style="width:120px"><tr><td style="padding:0"><p style="font-size:9pt;line-height:12px;margin:0">HEADER CELL<img data-inline-image="true" src="data:image/png;base64,${png}" width="10" height="10"></p></td></tr></table>`,
        footerHtml: '<p style="font-size:9pt;line-height:12px;margin:0">ODD FOOTER</p>', evenHeaderHtml: '<p style="font-size:9pt;line-height:12px;margin:0"><em>EVEN</em> <span data-page-number="PAGE">9</span></p>', evenFooterHtml: '<p style="font-size:9pt;line-height:12px;margin:0">EVEN FOOTER</p>', evenPageNumbers: false },
      content: parseDocumentHtml('<p>BODY ONE</p><div data-page-break="true"></div><p>BODY TWO</p><div data-page-break="true"></div><p>BODY THREE</p>') };
    if (firstPage) Object.assign(doc.page, { differentFirstPage: true, firstHeaderHtml: '<p style="font-size:9pt;line-height:12px;margin:0">FIRST <span data-page-number="PAGE">9</span></p>', firstFooter: 'FIRST FOOTER', firstPageNumbers: false });
    const states = [];
    for (let round = 0; round < 2; round++) {
      const exported = await exportEditableHwp(doc, format), session = new HwpReaderSession();
      try { await session.open(new File([exported.blob], 'furniture.' + format)); doc = importEditableHwp(await session.editable(), 'furniture.' + format); }
      finally { session.close(); }
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try { states.push({ warnings: exported.warnings, pages: [...layout.element.querySelectorAll('.print-page')].map(sheet => {
        const header = sheet.querySelector('.print-header')!;
        return { body: sheet.querySelector('.print-content')!.textContent, header: header.textContent?.replace(/\s+/g, ' ').trim(), fields: [...header.querySelectorAll('[data-page-number]')].map(f => f.textContent), bold: !!header.querySelector('strong,b,[style*="font-weight: bold"],[style*="font-weight:bold"]'), italic: !!header.querySelector('em,i,[style*="font-style: italic"],[style*="font-style:italic"]'), tables: header.querySelectorAll('table').length, images: header.querySelectorAll('img').length, footer: sheet.querySelector('.print-footer')!.textContent?.trim() };
      }) }); } finally { layout.dispose(); }
    }
    return states;
  }, { format, firstPage });
  for (const state of states) {
    expect(state.pages.map(p => p.body)).toEqual(['BODY ONE', 'BODY TWO', 'BODY THREE']);
    expect(state.pages.map(p => p.fields)).toEqual([firstPage ? ['1'] : ['1', '3'], ['2'], ['3', '3']]);
    expect(state.pages.map(p => p.footer)).toEqual([firstPage ? 'FIRST FOOTER' : 'ODD FOOTER', 'EVEN FOOTER', 'ODD FOOTER']);
    expect(state.pages[2]).toMatchObject({ bold: true, tables: 1, images: 1 });
    expect(state.pages[1]).toMatchObject({ italic: true, tables: 0, images: 0 });
    expect(state.pages[2].header).toContain('HEADER CELL');
    if (firstPage) expect(state.pages[0].header).toBe('FIRST 1');
    expect(state.warnings.join('\n')).not.toMatch(/내용 손실|기본 텍스트만 저장|자동 쪽 번호.*유지되지/);
  }
});

test('pretty-printed rich footer HTML has the same height and centered table position as compact HTML', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const footer = '<table style="margin-left:auto;margin-right:auto;width:90px"><tr><td style="padding:0"><p style="font-size:9pt;line-height:12px">ONE</p><p style="font-size:9pt;line-height:12px">TWO</p></td></tr></table>';
    const out = [];
    for (const html of [footer, footer.replaceAll('><', '>\n  <')]) {
      const layout = await createPrintLayout('<p>BODY</p>', { size: 'A4', orientation: 'portrait', margin: 30, footerDistance: 10, header: '', footer: '', footerHtml: html, pageNumbers: false });
      try {
        const sheet = layout.element.querySelector('.print-page')!.getBoundingClientRect(), table = layout.element.querySelector('.print-footer table')!.getBoundingClientRect();
        out.push({ height: table.height, center: table.left + table.width / 2 - (sheet.left + sheet.width / 2), bottom: sheet.bottom - table.bottom });
      } finally { layout.dispose(); }
    }
    return out;
  });
  expect(states[0].height).toBeLessThan(27);
  expect(states[1]).toEqual(states[0]);
  expect(Math.abs(states[0].center)).toBeLessThan(.1);
  expect(Math.abs(states[0].bottom - 10 * 96 / 25.4)).toBeLessThan(.1);
});
