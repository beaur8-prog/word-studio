import { test, expect } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: full-width and balanced column sections, explicit column/page breaks survive two edited saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async format => {
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Columns', comments: [], updatedAt: new Date().toISOString(),
      page: { size: 'Letter', orientation: 'landscape', margin: 20, header: 'KEEP HEADER', footer: '', pageNumbers: false },
      content: parseDocumentHtml('<div data-column-count="1" data-column-space="0" data-column-balance="false"><p>FULL WIDTH</p></div>' +
        '<div data-column-count="2" data-column-space="24" data-column-balance="true"><p>LEFT</p><div data-column-break="true"></div><p>RIGHT</p><div data-page-break="true"></div><p>NEW PAGE</p></div><p>AFTER</p>') };
    const states = [];
    for (let round = 0; round < 2; round++) {
      const before = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      if (before.pages !== 2) throw Error('Expected two source pages before save');
      before.dispose();
      const exported = await exportEditableHwp(doc, format), session = new HwpReaderSession();
      try { await session.open(new File([exported.blob], 'columns.' + format)); doc = importEditableHwp(await session.editable(), 'columns.' + format); }
      finally { session.close(); }
      const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
      const layout = await createPrintLayout(root.innerHTML, doc.page);
      try { states.push({ page: doc.page, warnings: exported.warnings, text: [...root.querySelectorAll('p')].map(p => p.textContent), groups: [...root.querySelectorAll<HTMLElement>('[data-column-count]')].map(g => ({ count: g.dataset.columnCount, space: g.dataset.columnSpace, balance: g.dataset.columnBalance, text: g.textContent })), columnBreaks: root.querySelectorAll('[data-column-break]').length, pages: [...layout.element.querySelectorAll('.print-content')].map(p => p.textContent) }); }
      finally { layout.dispose(); }
    }
    return states;
  }, format);
  for (const state of states) {
    expect(state.page.size).toBe('Letter'); expect(state.page.orientation).toBe('landscape');
    expect(state.page.header).toContain('KEEP HEADER');
    expect(state.text).toEqual(['FULL WIDTH', 'LEFT', 'RIGHT', 'NEW PAGE', 'AFTER']);
    expect(state.columnBreaks).toBe(1);
    expect(state.groups.some(g => g.count === '1' && g.text === 'FULL WIDTH')).toBe(true);
    expect(state.groups.filter(g => g.count === '2').every(g => g.space === '24' && g.balance === 'true')).toBe(true);
    expect(state.pages).toEqual(['FULL WIDTHLEFTRIGHT', 'NEW PAGEAFTER']);
    expect(state.warnings.join('\n')).not.toMatch(/다단 설정은 첫|내용 손실/);
  }
});
