import { test, expect } from '@playwright/test';

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: edited section paper, margins and plain furniture survive two saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async format => {
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const initial = { size: 'A4', orientation: 'portrait', margin: 20, header: 'FIRST HEADER', footer: 'FIRST FOOTER', pageNumbers: false };
    const second = { ...initial, size: 'Letter', orientation: 'landscape', margins: { top: 25, right: 15, bottom: 23, left: 17 }, header: 'SECOND HEADER', footer: 'SECOND FOOTER' };
    const third = { ...initial, size: 'Custom', width: 180, height: 240, margin: 18, header: '', footer: '' };
    const section = settings => `<div data-section-page="${JSON.stringify(settings).replaceAll('"', '&quot;')}"></div>`;
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Sections', comments: [], updatedAt: new Date().toISOString(), page: initial,
      content: parseDocumentHtml('<p>FIRST BODY</p>' + section(second) + '<p>SECOND BODY</p><div data-page-break="true"></div><p>SECOND NEXT PAGE</p>' + section(third) + '<p>THIRD BODY</p>') };
    const states = [];
    for (let round = 0; round < 3; round++) {
      const html = generateHTML(doc.content, editorExtensions), layout = await createPrintLayout(html, doc.page);
      try { states.push({ pages: [...layout.element.querySelectorAll<HTMLElement>('.print-page')].map(sheet => ({ width: parseFloat(sheet.style.width) * 25.4 / 96, height: parseFloat(sheet.style.height) * 25.4 / 96, body: sheet.querySelector('.print-content')!.textContent, header: sheet.querySelector('.print-header')!.textContent, footer: sheet.querySelector('.print-footer')!.textContent, left: parseFloat((sheet.querySelector('.print-content') as HTMLElement).style.left) * 25.4 / 96 })) }); }
      finally { layout.dispose(); }
      if (round === 2) break;
      const out = await exportEditableHwp(doc, format), session = new HwpReaderSession();
      if (out.warnings.some(w => /첫 용지 설정으로 통합|내용 손실/.test(w))) throw Error(out.warnings.join('\n'));
      try { await session.open(new File([out.blob], 'sections.' + format)); doc = importEditableHwp(await session.editable(), 'sections.' + format); }
      finally { session.close(); }
    }
    return states;
  }, format);
  for (const state of states) {
    expect(state.pages.map(p => p.body)).toEqual(['FIRST BODY', 'SECOND BODY', 'SECOND NEXT PAGE', 'THIRD BODY']);
    expect(state.pages.map(p => p.header)).toEqual(['FIRST HEADER', 'SECOND HEADER', 'SECOND HEADER', '']);
    expect(state.pages.map(p => p.footer)).toEqual(['FIRST FOOTER', 'SECOND FOOTER', 'SECOND FOOTER', '']);
    for (const [index, [width, height, left]] of [[210, 297, 20], [279.4, 215.9, 17], [279.4, 215.9, 17], [180, 240, 18]].entries()) {
      expect(Math.abs(state.pages[index].width - width)).toBeLessThan(.02);
      expect(Math.abs(state.pages[index].height - height)).toBeLessThan(.02);
      expect(Math.abs(state.pages[index].left - left)).toBeLessThan(.02);
    }
  }
});
