import { test, expect } from '@playwright/test';
import { Document, Footer, Packer, PageNumber, Paragraph, TextRun, LineRuleType } from 'docx';
import { writeFile } from 'node:fs/promises';

async function fixture(blank: boolean, exact: boolean) {
  const spacing = { before: 0, after: 200, line: exact ? 160 : 276, lineRule: exact ? LineRuleType.EXACT : LineRuleType.AUTO };
  const footer = new Footer({ children: [
    new Paragraph({ children: [new TextRun({ children: [PageNumber.CURRENT] })], spacing }),
    ...(blank ? [new Paragraph({ spacing })] : []),
  ] });
  return [...await Packer.toBuffer(new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
    sections: [{ properties: { page: { size: { width: 6000, height: 8000 }, margin: { top: 375, left: 375, right: 375, bottom: 300, footer: 300 } } }, footers: { default: footer }, children: [new Paragraph('BODY CONTENT PRESERVED')] }],
  }))];
}

for (const exact of [false, true]) test(`authored empty Word footer paragraph reserves a ${exact ? 'source exact' : 'source auto'} line and reduces body space`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async ({ plain, blank }) => {
    const { importFile } = await import('/src/document-io.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const states = [];
    for (const bytes of [plain, blank]) {
      const loaded = await importFile(new File([new Uint8Array(bytes)], 'footer.docx'));
      const layout = await createPrintLayout(loaded.html!, loaded.document.page);
      try {
        const footer = layout.element.querySelector<HTMLElement>('.print-footer .document-furniture')!;
        const ps = [...footer.querySelectorAll<HTMLElement>('p')];
        states.push({ paragraphs: ps.map(p => ({ text: p.textContent, height: p.getBoundingClientRect().height, lineHeight: getComputedStyle(p).lineHeight })), footerHeight: footer.getBoundingClientRect().height, bodyHeight: Number(layout.element.querySelector<HTMLElement>('.print-content')!.dataset.bodyHeight), bodyText: layout.element.querySelector('.print-content')!.textContent });
      } finally { layout.dispose(); }
    }
    return states;
  }, { plain: await fixture(false, exact), blank: await fixture(true, exact) });
  expect(result[0].paragraphs).toHaveLength(1);
  expect(result[1].paragraphs).toHaveLength(2);
  expect(result[1].paragraphs[1].text).toBe('');
  const first = result[1].paragraphs[0].height, empty = result[1].paragraphs[1].height;
  expect(empty).toBeGreaterThan(5);
  expect(Math.abs(empty - first)).toBeLessThan(.25);
  expect(result[1].footerHeight - result[0].footerHeight).toBeGreaterThan(first * .9);
  expect(result[1].bodyHeight).toBeLessThan(result[0].bodyHeight - first * .9);
  for (const state of result) expect(state.bodyText).toBe('BODY CONTENT PRESERVED');
});

test('empty footer paragraphs survive two DOCX saves and retain editor and print line height', async ({ page }, info) => {
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    let file = new File([new Uint8Array(bytes)], 'footer.docx');
    const states = [], saves = []; let doc: any;
    for (let i = 0; i < 3; i++) {
      const loaded = await importFile(file);
      doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      const html = new DOMParser().parseFromString(doc.page.footerHtml, 'text/html');
      states.push([...html.querySelectorAll('p')].map(p => ({ text: p.textContent, sourceLine: p.style.lineHeight })));
      if (i < 2) { const bytes = [...new Uint8Array(await (await exportDocx(doc)).arrayBuffer())]; saves.push(bytes); file = new File([new Uint8Array(bytes)], 'again.docx'); }
    }
    return { doc, states, saves };
  }, await fixture(true, false));
  for (const state of result.states) { expect(state).toHaveLength(2); expect(state[1].text).toBe(''); }
  expect(result.states[1]).toEqual(result.states[0]);
  expect(result.states[2]).toEqual(result.states[0]);
  for (const [i, bytes] of result.saves.entries()) await writeFile(info.outputPath(`textbox-save-empty-footer-${i + 1}.docx`), Buffer.from(bytes));
  await page.locator('input[type=file]').first().setInputFiles({ name: 'footer.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  await expect(editor).toContainText('BODY CONTENT PRESERVED');
  const empty = page.locator('.paper-footer .document-furniture p:empty').first();
  await expect(empty).toBeAttached();
  for (const zoom of [100, 50, 125]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => empty.evaluate(p => p.getBoundingClientRect().height)).toBeGreaterThan(5 * zoom / 100);
    await expect(editor).toHaveText('BODY CONTENT PRESERVED');
  }
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect.poll(() => page.locator('.print-footer .document-furniture p:empty').first().evaluate(p => p.getBoundingClientRect().height)).toBeGreaterThan(5);
  await expect(page.locator('.print-content')).toHaveText('BODY CONTENT PRESERVED');
});
