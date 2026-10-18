import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';

async function fixture(balance: boolean) {
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('fixture')] }] })));
  const run = '<w:r><w:rPr><w:rFonts w:ascii="DotumChe" w:eastAsia="DotumChe"/><w:sz w:val="20"/><w:spacing w:val="-6"/></w:rPr><w:t xml:space="preserve">  가나다  ABC  </w:t></w:r>';
  const p = (rule: string, line: number) => `<w:p><w:pPr><w:spacing w:line="${line}" w:lineRule="${rule}"/></w:pPr>${run}</w:p>`;
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    ${p('auto', 280)}${p('atLeast', 120)}${p('exact', 120)}
    <w:tbl><w:tblPr><w:tblCellMar><w:top w:w="30"/><w:bottom w:w="30"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>
    ${['atLeast', 'exact'].map(rule => `<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="${rule}"/></w:trPr><w:tc><w:tcPr/>${p('auto', 240)}</w:tc></w:tr>`).join('')}</w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  zip.file('word/settings.xml', `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:balanceSingleByteDoubleByteWidth w:val="${balance ? 1 : 0}"/></w:compat></w:settings>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

test('DOCX spaces, CJK tracking, line rules and padded minimum row heights survive roundtrip', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async cases => {
    // @ts-expect-error Vite source module.
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    // @ts-expect-error Vite source module.
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const output = [];
    for (const bytes of cases) {
      const loaded = await importFile(new File([new Uint8Array(bytes)], 'metrics.docx'));
      const doc = { ...loaded.document, content: parseDocumentHtml(loaded.html) };
      const again = await importFile(new File([await exportDocx(doc)], 'again.docx'));
      output.push({ doc, again: parseDocumentHtml(again.html) });
    }
    return output;
  }, [await fixture(true), await fixture(false)]);
  for (const [i, { doc, again }] of result.entries()) {
    const first = doc.content.content[0];
    expect(first.content.map((n: any) => n.text).join('')).toBe('  가나다  ABC  ');
    const korean = first.content.find((n: any) => n.text.includes('가나다'));
    expect(korean.marks.find((m: any) => m.type === 'textStyle').attrs.letterSpacing).toBe(i ? '-0.3pt' : '-0.6pt');
    expect(again.content[0].content.map((n: any) => n.text).join('')).toBe('  가나다  ABC  ');
    expect(again.content[1].attrs.wordLineRule).toBe('atLeast');
    expect(again.content[2].attrs.wordLineRule).toBe('exact');
    expect(doc.content.content[3].content.map((n: any) => n.attrs.rowHeight)).toEqual([44, 40]);
    expect(again.content[3].content.map((n: any) => n.attrs.rowHeight)).toEqual([44, 40]);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'metrics.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result[0].doc)) });
  const paragraphs = page.locator('.tiptap > p');
  await expect.poll(async () => parseFloat(await paragraphs.nth(0).evaluate(el => getComputedStyle(el).lineHeight))).toBeCloseTo(20.222, 1);
  await expect.poll(async () => parseFloat(await paragraphs.nth(1).evaluate(el => getComputedStyle(el).lineHeight))).toBeCloseTo(17.333, 1);
  await expect(paragraphs.nth(2)).toHaveCSS('line-height', '8px');
  // A view-only measurement must not replace the saved OOXML multiplier.
  const download = page.waitForEvent('download'); await page.keyboard.press('Control+s');
  const saved = JSON.parse(await (await import('node:fs/promises')).readFile((await (await download).path())!, 'utf8'));
  expect(Number(saved.content.content[0].attrs.lineSpacing)).toBeCloseTo(280 / 240, 4);
});

test('a form row can use the first sheet while its indivisible nested notice moves to the next', async ({ page }) => {
  await page.goto('/');
  const html = `<p style="line-height:20px;margin:0">Opening${'<br>line'.repeat(19)}</p><table style="margin-bottom:0"><tr><td style="padding:0"><p style="line-height:20px;margin:0">Form introduction${'<br>line'.repeat(6)}</p><table style="margin-bottom:0"><tr style="height:600px"><td style="padding:0"><p style="line-height:20px">Notice stays together</p></td></tr></table></td></tr></table>`;
  await page.locator('input[type=file]').first().setInputFiles({ name: 'nested.html', mimeType: 'text/html', buffer: Buffer.from(html) });
  const intro = page.locator('.tiptap p').filter({ hasText: 'Form introduction' });
  await expect(intro).toBeAttached();
  await expect(page.locator('.pagination-gap')).toHaveCount(1);
  const editor = await intro.evaluate(el => ({ y: el.getBoundingClientRect().top - el.closest('.tiptap')!.getBoundingClientRect().top }));
  expect(editor.y).toBeLessThan(1000);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page')).toHaveCount(2);
  await expect(page.locator('.print-page').first()).toContainText('Form introduction');
  await expect(page.locator('.print-page').first()).not.toContainText('Notice stays together');
  await expect(page.locator('.print-page').nth(1)).toContainText('Notice stays together');
  expect(await page.locator('.print-content [style]').evaluateAll(els => els.some(e => (e as HTMLElement).style.zoom))).toBe(false);
});
