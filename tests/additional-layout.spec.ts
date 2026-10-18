import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';

async function fixture() {
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('fixture')] }] })));
  const p = (text: string) => `<w:p><w:pPr><w:spacing w:line="249" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="DotumChe" w:eastAsia="DotumChe"/><w:sz w:val="18"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    ${p('서울･지역')}${p('서울 지역')}
    <w:tbl><w:tblPr><w:tblCellMar><w:top w:w="30"/><w:bottom w:w="30"/></w:tblCellMar></w:tblPr>
    <w:tblGrid><w:gridCol w:w="1350"/><w:gridCol w:w="30"/><w:gridCol w:w="1500"/></w:tblGrid>
    <w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="3"/><w:shd w:val="pct15" w:color="auto" w:fill="auto"/><w:tcBorders><w:bottom w:val="double" w:sz="8"/></w:tcBorders></w:tcPr>${p('Merged grid')}</w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${p('Left')}</w:tc><w:tc><w:tcPr/>${p('Right')}</w:tc></w:tr></w:tbl>
    <w:p><w:pPr><w:framePr w:w="5000" w:h="8000"/></w:pPr></w:p>
    <w:p><w:r><w:t>Signature </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="18"/></w:rPr><w:instrText> EQ \\o\\ac(○,</w:instrText></w:r><w:r><w:instrText>인)</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

test('narrow merged grids, percentage shading, double borders and overstrike fields survive DOCX roundtrips', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    // @ts-expect-error Vite source module.
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    // @ts-expect-error Vite source module.
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const loaded = await importFile(new File([new Uint8Array(bytes)], 'grid.docx'));
    let doc = { ...loaded.document, content: parseDocumentHtml(loaded.html) };
    const versions = [doc];
    for (let i = 0; i < 2; i++) {
      const again = await importFile(new File([await exportDocx(doc)], 'again.docx'));
      doc = { ...again.document, content: parseDocumentHtml(again.html) }; versions.push(doc);
    }
    return versions;
  }, await fixture());
  for (const doc of result) {
    expect(doc.content.content.map((n: any) => n.type)).toEqual(['paragraph', 'paragraph', 'table', 'paragraph']);
    const row = doc.content.content[2].content[0], cell = row.content[0];
    expect(cell.attrs.colwidth).toEqual([90, 2, 100]);
    expect(cell.attrs.backgroundColor).toBe('rgb(217, 217, 217)');
    expect(cell.attrs.borderBottom).toBe('3pt double rgb(0, 0, 0)');
    expect(row.attrs.rowHeight).toBe(48);
    const overlap = doc.content.content[3].content.find((n: any) => n.type === 'wordOverlap');
    expect(JSON.parse(overlap.attrs.parts)).toEqual(['○', '인']);
    expect(overlap.marks.find((m: any) => m.type === 'textStyle').attrs.fontSize).toBe('9pt');
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'grid.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result[0])) });
  const table = page.locator('.tiptap table'); await expect(table).toBeVisible();
  await expect(table.locator('col').nth(1)).toHaveCSS('width', '2px');
  expect((await table.boundingBox())!.width).toBeLessThan(196);
  await expect.poll(async () => parseFloat(await page.locator('.tiptap > p').nth(0).evaluate(e => getComputedStyle(e).lineHeight))).toBeCloseTo(19.2975, 1);
  await expect.poll(async () => parseFloat(await page.locator('.tiptap > p').nth(1).evaluate(e => getComputedStyle(e).lineHeight))).toBeCloseTo(16.185, 1);
  const overlap = page.locator('.tiptap .word-overlap'); await expect(overlap).toHaveText('○인');
  const centers = await overlap.locator('span').evaluateAll(els => els.map(e => { const r = e.getBoundingClientRect(); return r.x + r.width / 2; }));
  expect(Math.abs(centers[0] - centers[1])).toBeLessThan(.5);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page')).toHaveCount(1);
  expect((await page.locator('.print-content table').boundingBox())!.width).toBeLessThan(196);
  await expect(page.locator('.print-content .word-overlap')).toHaveText('○인');
  const beforePrint = await page.locator('.print-content table').boundingBox();
  await page.evaluate(() => document.body.classList.add('studio-printing'));
  await page.emulateMedia({ media: 'print' });
  await expect(page.locator('.print-content td').first()).toHaveCSS('position', 'static');
  const inPrint = await page.locator('.print-content table').boundingBox();
  expect(inPrint!.width).toBeCloseTo(beforePrint!.width, 1);
  expect(inPrint!.height).toBeCloseTo(beforePrint!.height, 1);
});

test('a minimum-height single-cell row moves intact when the remaining page is too short', async ({ page }) => {
  await page.goto('/');
  const html = `<p style="line-height:20px;margin:0">Opening${'<br>line'.repeat(33)}</p><table style="margin-bottom:0"><tr style="height:280px"><td style="padding:0"><p style="line-height:20px;margin:0">Notice beginning${'<br>line'.repeat(12)}<br>Notice end</p></td></tr></table>`;
  await page.locator('input[type=file]').first().setInputFiles({ name: 'minimum-row.html', mimeType: 'text/html', buffer: Buffer.from(html) });
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page')).toHaveCount(2);
  await expect(page.locator('.print-page').first()).not.toContainText('Notice');
  await expect(page.locator('.print-page').nth(1)).toContainText('Notice beginning');
  await expect(page.locator('.print-page').nth(1)).toContainText('Notice end');
});
