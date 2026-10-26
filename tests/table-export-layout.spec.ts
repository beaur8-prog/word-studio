import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
test('separate wide and narrow tables keep their grids and have a minimal Word paragraph between them', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async W => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const JSZip = (await import('/node_modules/.vite/deps/jszip.js')).default;
    const table = (a: number, b: number, label: string) => `<table><tr><td colwidth="${a}"><p>${label} LEFT</p></td><td colwidth="${b}"><p>${label} RIGHT</p></td></tr></table>`;
    const first = await importFile(new File([table(480, 320, 'WIDE') + table(180, 180, 'NARROW')], 'tables.html'));
    const doc = { ...first.document, page: { ...first.document.page, orientation: 'landscape' as const }, content: parseDocumentHtml(first.html!) };
    const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const body = xml.getElementsByTagNameNS(W, 'body')[0];
    const elements = [...body.children];
    const tables = elements.filter(e => e.localName === 'tbl');
    const separator = elements[elements.indexOf(tables[0]) + 1];
    const again = await importFile(new File([blob], 'tables.docx'));
    const saved = parseDocumentHtml(again.html!);
    return { grids: tables.map(t => [...t.getElementsByTagNameNS(W, 'gridCol')].map(c => Number(c.getAttributeNS(W, 'w')))),
      separator: separator.localName, line: separator.getElementsByTagNameNS(W, 'spacing')[0]?.getAttributeNS(W, 'line'),
      tables: saved.content?.filter(n => n.type === 'table').length, text: again.html };
  }, W);
  expect(result.grids).toEqual([[7200, 4800], [2700, 2700]]);
  expect(result.separator).toBe('p'); expect(result.line).toBe('1'); expect(result.tables).toBe(2);
  expect(result.text).toContain('WIDE LEFT'); expect(result.text).toContain('NARROW RIGHT');
});

test('nested tables use the padded parent cell width, including the row beneath a vertical merge', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async W => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const JSZip = (await import('/node_modules/.vite/deps/jszip.js')).default;
    const nested = (label: string) => `<table><tr><td><p>${label}</p></td></tr></table>`;
    const source = `<table><tr><td colwidth="200" rowspan="2"><p>MERGED</p></td><td colwidth="300" style="padding:10px">${nested('TOP')}</td></tr><tr><td colwidth="300" style="padding:10px">${nested('BOTTOM')}</td></tr></table>`;
    const first = await importFile(new File([source], 'nested.html'));
    const doc = { ...first.document, content: parseDocumentHtml(first.html!) };
    const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const tables = [...xml.getElementsByTagNameNS(W, 'tbl')];
    const widths = tables.slice(1).map(t => Number([...t.children].find(e => e.localName === 'tblPr')!.getElementsByTagNameNS(W, 'tblW')[0].getAttributeNS(W, 'w')));
    const trailing = tables.slice(1).map(t => {
      const paragraph = t.nextElementSibling!;
      return [paragraph.localName, paragraph.getElementsByTagNameNS(W, 'spacing')[0]?.getAttributeNS(W, 'line')];
    });
    const again = await importFile(new File([blob], 'nested.docx'));
    return { widths, trailing, text: again.html };
  }, W);
  expect(result.widths).toEqual([4200, 4200]); // 300px cell minus 10px on either side.
  expect(result.trailing).toEqual([['p', '1'], ['p', '1']]);
  expect(result.text).toContain('TOP'); expect(result.text).toContain('BOTTOM'); expect(result.text).toContain('MERGED');
});

test('a percentage table in a two-column section uses the column as its containing width', async ({ page }) => {
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('fixture')] }] })));
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>
    <w:tbl><w:tblPr><w:tblW w:type="pct" w:w="2500"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
    <w:tr>${['LEFT', 'RIGHT'].map(text => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr></w:tbl>
    <w:sectPr><w:pgSz w:w="12000" w:h="16000"/><w:pgMar w:left="1000" w:right="1000" w:top="1000" w:bottom="1000"/>
    <w:cols w:num="2" w:space="600"/></w:sectPr></w:body></w:document>`);
  const bytes = [...await zip.generateAsync({ type: 'uint8array' })];
  await page.goto('/');
  const result = await page.evaluate(async data => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const first = await importFile(new File([new Uint8Array(data)], 'columns.docx'));
    const content = parseDocumentHtml(first.html!);
    const again = await importFile(new File([await exportDocx({ ...first.document, content })], 'again.docx'));
    const widths = (html: string) => [...new DOMParser().parseFromString(html, 'text/html').querySelectorAll('[colwidth]')]
      .reduce((sum, cell) => sum + Number(cell.getAttribute('colwidth')), 0);
    return { first: widths(first.html!), again: widths(again.html!), text: again.html };
  }, bytes);
  // (10000twip content - 600twip gap) / 2 columns * 50% / 15twip per px,
  // rounded to the editor's integer pixel widths (78px per cell).
  expect(result.first).toBe(156); expect(result.again).toBeCloseTo(result.first, 0);
  expect(result.text).toContain('LEFT'); expect(result.text).toContain('RIGHT');
});
