import { test, expect } from '@playwright/test';
import JSZip from 'jszip';

async function layoutFixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Malgun Gothic"/><w:color w:val="000000"/><w:sz w:val="20"/></w:rPr><w:pPr><w:spacing w:before="60" w:after="100" w:line="280" w:lineRule="auto"/></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Base"/><w:pPr><w:ind w:left="120" w:right="60" w:hanging="30"/></w:pPr></w:style>
    <w:style w:type="character" w:styleId="Small"><w:name w:val="Small"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style>
  </w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:tbl><w:tblPr><w:tblInd w:w="150" w:type="dxa"/><w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="4500"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr/><w:p><w:pPr><w:pStyle w:val="Body"/><w:rPr><w:sz w:val="8"/><w:color w:val="0000FF"/></w:rPr></w:pPr></w:p>
    <w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Readable body</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr/><w:p><w:pPr><w:pStyle w:val="Body"/><w:rPr><w:color w:val="0000FF"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="Small"/><w:color w:val="008000"/><w:spacing w:val="-5"/></w:rPr><w:t>Explicit run</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Body"/><w:spacing w:line="300" w:lineRule="exact"/></w:pPr><w:r><w:t>Merged cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
    <w:p><w:r><w:t>Before break</w:t><w:br w:type="page"/><w:t>After break</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:bottom="1134" w:left="1134" w:right="1134"/></w:sectPr>
  </w:body></w:document>`);
  return [...await zip.generateAsync({ type: 'uint8array' })];
}

test('DOCX style inheritance, paragraph marks, table grid and spacing survive editing and export', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async bytes => {
    // @ts-expect-error Vite source module.
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    // @ts-expect-error Vite source module.
    const { editorExtensions } = await import('/src/extensions.ts');
    // @ts-expect-error Vite dependency.
    const { generateJSON } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const loaded = await importFile(new File([new Uint8Array(bytes)], 'style-layout.docx'));
    const doc = { ...loaded.document, content: generateJSON(loaded.html, editorExtensions) };
    const again = await importFile(new File([await exportDocx(doc)], 'roundtrip.docx'));
    return { doc, again: generateJSON(again.html, editorExtensions) };
  }, await layoutFixture());
  const table = result.doc.content.content[0], first = table.content[0].content[0], second = table.content[0].content[1];
  expect(first.attrs.fontSize).toBeNull();
  expect(first.attrs.textColor).toBeNull();
  expect(first.content[0].attrs.paragraphSize).toBe('4pt');
  expect(first.content[1].content[0].marks[0].attrs).toMatchObject({ fontSize: '10pt', color: 'rgb(0, 0, 0)', fontFamily: 'Arial' });
  expect(first.content[1].attrs).toMatchObject({ spaceBefore: '4px', spaceAfter: '6.66667px', indent: 8, rightIndent: '4px', firstLineIndent: '-2px' });
  expect(second.content[0].content[0].marks[0].attrs).toMatchObject({ fontSize: '9pt', color: 'rgb(0, 128, 0)', letterSpacing: '-0.25pt' });
  expect(first.attrs).toMatchObject({ colwidth: [120], paddingTop: '2px', paddingLeft: '6px' });
  expect(second.attrs.colwidth).toEqual([300]);
  expect(table.content[1].content[0].attrs.colwidth).toEqual([120, 300]);
  expect(result.doc.content.content.map((n: { type: string }) => n.type)).toEqual(['table', 'paragraph', 'pageBreak', 'paragraph']);
  const round = result.again.content[0];
  expect(round.content[0].content[0].attrs.colwidth).toEqual([120]);
  expect(round.content[0].content[0].content[1].attrs).toMatchObject({ spaceBefore: '4px', spaceAfter: '6.66667px', rightIndent: '4px', firstLineIndent: '-2px' });
  expect(round.content[1].content[0].content[0].attrs.lineSpacing).toBe('15pt');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'layout.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.doc)) });
  await expect(page.locator('.tiptap table').first()).toHaveCSS('width', '420px');
  await expect(page.locator('.tiptap p').filter({ hasText: 'Readable body' }).locator('span')).toHaveCSS('font-size', '13.3333px');
});

test('nested form rows paginate in editor and print without shrinking or losing text', async ({ page }) => {
  await page.goto('/');
  const paragraphs = Array.from({ length: 55 }, (_, i) => `<p style="font-size:10pt;line-height:1.5;margin-bottom:5px">Form paragraph ${i} — nested content stays readable.</p>`).join('');
  const rows = Array.from({ length: 60 }, (_, i) => `<tr><td><p>Row ${i}</p></td><td><p>Value ${i}</p></td></tr>`).join('');
  const html = `<table><tr><td>${paragraphs}<table>${rows}</table><p>Last form paragraph</p></td></tr></table>`;
  await page.locator('input[type=file]').first().setInputFiles({ name: 'form.html', mimeType: 'text/html', buffer: Buffer.from(html) });
  await expect(page.locator('.tiptap .pagination-gap').first()).toBeAttached();
  const before = await page.locator('.tiptap').innerText();
  const count = await page.locator('.document-stack').getAttribute('data-pages');
  await page.waitForTimeout(500);
  expect(await page.locator('.document-stack').getAttribute('data-pages')).toBe(count);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page').first()).toBeVisible();
  const result = await page.locator('.print-layout').evaluate(host => ({
    pages: host.querySelectorAll('.print-page').length,
    shrunk: [...host.querySelectorAll<HTMLElement>('[style]')].some(el => el.style.zoom && Number(el.style.zoom) < .999),
    paragraphs: [...host.querySelectorAll('p')].map(p => p.textContent).join(''),
    overflow: [...host.querySelectorAll('.print-content')].some(el => el.lastElementChild!.getBoundingClientRect().bottom > el.getBoundingClientRect().bottom + 1),
  }));
  expect(result.pages).toBeGreaterThan(3);
  expect(result.shrunk).toBe(false);
  expect(result.overflow).toBe(false);
  expect(result.paragraphs.replace(/\s/g, '')).toBe(before.replace(/\s/g, ''));
});
