import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, ImageRun } from 'docx';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, properties = '') => `<w:tc><w:tcPr>${properties}</w:tcPr>${p(text)}</w:tc>`;
const section = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="360"/>';

async function fixture(body: string, props = section, parts: Record<string, string> = {}) {
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('fixture')] }] })));
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr>${props}</w:sectPr></w:body></w:document>`);
  for (const [file, text] of Object.entries(parts)) zip.file(file, text);
  return Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
}
async function roundtrip(page: Page, bytes: Buffer) {
  return page.evaluate(async data => {
    // @ts-expect-error Vite module.
    const io = await import('/src/document-io.ts');
    // @ts-expect-error Vite module.
    const { editorExtensions } = await import('/src/extensions.ts');
    // @ts-expect-error Vite dependency.
    const { generateJSON } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const first = await io.importFile(new File([new Uint8Array(data)], 'case.docx'));
    const doc = { ...first.document, content: generateJSON(first.html, editorExtensions) };
    io.validateDocument(doc);
    const saved = await io.exportDocx(doc);
    const loaded = await io.importFile(new File([saved], 'again.docx'));
    const again = { ...loaded.document, content: generateJSON(loaded.html, editorExtensions) };
    io.validateDocument(again);
    return { doc, again, html: first.html, warnings: first.warnings, exported: [...new Uint8Array(await saved.arrayBuffer())] };
  }, [...bytes]);
}
test.beforeEach(async ({ page }) => { await page.goto('/'); });

test('resized inline pictures retain physical dimensions, paragraph position and save roundtrip', async ({ page }) => {
  const pixel = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aN1sAAAAASUVORK5CYII=', 'base64');
  const bytes = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph({ children: [
    new ImageRun({ type: 'png', data: pixel, transformation: { width: 376, height: 329 } }),
  ] }), new Paragraph('Caption after image')] }] }));
  const result = await roundtrip(page, bytes);
  for (const doc of [result.doc, result.again]) {
    expect(doc.content.content[0].type).toBe('paragraph');
    expect(doc.content.content[0].content[0]).toMatchObject({ type: 'inlineImage', attrs: { width: 376, height: 329 } });
    expect(doc.content.content).toHaveLength(2);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'picture.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
  await expect(page.locator('.tiptap > p > img[data-inline-image]')).toHaveCount(1);
  await expect(page.locator('.tiptap img[data-inline-image]')).toHaveCSS('width', '376px');
  await expect(page.locator('.tiptap img[data-inline-image]')).toHaveCSS('height', '329px');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-content img')).toHaveCSS('width', '376px');
  await expect(page.locator('.print-content img')).toHaveCSS('height', '329px');
});

test('omitted table grid cells, row offsets and vertical merges open without losing content', async ({ page }) => {
  const body = `<w:tbl><w:tblGrid>${'<w:gridCol w:w="1800"/>'.repeat(4)}</w:tblGrid>
    <w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>${cell('Top', '<w:gridSpan w:val="2"/>')}</w:tr>
    <w:tr>${cell('Merged', '<w:vMerge w:val="restart"/>')}${cell('B')}${cell('C')}${cell('D')}</w:tr>
    <w:tr>${cell('', '<w:vMerge/>')}${cell('E')}${cell('F')}${cell('G')}</w:tr>
    <w:tr>${cell('Short row', '<w:gridSpan w:val="3"/>')}</w:tr></w:tbl>${p('After table')}`;
  const result = await roundtrip(page, await fixture(body));
  for (const doc of [result.doc, result.again]) {
    const table = doc.content.content[0];
    expect(table.content).toHaveLength(4);
    expect(table.content[0].content.map((c: any) => c.attrs.colspan)).toEqual([1, 2, 1]);
    expect(table.content[1].content[0].attrs.rowspan).toBe(2);
    expect(JSON.stringify(doc)).toContain('Short row');
    expect(JSON.stringify(doc)).toContain('After table');
  }
});

test('legacy horizontal table merges retain both cell contents and width', async ({ page }) => {
  const result = await roundtrip(page, await fixture(`<w:tbl><w:tblGrid>${'<w:gridCol w:w="1800"/>'.repeat(3)}</w:tblGrid>
    <w:tr>${cell('Left', '<w:hMerge w:val="restart"/>')}${cell('Right', '<w:hMerge/>')}${cell('Last')}</w:tr></w:tbl>`));
  const cells = result.doc.content.content[0].content[0].content;
  expect(cells).toHaveLength(2);
  expect(cells[0].attrs.colspan).toBe(2);
  expect(cells[0].content.map((x: any) => x.content?.map((t: any) => t.text).join('')).join(' ')).toBe('Left Right');
});

test('first/odd/even headers and page numbers survive import, print and DOCX save', async ({ page }) => {
  const rels = ['header', 'footer'].flatMap(kind => ['default', 'first', 'even'].map(type => `<Relationship Id="${kind}-${type}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${kind}-${type}.xml"/>`)).join('');
  const refs = ['header', 'footer'].flatMap(kind => ['default', 'first', 'even'].map(type => `<w:${kind}Reference w:type="${type}" r:id="${kind}-${type}"/>`)).join('');
  const parts: Record<string, string> = {
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
    'word/settings.xml': `<w:settings xmlns:w="${W}"><w:evenAndOddHeaders/></w:settings>`,
  };
  for (const kind of ['header', 'footer']) for (const type of ['default', 'first', 'even']) parts[`word/${kind}-${type}.xml`] = `<w:${kind === 'header' ? 'hdr' : 'ftr'} xmlns:w="${W}">${p(`${type} ${kind}`)}${kind === 'footer' && type !== 'first' ? '<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>' : ''}</w:${kind === 'header' ? 'hdr' : 'ftr'}>`;
  const bytes = await fixture(`${p('Cover')}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${p('Even body')}<w:p><w:r><w:br w:type="page"/></w:r></w:p>${p('Odd body')}`, refs + '<w:titlePg/>' + section, parts);
  const result = await roundtrip(page, bytes);
  for (const doc of [result.doc, result.again]) expect(doc.page).toMatchObject({ differentFirstPage: true, differentOddEven: true,
    firstHeader: 'first header', evenHeader: 'even header', header: 'default header', firstFooter: 'first footer', evenPageNumbers: true, pageNumbers: true });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'headers.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.again)) });
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-header')).toHaveText(['first header', 'even header', 'default header']);
  await expect(page.locator('.print-footer')).toHaveText(['first footer', 'even footer2', 'default footer3']);
  await expect(page.locator('.print-header').first()).toHaveCSS('top', '48px');
});

test('binding gutter adds to the content margin once, including after export', async ({ page }) => {
  const result = await roundtrip(page, await fixture(p('Bound report')));
  for (const doc of [result.doc, result.again]) expect(doc.page.gutter).toBeCloseTo(6.35, 3);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'bound.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.again)) });
  await expect(page.locator('.tiptap')).toHaveCSS('padding-left', '120px');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-content')).toHaveCSS('left', '120px');
});

test('long tables repeat enabled header rows and preserve all body rows when printing', async ({ page }) => {
  const grid = '<w:tblGrid><w:gridCol w:w="3500"/><w:gridCol w:w="3500"/></w:tblGrid>';
  const body = `<w:tbl>${grid}<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('Item')}${cell('Amount')}</w:tr>${Array.from({ length: 80 }, (_, i) => `<w:tr>${cell(`Row ${i}`)}${cell(String(i * 100))}</w:tr>`).join('')}</w:tbl>`;
  await page.locator('input[type=file]').first().setInputFiles({ name: 'report.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await fixture(body) });
  await expect(page.getByRole('textbox', { name: '문서 편집 영역' }).locator('th')).toHaveCount(2);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect.poll(() => page.locator('.print-page').count()).toBeGreaterThan(1);
  const sheets = await page.locator('.print-page').count();
  await expect(page.locator('.print-content th')).toHaveCount(sheets * 2);
  await expect(page.locator('.print-content td').filter({ hasText: /^Row \d+$/ })).toHaveCount(80);
});

test('right and center tab stops with leaders align after typing and survive DOCX save', async ({ page }) => {
  const tabbed = (type: string, text: string) => `<w:p><w:pPr><w:tabs><w:tab w:val="${type}" w:pos="7200" w:leader="dot"/></w:tabs></w:pPr><w:r><w:t>Item</w:t><w:tab/><w:t>${text}</w:t></w:r></w:p>`;
  const bytes = await fixture(tabbed('right', '123') + tabbed('center', 'Center'));
  const result = await roundtrip(page, bytes);
  for (const doc of [result.doc, result.again]) {
    expect(doc.content.content[0].content.map((n: any) => n.type)).toEqual(['text', 'documentTab', 'text']);
    expect(JSON.parse(doc.content.content[0].attrs.tabStops)).toEqual([{ position: 480, type: 'right', leader: 'dot' }]);
  }
  await page.locator('input[type=file]').first().setInputFiles({ name: 'menu.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
  await expect(page.locator('.tiptap [data-docx-tab]')).toHaveCount(2);
  const aligned = async (selector: string) => page.locator(selector).evaluateAll(ps => ps.map((p, index) => {
    const range = document.createRange(); range.selectNodeContents(p); range.setStartAfter(p.querySelector('[data-docx-tab]')!);
    const box = range.getBoundingClientRect(); return (index ? (box.left + box.right) / 2 : box.right) - p.getBoundingClientRect().left;
  }));
  await expect.poll(async () => (await aligned('.tiptap > p')).map(x => Math.round(x))).toEqual([480, 480]);
  await expect(page.locator('[data-docx-tab]').first()).toHaveAttribute('data-leader', 'dot');
  await page.locator('.tiptap').click(); await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowUp' : 'Control+Home'); await page.keyboard.type('Longer item: ');
  await expect.poll(async () => (await aligned('.tiptap > p')).map(x => Math.round(x))).toEqual([480, 480]);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect.poll(async () => (await aligned('.print-content > p')).map(x => Math.round(x))).toEqual([480, 480]);
});

test('Tab inserts a document tab outside tables without losing keyboard focus', async ({ page }) => {
  await page.locator('input[type=file]').first().setInputFiles({ name: 'typing.txt', mimeType: 'text/plain', buffer: Buffer.from('Label') });
  await page.locator('.tiptap').click(); await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowDown' : 'Control+End'); await page.keyboard.press('Tab'); await page.keyboard.type('Value');
  await expect(page.locator('.tiptap')).toBeFocused();
  await expect(page.locator('.tiptap [data-docx-tab]')).toHaveCount(1);
  await expect(page.locator('.tiptap')).toHaveText('Label\tValue');
});

test('negative tab stops and consecutive tabs retain the stop at the text margin', async ({ page }) => {
  const body = '<w:p><w:pPr><w:ind w:hanging="1080"/><w:tabs><w:tab w:val="left" w:pos="-720"/><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:tab/><w:t>-0.5</w:t><w:tab/><w:tab/><w:t>0.5</w:t></w:r></w:p>';
  await page.locator('input[type=file]').first().setInputFiles({ name: 'negative.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await fixture(body) });
  await expect(page.locator('.tiptap [data-docx-tab]')).toHaveCount(3);
  await expect.poll(async () => page.locator('.tiptap > p').evaluate(p => {
    const range = document.createRange(); range.selectNodeContents(p); range.setStartAfter(p.querySelectorAll('[data-docx-tab]')[2]);
    return Math.round(range.getBoundingClientRect().left - p.getBoundingClientRect().left);
  })).toBe(48);
});

test('theme fonts and mixed Korean/Latin runs keep their intended typefaces after saving', async ({ page }) => {
  const bytes = await fixture('<w:p><w:r><w:t>English 한글</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t> Direct</w:t></w:r></w:p>', section, {
    'word/styles.xml': `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsia="Malgun Gothic"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/theme/theme1.xml': '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="Test"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
  });
  const result = await roundtrip(page, bytes);
  for (const doc of [result.doc, result.again]) {
    const runs = doc.content.content[0].content;
    expect(runs.map((n: any) => [n.text, n.marks.find((m: any) => m.type === 'textStyle')?.attrs.fontFamily])).toEqual([
      ['English ', 'Calibri'], ['한글', '"Malgun Gothic"'], [' Direct', 'Arial'],
    ]);
  }
});

test('header dialog creates and edits independent first and even page content', async ({ page }) => {
  await page.locator('input[type=file]').first().setInputFiles({ name: 'book.txt', mimeType: 'text/plain', buffer: Buffer.from('Book body') });
  await page.getByRole('button', { name: '삽입', exact: true }).click();
  await page.getByRole('button', { name: '머리글·바닥글', exact: true }).click();
  await page.getByRole('textbox', { name: '머리글', exact: true }).fill('Chapter');
  await page.getByLabel('첫 페이지 다르게 지정', { exact: true }).check();
  await page.getByRole('textbox', { name: '첫 페이지 머리글', exact: true }).fill('Cover');
  await page.getByLabel('홀수·짝수 페이지 다르게 지정', { exact: true }).check();
  await page.getByRole('textbox', { name: '짝수 페이지 머리글', exact: true }).fill('Even');
  await page.getByRole('button', { name: '완료', exact: true }).click();
  await expect(page.locator('.paper-header').first()).toHaveText('Cover');
  // Rendering a header is immediate; recovery is deliberately debounced.
  // Reload only after persistence completes, as in the lifecycle tests.
  await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
  await page.reload();
  await expect(page.locator('.paper-header').first()).toHaveText('Cover');
});
