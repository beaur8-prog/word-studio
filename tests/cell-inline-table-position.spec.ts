import type { Editor } from '@tiptap/core';
import { expect, test, type Locator, type Page } from '@playwright/test';
import { AlignmentType, Document, Packer, Paragraph, Table, TableAnchorType, TableCell, TableLayoutType, TableRow, TextRun, WidthType } from 'docx';
import JSZip from 'jszip';

const paragraph = (text: string) => new Paragraph({ children: [new TextRun(text)] });
const columns = [675, 675];
const grid = (text: string, x: number) => new Table({
  layout: TableLayoutType.FIXED, columnWidths: columns, width: { size: 1350, type: WidthType.DXA },
  alignment: AlignmentType.LEFT,
  float: { horizontalAnchor: TableAnchorType.MARGIN, verticalAnchor: TableAnchorType.MARGIN,
    absoluteHorizontalPosition: x, absoluteVerticalPosition: 300 },
  rows: [new TableRow({ children: [new TableCell({ children: [paragraph(text)] }), new TableCell({ children: [paragraph('blank grid marker')] })] })],
});

function fixture(rightPosition = 1950) {
  const host = new Table({
    layout: TableLayoutType.FIXED, columnWidths: [4500], width: { size: 4500, type: WidthType.DXA },
    rows: [new TableRow({ children: [new TableCell({ children: [paragraph('    '),
      grid('LEFT LABEL', 300), new Paragraph({}), grid('RIGHT LABEL', rightPosition), new Paragraph({}),
    ] })] })],
  });
  return Packer.toBuffer(new Document({ sections: [{ children: [host, paragraph('Tail remains editable')] }] }));
}

function pairedLayout(editor: Locator) {
  return editor.locator('table[data-cell-table-position]').evaluateAll(tables => tables.map(table => {
    const owner = table.parentElement!.closest('td')!, rect = table.getBoundingClientRect(), host = owner.getBoundingClientRect();
    return { paired: table.parentElement!.classList.contains('cell-table-frame'), x: Math.round(rect.x - host.x),
      y: Math.round(rect.y - host.y), width: Math.round(rect.width) };
  }));
}

async function saveDocx(page: Page) {
  await page.getByRole('button', { name: '파일', exact: true }).click();
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
  const stream = await (await download).createReadStream(), bytes: Buffer[] = [];
  for await (const block of stream) bytes.push(Buffer.from(block));
  return Buffer.concat(bytes);
}

async function savedXml(page: Page) {
  const zip = await JSZip.loadAsync(await saveDocx(page));
  return zip.file('word/document.xml')!.async('string');
}

async function logicalMarginFixture() {
  const zip = await JSZip.loadAsync(await fixture(3075));
  const part = 'word/document.xml';
  const body = await zip.file(part)!.async('string');
  const logical = '<w:tblCellMar><w:start w:w="30" w:type="dxa"/><w:end w:w="30" w:type="dxa"/></w:tblCellMar>';
  const ordinary = `<w:tbl><w:tblPr><w:tblStyle w:val="WriterTable"/>${logical}</w:tblPr>
    <w:tblGrid><w:gridCol w:w="1350"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>UNRELATED TABLE</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
  const firstTable = body.replace('<w:tblPr>', '<w:tblPr><w:tblStyle w:val="WriterTable"/>')
    .replace('</w:tblPr>', logical + '</w:tblPr>')
    .replace('<w:sectPr', ordinary + '<w:sectPr');
  zip.file(part, firstTable);
  const styles = await zip.file('word/styles.xml')!.async('string');
  zip.file('word/styles.xml', styles.replace('</w:styles>', `<w:style w:type="table" w:styleId="WriterTable">
    <w:name w:val="Writer Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/>
    <w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`));
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('explicit logical table margins preserve a fitting pair across two DOCX saves', async ({ page }) => {
  await page.goto('/');
  let payload = await logicalMarginFixture();
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  for (let stage = 0; stage < 3; stage++) {
    await page.locator('input[type=file]').first().setInputFiles({
      name: `side-by-side-with-logical-margins-${stage}.docx`,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: payload,
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(`side-by-side-with-logical-margins-${stage}`);
    const host = editor.locator('td').first();
    await expect(host).toHaveCSS('padding-left', '2px');
    await expect(host).toHaveCSS('padding-right', '2px');
    const other = editor.locator('td').filter({ hasText: 'UNRELATED TABLE' });
    await expect(other).toHaveCSS('padding-right', '7.2px');
    await expect.poll(async () => (await pairedLayout(editor)).map(({ paired, x, y }) => ({ paired, x, y }))).toEqual([
      { paired: true, x: 20, y: 20 },
      { paired: true, x: 205, y: 20 },
    ]);
    await expect(editor.getByText('LEFT LABEL', { exact: true })).toBeVisible();
    await expect(editor.getByText('RIGHT LABEL', { exact: true })).toBeVisible();
    if (stage < 2) payload = await saveDocx(page);
  }
});

test('nested paired tables remain ordinary blocks after their cell is narrowed or gains text', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'side-by-side nested tables.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await fixture(),
  });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const twoAtSource = () => pairedLayout(editor).then(layout => layout.map((item, index) =>
    item.paired && Math.abs(item.x - (index === 0 ? 20 : 130)) <= 1 && item.y === 20));
  await expect.poll(twoAtSource).toEqual([true, true]);
  const first = editor.locator('table[data-cell-table-position]').first();
  const before = await savedXml(page);
  expect(before.match(/<w:tblpPr\b/g)).toHaveLength(2);
  expect(before.match(/<w:tbl\b/g)).toHaveLength(3);

  await first.evaluate(table => {
    const host = table.parentElement!.closest('td')!;
    const view = (table.closest('.tiptap') as HTMLElement & { editor: Editor }).editor.view;
    const position = view.posAtDOM(host, 0) - 1;
    const node = view.state.doc.nodeAt(position);
    view.dispatch(view.state.tr.setNodeMarkup(position, null, { ...node.attrs, colwidth: [180] }));
  });
  await expect.poll(async () => (await pairedLayout(editor)).map(item => item.paired)).toEqual([false, false]);
  const narrow = await savedXml(page);
  expect(narrow.match(/<w:tbl\b/g)).toHaveLength(3);
  expect(narrow).not.toMatch(/<w:tblpPr\b/);
  expect(narrow).toContain('LEFT LABEL');
  expect(narrow).toContain('RIGHT LABEL');

  await editor.focus();
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(twoAtSource).toEqual([true, true]);
  await first.evaluate(table => {
    const host = table.parentElement!.closest('td')!;
    const view = (table.closest('.tiptap') as HTMLElement & { editor: Editor }).editor.view;
    const p = host.querySelector(':scope > p')!;
    view.dispatch(view.state.tr.insertText('A real cell paragraph', view.posAtDOM(p, 0)));
  });
  await expect(editor.getByText('A real cell paragraph', { exact: false })).toBeVisible();
  await expect.poll(async () => (await pairedLayout(editor)).map(item => item.paired)).toEqual([false, false]);
  const text = await savedXml(page);
  expect(text.match(/<w:tbl\b/g)).toHaveLength(3);
  expect(text).toContain('A real cell paragraph');
  expect(text).not.toMatch(/<w:tblpPr\b/);

  await editor.focus();
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(twoAtSource).toEqual([true, true]);
  await first.locator('td > p').first().evaluate(p => {
    const view = (p.closest('.tiptap') as HTMLElement & { editor: Editor }).editor.view;
    const at = view.posAtDOM(p, p.childNodes.length);
    view.dispatch(view.state.tr.insert(at, view.state.schema.nodes.hardBreak.create()));
  });
  await expect.poll(async () => (await pairedLayout(editor)).map(item => item.paired)).toEqual([false, false]);
  const control = await savedXml(page);
  expect(control).toMatch(/<w:br\b/);
  expect(control).not.toMatch(/<w:tblpPr\b/);
});
