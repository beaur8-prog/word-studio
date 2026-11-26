import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';

type JsonNode = { type?: string; attrs?: Record<string, unknown>; text?: string; content?: JsonNode[] };
const bodyPage = [1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3];
const headers = ['HEAD-A', 'HEAD-B', 'HEAD-C', 'HEAD-D'];
const labels = ['MERGE-C', 'C-LINE-02', 'MERGE-A'];

async function makeFixture(lastRowTooTall: boolean) {
  const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const widths = [1560, 1440, 1560, 1440];
  const paragraph = (text: string, line = 300) => `<w:p><w:pPr><w:keepNext w:val="off"/>
    <w:keepLines w:val="off"/><w:widowControl w:val="off"/>
    <w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="exact"/><w:jc w:val="left"/></w:pPr>${
    text ? `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>${
      text.split('\n').map(value => `<w:t>${value}</w:t>`).join('<w:br/>')}</w:r>` : ''}</w:p>`;
  const cell = (col: number, text: string, merge = '', alignment = 'top') => `<w:tc><w:tcPr>
    <w:tcW w:w="${widths[col]}" w:type="dxa"/>${merge ? `<w:vMerge w:val="${merge}"/>` : ''}
    <w:vAlign w:val="${alignment}"/></w:tcPr>${paragraph(text)}</w:tc>`;
  const row = (cells: string[], height: number, header = false) => `<w:tr><w:trPr><w:cantSplit/>
    <w:trHeight w:val="${height}" w:hRule="atLeast"/>${header ? '<w:tblHeader/>' : ''}</w:trPr>${cells.join('')}</w:tr>`;
  const body = markerIDs('ROW').map((id, index) => row([
    cell(0, index === 1 ? 'MERGE-A' : '', index === 0 ? '' : index === 1 ? 'restart' : 'continue', index === 1 ? 'center' : 'top'),
    cell(1, id),
    cell(2, index === 0 ? 'MERGE-C\nC-LINE-02' : index >= 9 ? 'SINGLE-C' : '',
      index === 0 ? 'restart' : index < 9 ? 'continue' : '', index === 0 ? 'center' : 'top'),
    cell(3, id.replace('ROW', 'CHECK')),
  ], lastRowTooTall && index === 10 ? 8400 : 1080)).join('');
  const edges = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side =>
    `<w:${side} w:val="single" w:sz="4" w:color="000000"/>`).join('');
  const margins = ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="30" w:type="dxa"/>`).join('');
  const table = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>${edges}</w:tblBorders>
    <w:tblLayout w:type="fixed"/><w:tblCellMar>${margins}</w:tblCellMar></w:tblPr>
    <w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
    ${row(headers.map((head, index) => cell(index, head)), 480, true)}${body}</w:tbl>`;
  const lead = [1, 2, 3, 4].map(i => paragraph(`PREFIX-0${i}`, 1060)).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="${word}"><w:body>${lead}${table}${paragraph('TAIL')}
    <w:sectPr><w:pgSz w:w="8640" w:h="8640"/>
    <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>
    </w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

function markerIDs(prefix: string) {
  return bodyPage.map((_, index) => prefix + '-' + String(index + 1).padStart(2, '0'));
}

function snapshot(root: Locator, printing: boolean) {
  return root.evaluate(async (scope, printing) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const wrapper = scope as HTMLElement;
    const editor = wrapper.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    const base = editor ?? wrapper;
    const baseBox = base.getBoundingClientRect();
    const scale = documentScale(base);
    const height = Number(wrapper.dataset.pageHeight);
    const stride = height + 24;
    const markers: { text: string; page: number; y: number; fontSize: number; inside: boolean }[] = [];
    const walker = document.createTreeWalker(wrapper, NodeFilter.SHOW_TEXT);
    const range = document.createRange();
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      const p = node.parentElement?.closest('p');
      const cell = p?.closest('td,th');
      if (!p || p.closest('[data-flow-widget]')) continue;
      for (const match of node.data.matchAll(/HEAD-[A-D]|(?:ROW|CHECK)-\d{2}|MERGE-[AC]|C-LINE-02|TAIL/g)) {
        const index = match.index!;
        range.setStart(node, index); range.setEnd(node, index + match[0].length);
        const ink = range.getBoundingClientRect();
        const sheet = printing ? p.closest<HTMLElement>('.print-page') : undefined;
        if (printing && !sheet) continue;
        const sheetScale = sheet ? documentScale(sheet) : scale;
        const local = (ink.top - baseBox.top) / scale;
        const page = sheet ? Number(sheet.dataset.page) : 1 + Math.floor((local + .5) / stride);
        const pageOffset = sheet ? sheet.getBoundingClientRect().top : baseBox.top + (page - 1) * stride * scale;
        const y = (ink.top - pageOffset) / sheetScale;
        const body = sheet?.querySelector<HTMLElement>('.print-content')?.getBoundingClientRect();
        const contentTop = body ? (body.top - pageOffset) / sheetScale : Number(wrapper.dataset.margin);
        const contentBottom = body ? (body.bottom - pageOffset) / sheetScale : height - Number(wrapper.dataset.bottomMargin);
        const box = cell?.getBoundingClientRect();
        markers.push({
          text: match[0], page, y,
          fontSize: parseFloat(getComputedStyle(node.parentElement!).fontSize),
          inside: ink.width > 0 && ink.height > 0 && y >= contentTop - 1 &&
            (ink.bottom - pageOffset) / sheetScale <= contentBottom + 1 &&
            (!box || (ink.left >= box.left - 1 && ink.right <= box.right + 1 &&
              ink.top >= box.top - 1 && ink.bottom <= box.bottom + 1)),
        });
      }
    }
    const pages = [...wrapper.querySelectorAll<HTMLElement>('.print-page')];
    const pageTables = pages.map(p => p.querySelector<HTMLTableElement>('.print-content > table'));
    const rootTables = printing ? pageTables.flatMap((table, index) =>
      table ? [{ table, baseScale: documentScale(pages[index]) }] : []) :
      [{ table: editor!.querySelector<HTMLTableElement>('table')!, baseScale: scale }];
    const tableWidths = rootTables.map(({ table, baseScale }) => table.getBoundingClientRect().width / baseScale);
    const scaleRatios = rootTables.map(({ table, baseScale }) => documentScale(table) / baseScale);
    return {
      pages: printing ? pages.length : Number(wrapper.dataset.pages),
      markers, tableWidths, scaleRatios,
      bodyRowsPerPrintPage: printing ? pageTables.map(table => table ? [...table.rows].filter(row =>
        [...row.cells].some(cell => cell.tagName === 'TD')).length : 0) : [],
    };
  }, printing);
}

async function openFixture(page: Page, lastRowTooTall = false) {
  const title = lastRowTooTall ? 'oversized-protected-final-row' : 'long-protected-rows';
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: title + '.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await makeFixture(lastRowTooTall),
  });
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(title);
  await expect(page.locator('.busy-indicator')).toBeHidden();
  await page.evaluate(() => document.fonts.ready);
}

test('staggered protected merge crosses two page boundaries at full size', async ({ page }) => {
  await openFixture(page);
  const stack = page.locator('.document-stack');
  const editor = stack.locator('.tiptap[contenteditable=true]');
  const readModel = () => editor.evaluate(element => {
    const model = (element as HTMLElement & { editor: { getJSON(): JsonNode } }).editor.getJSON();
    const table = model.content!.find(node => node.type === 'table')!;
    const text = (node: JsonNode): string => node.text ?? (node.content ?? []).map(text).join('');
    return table.content!.map(row => ({
      protection: row.attrs?.rowSplit,
      rowRule: row.attrs?.wordRowRule,
      cells: row.content!.map(cell => ({
        span: cell.attrs?.rowspan, valign: cell.attrs?.verticalAlign, type: cell.type, text: text(cell),
      })),
    }));
  });
  const sourceTable = await readModel();
  expect(sourceTable).toHaveLength(12);
  expect(sourceTable.every(row => row.protection === 'avoid' && row.rowRule === 'atLeast')).toBe(true);
  expect(sourceTable[0].cells.map(cell => cell.type)).toEqual(Array(4).fill('tableHeader'));
  expect(sourceTable[1].cells[2]).toMatchObject({ span: 9, valign: 'middle', text: 'MERGE-CC-LINE-02' });
  expect(sourceTable[2].cells[0]).toMatchObject({ span: 10, valign: 'middle', text: 'MERGE-A' });

  await expect.poll(async () => (await snapshot(stack, false)).markers
    .filter(m => m.text.startsWith('ROW-')).map(m => m.page)).toEqual(bodyPage);
  await expect(stack.locator('.pagination-headers th')).toHaveText([...headers, ...headers]);
  const edited = await snapshot(stack, false);

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(3);
  const printed = await snapshot(preview, true);

  expect(printed.bodyRowsPerPrintPage).toEqual([2, 5, 4]);
  for (const mode of [edited, printed]) {
    expect(mode.pages).toBe(3);
    for (const prefix of ['ROW', 'CHECK']) {
      const markers = mode.markers.filter(m => m.text.startsWith(prefix + '-'));
      expect(markers.map(m => m.text)).toEqual(markerIDs(prefix));
      expect(markers.map(m => m.page)).toEqual(bodyPage);
    }
    for (const header of headers)
      expect(mode.markers.filter(m => m.text === header).map(m => m.page).sort()).toEqual([1, 2, 3]);
    expect(mode.markers.filter(m => m.text === 'TAIL').map(m => m.page)).toEqual([3]);
    const owners = mode.markers.filter(m => labels.includes(m.text));
    expect(owners.map(m => m.text)).toEqual(labels);
    expect(owners.map(m => m.page)).toEqual([1, 1, 1]);
    expect(owners[1].y - owners[0].y).toBeCloseTo(20, 1);
    expect(mode.markers.filter(m => !m.inside)).toEqual([]);
    for (const marker of mode.markers) expect(marker.fontSize, marker.text).toBeCloseTo(40 / 3, 2);
    for (const ratio of mode.scaleRatios) expect(ratio).toBeCloseTo(1, 4);
    for (const width of mode.tableWidths) expect(width).toBeCloseTo(edited.tableWidths[0], 1);
  }
  for (const mark of printed.markers.filter(m => !m.text.startsWith('HEAD-'))) {
    const match = edited.markers.find(m => m.text === mark.text)!;
    expect(Math.abs(match.y - mark.y), mark.text).toBeLessThan(0.5);
  }

  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  expect(await readModel()).toEqual(sourceTable);
});


test('an oversized final protected row abandons the entire planned continuation', async ({ page }) => {
  await openFixture(page, true);
  const stack = page.locator('.document-stack');

  // The existing fallback is intentionally different between the two modes.
  // Detect even a single committed prefix from the new staged plan.
  const editorFallback = [1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2];
  await expect.poll(async () => (await snapshot(stack, false)).markers
    .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page)).toEqual(editorFallback);
  await expect(stack.locator('.pagination-merge-row')).toHaveCount(0);

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(3);
  const printed = await snapshot(preview, true);
  expect(printed.bodyRowsPerPrintPage).toEqual([0, 11, 0]);
  for (const prefix of ['ROW', 'CHECK']) {
    const markers = printed.markers.filter(marker => marker.text.startsWith(prefix + '-'));
    expect(markers.map(marker => marker.text)).toEqual(markerIDs(prefix));
    expect(markers.map(marker => marker.page)).toEqual(Array(11).fill(2));
  }
  for (const header of headers)
    expect(printed.markers.filter(marker => marker.text === header).map(marker => marker.page)).toEqual([2]);
  expect(printed.markers.filter(marker => labels.includes(marker.text)).map(marker => [marker.text, marker.page]))
    .toEqual(labels.map(label => [label, 2]));
  await expect(preview.locator('[data-print-merge-continuation]')).toHaveCount(0);
});

test('a failed live continuation restores the original pages and table', async ({ page }) => {
  await openFixture(page);
  // Only the physical third paper is obstructed. The offscreen plan still
  // uses the ordinary fresh-body height, so this exercises the live recheck.
  await page.addStyleTag({ content: '.print-layout .print-page[data-page="3"] > .print-content { height: 160px !important; }' });

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(3);
  const pageNumbers = await preview.locator('.print-page').evaluateAll(papers => papers.map(paper => paper.getAttribute('data-page')));
  expect(pageNumbers).toEqual(['1', '2', '3']);
  const printed = await snapshot(preview, true);
  expect(printed.bodyRowsPerPrintPage).toEqual([0, 11, 0]);
  const rows = printed.markers.filter(mark => mark.text.startsWith('ROW-'));
  expect(rows.map(mark => mark.text)).toEqual(markerIDs('ROW'));
  expect(rows.map(mark => mark.page)).toEqual(Array(11).fill(2));
  await expect(preview.locator('[data-print-merge-continuation]')).toHaveCount(0);
});
