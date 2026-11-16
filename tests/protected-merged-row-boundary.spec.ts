import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';

const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wordRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const widths = [1560, 1440, 1560, 1440];
const heads = ['HEAD-A', 'HEAD-B', 'HEAD-C', 'HEAD-D'];
const rows = ['01', '02', '03', '04', '05'];

type FooterFixture = 'none' | 'different-first-page' | 'empty';

async function fixture(alignment: 'top' | 'bottom' | 'center', footerFixture: FooterFixture, cLineCount = 1, edge: HeaderEdgeFixture = 'none') {
  const differentFirstPage = footerFixture === 'different-first-page';
  const hasFooter = footerFixture !== 'none';
  const paragraph = (text: string, line = 300) => `<w:p><w:pPr>
    <w:keepNext w:val="off"/><w:keepLines w:val="off"/><w:widowControl w:val="off"/>
    <w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="exact"/><w:jc w:val="left"/>
    </w:pPr>${text ? `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>${text.split('\n').map(line => `<w:t>${line}</w:t>`).join('<w:br/>')}</w:r>` : ''}</w:p>`;
  const thickHeader = edge === 'thick' || edge === 'thick-asym' || edge === 'thick-mixed';
  const headerBottom = !thickHeader ? '' :
    '<w:tcBorders><w:bottom w:val="single" w:sz="32" w:color="000000"/></w:tcBorders>';
  const ownerMargins = edge === 'thick-asym' ?
    '<w:tcMar><w:top w:w="90" w:type="dxa"/><w:bottom w:w="15" w:type="dxa"/></w:tcMar>' : '';
  const cell = (column: number, text: string, merge = '', align = 'top', extraProps = '') => `<w:tc><w:tcPr>
    <w:tcW w:w="${widths[column]}" w:type="dxa"/>${merge ? `<w:vMerge w:val="${merge}"/>` : ''}${extraProps}<w:vAlign w:val="${align}"/>
    </w:tcPr>${paragraph(text)}</w:tc>`;
  const row = (cells: string[], header = false) => `<w:tr><w:trPr><w:cantSplit/>
    <w:trHeight w:val="${header ? 480 : 1080}" w:hRule="atLeast"/>${header ? '<w:tblHeader/>' : ''}
    </w:trPr>${cells.join('')}</w:tr>`;
  const bodyRows = rows.map((id, index) => {
    const merge = index === 0 ? 'restart' : 'continue';
    return row([
      cell(0, index === 0 ? 'MERGE-A' : '', merge, index === 0 && edge !== 'thick-mixed' ? alignment : 'top', index === 0 ? ownerMargins : ''),
      cell(1, `ROW-${id}`, '', index === 2 && edge === 'bottom-on-continuation' ? 'bottom' : 'top'),
      cell(2, index === 0 ? centeredCLines(cLineCount).join('\n') : index === 4 ? 'SINGLE-C' : '',
        index === 4 ? '' : merge, index === 0 && alignment === 'center' ? 'center' : 'top', index === 0 ? ownerMargins : ''),
      cell(3, `CHECK-${id}`),
    ]);
  }).join('');
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side =>
    `<w:${side} w:val="single" w:sz="4" w:color="000000"/>`).join('');
  const margins = ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="30" w:type="dxa"/>`).join('');
  const table = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders>
    <w:tblLayout w:type="fixed"/><w:tblCellMar>${margins}</w:tblCellMar></w:tblPr>
    <w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
    ${row(heads.map((head, index) => cell(index, head, '', 'top', headerBottom)), true)}${bodyRows}</w:tbl>`;
  const lead = [1, 2, 3, 4].map(i => paragraph(`PREFIX-0${i}`, i === 4 && thickHeader ? 940 : 1060)).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    (hasFooter ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : '') +
    '</Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  if (hasFooter) {
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${wordRelationships}/footer" Target="footer1.xml"/></Relationships>`);
    const lines = differentFirstPage ? Array.from({ length: 11 }, (_, i) => paragraph(`FOOTER-${i}`, 480)).join('') : paragraph('', 1);
    zip.file('word/footer1.xml', `<w:ftr xmlns:w="${word}">${lines}</w:ftr>`);
  }
  const footer = hasFooter ? '<w:footerReference w:type="default" r:id="rId1"/>' : '';
  zip.file('word/document.xml', `<w:document xmlns:w="${word}" xmlns:r="${wordRelationships}"><w:body>${lead}${table}${paragraph('TAIL')}
    <w:sectPr>${footer}<w:pgSz w:w="8640" w:h="8640"/>
    <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>
    ${differentFirstPage ? '<w:titlePg/>' : ''}</w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

function snapshot(root: Locator, print: boolean) {
  return root.evaluate(async (scope, print) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const paper = scope as HTMLElement;
    const editor = paper.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    const top = editor?.getBoundingClientRect().top ?? paper.getBoundingClientRect().top;
    const scale = documentScale(editor ?? paper);
    const height = Number(paper.dataset.pageHeight);
    const stride = height + 24;
    const markers = [...paper.querySelectorAll<HTMLParagraphElement>('p')].filter(p =>
      /^(?:HEAD-[A-D]|(?:ROW|CHECK)-0[1-5]|MERGE-[AC]|SINGLE-C|TAIL)$/.test(p.textContent ?? ''));
    return {
      pages: print ? paper.querySelectorAll('.print-page').length : Number(paper.dataset.pages),
      markers: markers.map(p => {
        const range = document.createRange();
        range.selectNodeContents(p);
        const ink = range.getBoundingClientRect();
        const sheet = p.closest<HTMLElement>('.print-page');
        const local = (ink.top - top) / scale;
        const page = sheet ? Number(sheet.dataset.page) : 1 + Math.floor((local + .5) / stride);
        const y = sheet ? (ink.top - sheet.getBoundingClientRect().top) / documentScale(sheet) : local - (page - 1) * stride;
        const cell = p.closest('td,th')?.getBoundingClientRect();
        return {
          text: p.textContent!, page, y,
          inside: ink.width > 0 && ink.height > 0 && (!cell || (ink.top >= cell.top - 1 &&
            ink.bottom <= cell.bottom + 1 && ink.left >= cell.left - 1 && ink.right <= cell.right + 1)),
        };
      }),
    };
  }, print);
}

async function openFixture(page: Page, alignment: 'top' | 'bottom' | 'center', footer: FooterFixture = 'none', cLineCount = 1, edge: HeaderEdgeFixture = 'none') {
  await page.goto('/');
  const name = `protected-crossed-${alignment}`;
  await page.locator('input[type=file]').first().setInputFiles({
    name: `${name}.docx`, buffer: await fixture(alignment, footer, cLineCount, edge),
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
  await expect(page.locator('.busy-indicator')).toBeHidden();
  await page.evaluate(() => document.fonts.ready);
}

test('protected rows can continue between complete rows across two short top-aligned merges', async ({ page }) => {
  for (const alignment of ['top', 'bottom'] as const) await test.step(alignment, async () => {
    await openFixture(page, alignment);

    // The first page can fit one header plus two complete body rows.
    // The non-top control still keeps the protected group together.
    const expectedRowPages = alignment === 'top' ? [1, 1, 2, 2, 2] : [2, 2, 2, 2, 2];
    const expectedHeaderPages = alignment === 'top' ? [1, 2] : [2];
    const stack = page.locator('.document-stack');
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor.locator('th')).toHaveText(heads);
    await expect(editor.locator('td[rowspan="5"]')).toHaveText('MERGE-A');
    await expect(editor.locator('td[rowspan="4"]')).toHaveText('MERGE-C');
    const overlayHeads = stack.locator('.pagination-headers th');
    if (alignment === 'top') await expect(overlayHeads).toHaveText(heads);
    else await expect(overlayHeads).toHaveCount(0);
    await expect.poll(async () => (await snapshot(stack, false)).markers
      .filter(m => m.text.startsWith('ROW-')).map(m => m.page)).toEqual(expectedRowPages);
    const edited = await snapshot(stack, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    const printed = await snapshot(preview, true);
    for (const mode of [edited, printed]) {
      expect(mode.pages).toBe(2);
      expect(mode.markers.filter(m => !m.inside)).toEqual([]);
      for (const prefix of ['ROW', 'CHECK']) {
        expect(mode.markers.filter(m => m.text.startsWith(`${prefix}-`)).map(m => [m.text, m.page]))
          .toEqual(rows.map((id, index) => [`${prefix}-${id}`, expectedRowPages[index]]));
      }
      for (const id of rows) {
        const row = mode.markers.find(m => m.text === `ROW-${id}`)!;
        const check = mode.markers.find(m => m.text === `CHECK-${id}`)!;
        expect(Math.abs(row.y - check.y), id).toBeLessThan(1.25);
      }
      for (const head of heads)
        expect(mode.markers.filter(m => m.text === head).map(m => m.page).sort()).toEqual(expectedHeaderPages);
      expect(mode.markers.filter(m => m.text === 'MERGE-A').map(m => m.page)).toEqual([expectedRowPages[0]]);
      expect(mode.markers.filter(m => m.text === 'MERGE-C').map(m => m.page)).toEqual([expectedRowPages[0]]);
      expect(mode.markers.filter(m => m.text === 'SINGLE-C').map(m => m.page)).toEqual([2]);
      expect(mode.markers.filter(m => m.text === 'TAIL').map(m => m.page)).toEqual([2]);
      if (alignment === 'bottom')
        expect(mode.markers.find(m => m.text === 'MERGE-A')!.y)
          .toBeGreaterThan(mode.markers.find(m => m.text === 'ROW-05')!.y);
    }
    for (const marker of printed.markers.filter(m => !m.text.startsWith('HEAD-'))) {
      const corresponding = edited.markers.find(m => m.text === marker.text)!;
      expect(Math.abs(corresponding.y - marker.y), marker.text).toBeLessThan(1.25);
    }
  });
});

test('a larger regular footer leaves a protected merged block together', async ({ page }) => {
  await openFixture(page, 'top', 'different-first-page');
  const stack = page.locator('.document-stack');
  const rowPages = async (root: Locator, print: boolean) => (await snapshot(root, print)).markers
    .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page);

  await expect.poll(() => rowPages(stack, false)).toEqual([2, 2, 2, 2, 2]);
  await expect(stack.locator('.pagination-headers th')).toHaveCount(0);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(3);
  await expect(preview.locator('.print-footer').nth(1)).toContainText('FOOTER-10');
  await expect(preview.locator('.print-footer').first()).not.toContainText('FOOTER');
  expect(await rowPages(preview, true)).toEqual([2, 2, 2, 2, 2]);
  await expect(preview.locator('th')).toHaveText(heads);
});

test('an empty DOCX footer keeps the whole-row continuation available', async ({ page }) => {
  await openFixture(page, 'top', 'empty');
  const stack = page.locator('.document-stack');
  await expect(stack.locator('.pagination-headers th')).toHaveText(heads);
  await expect.poll(async () => (await snapshot(stack, false)).markers
    .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page)).toEqual([1, 1, 2, 2, 2]);

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(2);
  await expect(preview.locator('.print-page').first().locator('th')).toHaveText(heads);
  await expect(preview.locator('.print-page').last().locator('th')).toHaveText(heads);
});


function centeredCLines(count: number) {
  return ['MERGE-C', ...Array.from({ length: count - 1 }, (_, i) => `C-LINE-0${i + 2}`)];
}

function ownerLinePositions(root: Locator, print: boolean) {
  return root.evaluate(async (scope, print) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const paper = scope as HTMLElement;
    const editor = paper.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    const top = (editor ?? paper).getBoundingClientRect().top;
    const scale = documentScale(editor ?? paper);
    const stride = Number(paper.dataset.pageHeight) + 24;
    const lines: { text: string; page: number; y: number; inside: boolean }[] = [];
    const walker = document.createTreeWalker(paper, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      const paragraph = node.parentElement?.closest('p');
      const cell = paragraph?.closest('td');
      if (!cell || paragraph!.closest('.pagination-headers,[data-flow-widget]')) continue;
      for (const match of node.data.matchAll(/MERGE-[AC]|C-LINE-0[2-9]/g)) {
        const rng = document.createRange();
        rng.setStart(node, match.index!);
        rng.setEnd(node, match.index! + match[0].length);
        const ink = rng.getBoundingClientRect();
        const rect = cell.getBoundingClientRect();
        const sheet = print ? paragraph!.closest<HTMLElement>('.print-page') : null;
        const local = (ink.top - top) / scale;
        const page = sheet ? Number(sheet.dataset.page) : 1 + Math.floor((local + .5) / stride);
        const y = sheet ? (ink.top - sheet.getBoundingClientRect().top) / documentScale(sheet)
          : local - (page - 1) * stride;
        lines.push({
          text: match[0], page, y,
          inside: ink.width > 0 && ink.height > 0 && ink.left >= rect.left - 1 && ink.right <= rect.right + 1 &&
            ink.top >= rect.top - 1 && ink.bottom <= rect.bottom + 1,
        });
      }
    }
    return lines;
  }, print);
}

test('short centered merge labels stay whole in the first protected row fragment', async ({ page }) => {
  for (const lineCount of [2, 9]) await test.step(`${lineCount} exact lines in C`, async () => {
    await openFixture(page, 'center', 'none', lineCount);
    const short = lineCount === 2;
    const rowPages = short ? [1, 1, 2, 2, 2] : [2, 2, 2, 2, 2];
    const headerPages = short ? [1, 2] : [2];
    const stack = page.locator('.document-stack');
    const readRows = async (root: Locator, print: boolean) => (await snapshot(root, print)).markers
      .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page);
    await expect.poll(() => readRows(stack, false)).toEqual(rowPages);
    if (short) await expect(stack.locator('.pagination-headers th')).toHaveText(heads);
    else await expect(stack.locator('.pagination-headers th')).toHaveCount(0);
    const edited = { table: await snapshot(stack, false), lines: await ownerLinePositions(stack, false) };

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    const printed = { table: await snapshot(preview, true), lines: await ownerLinePositions(preview, true) };

    const expectedLabels = ['MERGE-A', ...centeredCLines(lineCount)];
    for (const mode of [edited, printed]) {
      expect(mode.table.pages).toBe(2);
      expect(mode.lines.map(line => line.text)).toEqual(expectedLabels);
      expect(mode.lines.filter(line => !line.inside)).toEqual([]);
      expect(mode.lines.map(line => line.page)).toEqual(expectedLabels.map(() => short ? 1 : 2));
      const [a, ...c] = mode.lines;
      for (let i = 1; i < c.length; i++)
        expect(c[i].y - c[i - 1].y, c[i].text).toBeCloseTo(20, 1);
      for (const prefix of ['ROW', 'CHECK']) {
        expect(mode.table.markers.filter(marker => marker.text.startsWith(`${prefix}-`)).map(marker => marker.page))
          .toEqual(rowPages);
      }
      for (const head of heads)
        expect(mode.table.markers.filter(marker => marker.text === head).map(marker => marker.page).sort())
          .toEqual(headerPages);
      if (short) {
        const row1 = mode.table.markers.find(marker => marker.text === 'ROW-01')!;
        const row2 = mode.table.markers.find(marker => marker.text === 'ROW-02')!;
        expect(a.y).toBeGreaterThan((row1.y + row2.y) / 2);
        expect(a.y).toBeLessThan(row2.y);
        // The two owners have different source spans. Both have consumed the
        // same two physical rows, so the centers of their ink blocks align.
        expect(Math.abs(a.y - (c[0].y + c[1].y) / 2)).toBeLessThan(1.25);
      }
    }
    for (const [index, line] of printed.lines.entries())
      expect(Math.abs(line.y - edited.lines[index].y), line.text).toBeLessThan(1.25);
  });
});


type HeaderEdgeFixture = 'none' | 'thick' | 'thick-asym' | 'thick-mixed' | 'bottom-on-continuation';

test('protected centered merges retain the repeated thick header edge on continuation', async ({ page }) => {
  const ownerLabels = ['MERGE-A', 'MERGE-C', 'C-LINE-02'];
  const rowPages = [1, 1, 2, 2, 2];
  const centeredY: number[][] = [];
  for (const edge of ['thick', 'thick-asym'] as const) await test.step(edge, async () => {
    await openFixture(page, 'center', 'none', 2, edge);
    const stack = page.locator('.document-stack');
    const collectRows = (markers: Awaited<ReturnType<typeof snapshot>>['markers'], prefix: string) =>
      markers.filter(marker => marker.text.startsWith(`${prefix}-`));
    await expect.poll(async () => collectRows((await snapshot(stack, false)).markers, 'ROW')
      .map(marker => marker.page)).toEqual(rowPages);
    await expect(stack.locator('.pagination-headers th')).toHaveText(heads);
    const edited = { table: await snapshot(stack, false), owners: await ownerLinePositions(stack, false) };

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    const printed = { table: await snapshot(preview, true), owners: await ownerLinePositions(preview, true) };

    for (const mode of [edited, printed]) {
      expect(mode.table.pages).toBe(2);
      for (const head of heads)
        expect(mode.table.markers.filter(marker => marker.text === head).map(marker => marker.page).sort())
          .toEqual([1, 2]);
      for (const prefix of ['ROW', 'CHECK']) {
        const positions = collectRows(mode.table.markers, prefix);
        expect(positions.map(marker => marker.text)).toEqual(rows.map(id => `${prefix}-${id}`));
        expect(positions.map(marker => marker.page)).toEqual(rowPages);
        expect(positions.every(marker => marker.inside)).toBe(true);
        // Matching only paper assignments misses a shifted continuation row.
        // Compare neighboring rows to cancel any platform-dependent glyph offset.
        expect(positions[3].y - positions[2].y).toBeCloseTo(74 + 2 / 3, 1);
        expect(positions[4].y - positions[3].y).toBeCloseTo(76 + 2 / 3, 1);
      }
      expect(mode.owners.map(owner => owner.text)).toEqual(ownerLabels);
      expect(mode.owners.map(owner => owner.page)).toEqual([1, 1, 1]);
      expect(mode.owners.every(owner => owner.inside)).toBe(true);
      const [a, c1, c2] = mode.owners;
      expect(c2.y - c1.y).toBeCloseTo(20, 1);
      expect(Math.abs(a.y - (c1.y + c2.y) / 2)).toBeLessThan(0.5);
    }

    for (const prefix of ['ROW', 'CHECK']) {
      const editorRows = collectRows(edited.table.markers, prefix);
      const printRows = collectRows(printed.table.markers, prefix);
      for (const [index, row] of editorRows.entries())
        expect(Math.abs(row.y - printRows[index].y), `${edge} ${row.text} on paper ${row.page}`)
          .toBeLessThan(0.5);
    }
    for (const [index, owner] of edited.owners.entries())
      expect(Math.abs(owner.y - printed.owners[index].y), `${edge} ${owner.text}`).toBeLessThan(0.5);
    centeredY.push(edited.owners.map(owner => owner.y));
  });
  for (const [index, y] of centeredY[0].entries())
    expect(centeredY[1][index] - y, ownerLabels[index]).toBeCloseTo(4, 1);
});

test('a bottom-aligned continuation cell keeps the centered merge together in both views', async ({ page }) => {
  await openFixture(page, 'center', 'none', 2, 'bottom-on-continuation');
  const stack = page.locator('.document-stack');
  const rowPages = async (root: Locator, print: boolean) => (await snapshot(root, print)).markers
    .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page);
  await expect.poll(() => rowPages(stack, false)).toEqual([2, 2, 2, 2, 2]);

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(2);
  expect(await rowPages(preview, true)).toEqual([2, 2, 2, 2, 2]);
  await expect(preview.locator('th')).toHaveText(heads);
  expect((await ownerLinePositions(preview, true)).map(line => [line.text, line.page]))
    .toEqual([['MERGE-A', 2], ['MERGE-C', 2], ['C-LINE-02', 2]]);
});

test('top and centered owners use the same collapsed edge when they cross together', async ({ page }) => {
  await openFixture(page, 'center', 'none', 2, 'thick-mixed');
  const stack = page.locator('.document-stack');
  await expect(stack.locator('.pagination-headers th')).toHaveText(heads);
  const edited = await ownerLinePositions(stack, false);

  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(2);
  const printed = await ownerLinePositions(preview, true);
  expect(printed.map(line => [line.text, line.page, line.inside]))
    .toEqual([['MERGE-A', 1, true], ['MERGE-C', 1, true], ['C-LINE-02', 1, true]]);
  expect(edited).toEqual(printed);
  expect(printed[1].y - printed[0].y).toBeGreaterThan(20);
});
