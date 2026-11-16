import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wordRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const widths = [1560, 1440, 1560, 1440];
const heads = ['HEAD-A', 'HEAD-B', 'HEAD-C', 'HEAD-D'];
const rows = ['01', '02', '03', '04', '05'];

type FooterFixture = 'none' | 'different-first-page' | 'empty';

const hwpCellLineNs = 'urn:word-studio:hwp-cell-lines:1';
const markupCompatibilityNs = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const fixedCellLine = '1 Fixed 13.3333 20 20 20';
const percentCellLine = '1 Percent 13.3333 21.3333 13.3 13.3';
const hwpGuards = {
  fixed: { policy: fixedCellLine, line: 300, single: false, allowed: true },
  percent: { policy: percentCellLine, line: 200, single: true, allowed: true },
  'percent-break': { policy: percentCellLine, line: 200, single: false, allowed: false },
  mixed: { policy: fixedCellLine, line: 300, single: false, allowed: false },
  'compact-owner': { policy: percentCellLine, line: 200, single: true, allowed: true },
  'terminal-owner': { policy: '1 Percent 13.3333 21.3 13.3 21.3', line: 320, single: false, allowed: false },
  'natural-owner': { policy: '1 Percent 13.3333 21.3 13.3 21.3', line: 320, single: false, allowed: true },
  'three-line-owner': { policy: '1 Percent 13.3333 21.3 13.3 21.3', line: 320, single: false, allowed: false },
  'compact-header': { policy: percentCellLine, line: 200, single: true, allowed: true },
  'mixed-header': { policy: percentCellLine, line: 200, single: true, allowed: false },
  'ordinary-terminal': { policy: '1 Percent 13.3333 21.3 13.3 21.3', line: 320, single: true, allowed: true },
  'ordinary-terminal-three': { policy: '1 Percent 13.3333 21.3 13.3 21.3', line: 320, single: true, allowed: false },
} as const;
type HwpGuardFixture = keyof typeof hwpGuards;

function hwpGuardAttribute(policy: string, line: number) {
  const data = JSON.stringify({ policy, font: 13.3333, line }).replaceAll('"', '&quot;');
  return ` xmlns:mc="${markupCompatibilityNs}" xmlns:wsh="${hwpCellLineNs}" mc:Ignorable="wsh" mc:PreserveAttributes="wsh:cellLine" wsh:cellLine="${data}"`;
}


async function fixture(alignment: 'top' | 'bottom' | 'center', footerFixture: FooterFixture, cLineCount = 1, edge: HeaderEdgeFixture = 'none', hwpGuard?: HwpGuardFixture) {
  const differentFirstPage = footerFixture === 'different-first-page';
  const hasFooter = footerFixture !== 'none';
  const headerGuard = hwpGuard === 'compact-header' || hwpGuard === 'mixed-header';
  const naturalOwner = hwpGuard === 'natural-owner' || hwpGuard === 'three-line-owner';
  const ownerA = hwpGuard === 'three-line-owner' ? 'MERGE-A OWNER-A OWNER-B'
    : naturalOwner ? 'MERGE-A OWNER-A' : 'MERGE-A';
  const ownerGuard = hwpGuard === 'compact-owner' || hwpGuard === 'terminal-owner' || naturalOwner;
  const ordinaryTerminal = hwpGuard === 'ordinary-terminal' || hwpGuard === 'ordinary-terminal-three';
  const row03Suffix = ordinaryTerminal ? (hwpGuard === 'ordinary-terminal-three' ? ' WIDEPARTA WIDEPARTB' : ' WIDEPARTA') : '';
  const paragraph = (text: string, line = 300) => {
    const marked = ordinaryTerminal ? text === 'ROW-03' + row03Suffix : headerGuard ? text === (hwpGuard === 'mixed-header' ? 'HEAD-A X' : 'HEAD-A') :
      hwpGuard === 'compact-owner' || naturalOwner ? text === ownerA :
      hwpGuard === 'terminal-owner' ? text.startsWith('MERGE-C\n') : /^ROW-03(?:\n|$)/.test(text);
    const guard = hwpGuard && marked ? hwpGuards[hwpGuard] : undefined;
    const textElements = text.split('\n').map(value => `<w:t>${value}</w:t>`);
    const run = (content: string, halfPoints = 20) => `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${halfPoints}"/></w:rPr>${content}</w:r>`;
    const runs = !text ? '' : guard && hwpGuard === 'mixed-header'
      ? run('<w:t xml:space="preserve">HEAD-A </w:t>') + run('<w:t>X</w:t>', 16)
      : guard && hwpGuard === 'mixed'
      ? run(textElements[0] + '<w:br/>') + run(textElements[1], 16)
      : run(textElements.join('<w:br/>'));
    return `<w:p${guard ? hwpGuardAttribute(guard.policy, guard.line) : ''}><w:pPr>
      <w:keepNext w:val="off"/><w:keepLines w:val="off"/><w:widowControl w:val="off"/>
      <w:spacing w:before="0" w:after="0" w:line="${guard?.line ?? line}" w:lineRule="exact"/><w:jc w:val="left"/>
      </w:pPr>${runs}</w:p>`;
  };
  const ordinary = edge === 'thick-ordinary' || edge === 'thick-ordinary-asym' || edge === 'thick-ordinary-bottom';
  const ordinaryAlign = ordinary ? 'center' : 'top';
  const continuationMargins = (top: number, bottom: number) => ordinary && edge !== 'thick-ordinary'
    ? `<w:tcMar><w:top w:w="${top}" w:type="dxa"/><w:bottom w:w="${bottom}" w:type="dxa"/></w:tcMar>` : '';
  const thickHeader = ordinary || edge === 'thick' || edge === 'thick-asym' || edge === 'thick-mixed';
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
      cell(0, index === 0 ? ownerA : '', merge, index === 0 && edge !== 'thick-mixed' ? alignment : 'top', index === 0 ? ownerMargins : ''),
      cell(1, `ROW-${id}${index === 2 ? row03Suffix : ''}${ordinary && index === 2 && !ownerGuard && (!hwpGuard || !hwpGuards[hwpGuard].single) ? '\nR3-LINE-02' : ''}`, '',
        index === 2 && (edge === 'bottom-on-continuation' || edge === 'thick-ordinary-bottom') ? 'bottom' : ordinaryAlign,
        index === 2 ? continuationMargins(90, 15) : ''),
      cell(2, index === 0 ? centeredCLines(cLineCount).join('\n') : index === 4 ? 'SINGLE-C' : '',
        index === 4 ? '' : merge, index === 4 ? ordinaryAlign : index === 0 && alignment === 'center' ? 'center' : 'top', index === 0 ? ownerMargins : ''),
      cell(3, `CHECK-${id}`, '', ordinaryAlign, index === 2 ? continuationMargins(15, 90) : ''),
    ]);
  }).join('');
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side =>
    `<w:${side} w:val="single" w:sz="4" w:color="000000"/>`).join('');
  const margins = ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="30" w:type="dxa"/>`).join('');
  const table = `<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders>
    <w:tblLayout w:type="fixed"/><w:tblCellMar>${margins}</w:tblCellMar></w:tblPr>
    <w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
    ${row(heads.map((head, index) => cell(index, index === 0 && hwpGuard === 'mixed-header' ? 'HEAD-A X' : head,
      '', index === 0 && headerGuard ? 'center' : 'top', headerBottom)), true)}${bodyRows}</w:tbl>`;
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
      /^(?:HEAD-[A-D]|(?:ROW|CHECK)-0[1-5]|MERGE-[AC]|SINGLE-C|TAIL|ROW-03\s*(?:R3-LINE-02|WIDEPARTA(?:\s*WIDEPARTB)?))$/.test(p.textContent ?? ''));
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
          text: p.textContent!.replace(/^ROW-03\s*(?:R3-LINE-02|WIDEPARTA(?:\s*WIDEPARTB)?)$/, 'ROW-03'), page, y,
          inside: ink.width > 0 && ink.height > 0 && (!cell || (ink.top >= cell.top - 1 &&
            ink.bottom <= cell.bottom + 1 && ink.left >= cell.left - 1 && ink.right <= cell.right + 1)),
        };
      }),
    };
  }, print);
}

async function openFixture(page: Page, alignment: 'top' | 'bottom' | 'center', footer: FooterFixture = 'none', cLineCount = 1, edge: HeaderEdgeFixture = 'none', hwpGuard?: HwpGuardFixture) {
  await page.goto('/');
  const name = `protected-crossed-${alignment}`;
  await page.locator('input[type=file]').first().setInputFiles({
    name: `${name}.docx`, buffer: await fixture(alignment, footer, cLineCount, edge, hwpGuard),
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

function ownerLinePositions(root: Locator, print: boolean, withX = false) {
  return root.evaluate(async (scope, { print, withX }) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const paper = scope as HTMLElement;
    const editor = paper.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    const top = (editor ?? paper).getBoundingClientRect().top;
    const scale = documentScale(editor ?? paper);
    const stride = Number(paper.dataset.pageHeight) + 24;
    const lines: { text: string; page: number; y: number; x?: number; inside: boolean }[] = [];
    const walker = document.createTreeWalker(paper, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      const paragraph = node.parentElement?.closest('p');
      const cell = paragraph?.closest('td');
      if (!cell || paragraph!.closest('.pagination-headers,[data-flow-widget]')) continue;
      for (const match of node.data.matchAll(/MERGE-[AC]|C-LINE-0[2-9]|OWNER-[AB]/g)) {
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
          ...(withX ? { x: (ink.left - (sheet ?? editor ?? paper).getBoundingClientRect().left) /
            (sheet ? documentScale(sheet) : scale) } : {}),
          inside: ink.width > 0 && ink.height > 0 && ink.left >= rect.left - 1 && ink.right <= rect.right + 1 &&
            ink.top >= rect.top - 1 && ink.bottom <= rect.bottom + 1,
        });
      }
    }
    return lines;
  }, { print, withX });
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


type HeaderEdgeFixture = 'none' | 'thick' | 'thick-asym' | 'thick-mixed' | 'bottom-on-continuation'
  | 'thick-ordinary' | 'thick-ordinary-asym' | 'thick-ordinary-bottom';

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

function continuationLinePositions(root: Locator, print: boolean, withX = false) {
  return root.evaluate(async (scope, { print, withX }) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const paper = scope as HTMLElement;
    const editor = paper.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    const origin = (editor ?? paper).getBoundingClientRect().top;
    const editScale = documentScale(editor ?? paper);
    const stride = Number(paper.dataset.pageHeight) + 24;
    const lines: { text: string; page: number; y: number; x?: number; inside: boolean }[] = [];
    const shape: { bodyParagraphs: number; hardBreaks: number }[] = [];

    for (const paragraph of paper.querySelectorAll<HTMLParagraphElement>('td > p')) {
      if (paragraph.closest('.pagination-headers,[data-flow-widget]')) continue;
      const text = paragraph.textContent?.replace(/\s+/g, '');
      if (text !== 'ROW-03' && text !== 'ROW-03R3-LINE-02' && text !== 'CHECK-03' && !/^ROW-03WIDEPARTA(?:WIDEPARTB)?$/.test(text ?? '')) continue;
      const cell = paragraph.closest('td')!;
      if (text !== 'CHECK-03') {
        shape.push({
          bodyParagraphs: cell.querySelectorAll(':scope > p').length,
          hardBreaks: paragraph.querySelectorAll('br:not(.ProseMirror-trailingBreak)').length,
        });
      }
      const sheet = print ? paragraph.closest<HTMLElement>('.print-page') : null;
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        for (const match of node.data.matchAll(/ROW-03|R3-LINE-02|WIDEPART[AB]|CHECK-03/g)) {
          const range = document.createRange();
          range.setStart(node, match.index!);
          range.setEnd(node, match.index! + match[0].length);
          const ink = range.getBoundingClientRect(), box = cell.getBoundingClientRect();
          const local = (ink.top - origin) / editScale;
          const page = sheet ? Number(sheet.dataset.page) : 1 + Math.floor((local + .5) / stride);
          const y = sheet ? (ink.top - sheet.getBoundingClientRect().top) / documentScale(sheet)
            : local - (page - 1) * stride;
          lines.push({
            text: match[0], page, y,
            ...(withX ? { x: (ink.left - (sheet ?? editor ?? paper).getBoundingClientRect().left) /
              (sheet ? documentScale(sheet) : editScale) } : {}),
            inside: ink.width > 0 && ink.height > 0 && ink.left >= box.left - 1 &&
              ink.right <= box.right + 1 && ink.top >= box.top - 1 && ink.bottom <= box.bottom + 1,
          });
        }
      }
    }
    return { lines, shape };
  }, { print, withX });
}

test('center alignment follows each ordinary continuation cell and its exact lines', async ({ page }) => {
  const lineLabels = ['ROW-03', 'R3-LINE-02', 'CHECK-03'];
  const boxes: { midpoint: number; check: number; row4: number }[] = [];
  const variants = ['thick-ordinary', 'thick-ordinary-asym', 'thick-ordinary-bottom'] as const;
  for (const edge of variants) await test.step(edge, async () => {
    const atomic = edge === 'thick-ordinary-bottom';
    const expectedPages = atomic ? [2, 2, 2, 2, 2] : [1, 1, 2, 2, 2];
    const expectedHeaders = atomic ? [2] : [1, 2];
    const rowMarkers = (table: Awaited<ReturnType<typeof snapshot>>, prefix: 'ROW' | 'CHECK') =>
      table.markers.filter(marker => marker.text.startsWith(`${prefix}-`));
    await openFixture(page, 'center', 'none', 2, edge);
    const editor = page.locator('.document-stack');
    await expect.poll(async () => rowMarkers(await snapshot(editor, false), 'ROW').map(marker => marker.page))
      .toEqual(expectedPages);
    if (atomic) await expect(editor.locator('.pagination-headers th')).toHaveCount(0);
    else await expect(editor.locator('.pagination-headers th')).toHaveText(heads);
    const read = async (root: Locator, print: boolean) => ({
      table: await snapshot(root, print),
      continuation: await continuationLinePositions(root, print),
      owners: await ownerLinePositions(root, print),
    });
    const edited = await read(editor, false);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    const printed = await read(preview, true);

    for (const mode of [edited, printed]) {
      expect(mode.table.pages).toBe(2);
      for (const prefix of ['ROW', 'CHECK'] as const) {
        const markers = rowMarkers(mode.table, prefix);
        expect(markers.map(marker => marker.text)).toEqual(rows.map(id => `${prefix}-${id}`));
        expect(markers.map(marker => marker.page)).toEqual(expectedPages);
        expect(markers.every(marker => marker.inside)).toBe(true);
      }
      for (const label of heads)
        expect(mode.table.markers.filter(marker => marker.text === label).map(marker => marker.page).sort())
          .toEqual(expectedHeaders);
      expect(mode.owners.map(owner => [owner.text, owner.page, owner.inside])).toEqual(
        ['MERGE-A', 'MERGE-C', 'C-LINE-02'].map(label => [label, atomic ? 2 : 1, true]),
      );
      expect(mode.continuation.shape).toEqual([{ bodyParagraphs: 1, hardBreaks: 1 }]);
      const [first, second, check] = mode.continuation.lines;
      expect(mode.continuation.lines.map(line => [line.text, line.page, line.inside]))
        .toEqual(lineLabels.map(label => [label, 2, true]));
      expect(second.y - first.y).toBeCloseTo(20, 1);
      const midpoint = (first.y + second.y) / 2;
      // Compare the two-line text block's midpoint with the single-line
      // text in the same physical row. Comparing only first lines is wrong.
      if (atomic) expect(first.y).toBeGreaterThan(check.y);
      else expect(midpoint - check.y).toBeCloseTo(edge === 'thick-ordinary-asym' ? 5 : 0, 1);
    }
    for (const [index, line] of edited.continuation.lines.entries())
      expect(Math.abs(line.y - printed.continuation.lines[index].y), edge + ' ' + line.text).toBeLessThan(0.5);
    for (const [index, owner] of edited.owners.entries())
      expect(Math.abs(owner.y - printed.owners[index].y), edge + ' ' + owner.text).toBeLessThan(0.5);
    for (const prefix of ['ROW', 'CHECK'] as const) {
      const editorRows = rowMarkers(edited.table, prefix), printRows = rowMarkers(printed.table, prefix);
      for (const [index, mark] of editorRows.entries())
        expect(Math.abs(mark.y - printRows[index].y), edge + ' ' + mark.text).toBeLessThan(0.5);
    }
    if (!atomic) {
      const [first, second, check] = edited.continuation.lines;
      boxes.push({ midpoint: (first.y + second.y) / 2, check: check.y, row4: rowMarkers(edited.table, 'ROW')[3].y });
    }
  });
  // Both explicit cells gain 45 twips of total padding. The 3px higher row
  // contributes 1.5px uniformly to centering, the 75-twip imbalance then
  // biases the respective cells 2.5px in opposite directions.
  const [balanced, asymmetric] = boxes;
  expect(asymmetric.midpoint - balanced.midpoint).toBeCloseTo(4, 1);
  expect(asymmetric.check - balanced.check).toBeCloseTo(-1, 1);
  expect(asymmetric.row4 - balanced.row4).toBeCloseTo(3, 1);
});

async function exportedHwpGuard(page: Page, xml: string) {
  return page.evaluate(({ xml, word, hwpCellLineNs, markupCompatibilityNs }) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved DOCX XML');
    return Array.from(document.getElementsByTagNameNS(word, 'p'))
      .filter(p => p.hasAttributeNS(hwpCellLineNs, 'cellLine')).map(p => {
        const saved = JSON.parse(p.getAttributeNS(hwpCellLineNs, 'cellLine')!);
        const props = Array.from(p.children).find(child => child.namespaceURI === word && child.localName === 'pPr');
        const spacing = props?.getElementsByTagNameNS(word, 'spacing')[0];
        const prefix = p.lookupPrefix(hwpCellLineNs);
        const ignored = p.getAttributeNS(markupCompatibilityNs, 'Ignorable')?.split(/\s+/);
        const preserved = p.getAttributeNS(markupCompatibilityNs, 'PreserveAttributes')?.split(/\s+/);
        return {
          policy: saved.policy,
          policyLine: saved.line,
          standardLine: Number(spacing?.getAttributeNS(word, 'line')),
          standardRule: spacing?.getAttributeNS(word, 'lineRule'),
          ignorable: !!prefix && !!ignored?.includes(prefix),
          preserved: !!prefix && !!preserved?.includes(prefix + ':cellLine'),
          labels: Array.from(p.getElementsByTagNameNS(word, 't')).map(t => t.textContent),
          breaks: p.getElementsByTagNameNS(word, 'br').length,
        };
      });
  }, { xml, word, hwpCellLineNs, markupCompatibilityNs });
}

for (const name of ['fixed', 'percent', 'percent-break', 'mixed'] as const) {
  test(`protected ordinary HWP cell line ${name} keeps its safe row placement`, async ({ page }, info) => {
    const guard = hwpGuards[name];
    const expectedRows = guard.allowed ? [1, 1, 2, 2, 2] : [2, 2, 2, 2, 2];
    const expectedLines = guard.single ? ['ROW-03', 'CHECK-03'] : ['ROW-03', 'R3-LINE-02', 'CHECK-03'];
    await openFixture(page, 'center', 'none', 2, 'thick-ordinary', name);
    const editor = page.locator('.document-stack');
    const assertPolicy = async (root: Locator) => {
      const active = root.locator('td > p[data-hwp-cell-line]');
      await expect(active).toHaveCount(1);
      await expect(active).toHaveAttribute('data-hwp-cell-line', guard.policy);
      await expect(active).toHaveAttribute('data-word-line-rule', 'exact');
      if (guard.allowed)
        expect(await active.evaluate(p => parseFloat(getComputedStyle(p).lineHeight)))
          .toBeCloseTo(guard.single ? 13.3 : 20, 1);
    };
    const placements = async (root: Locator, print: boolean) => {
      const table = await snapshot(root, print);
      return { table, continuation: await continuationLinePositions(root, print),
        owners: await ownerLinePositions(root, print) };
    };
    await assertPolicy(editor);
    await expect.poll(async () => (await snapshot(editor, false)).markers
      .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page)).toEqual(expectedRows);
    const edited = await placements(editor, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    await assertPolicy(preview);
    const printed = await placements(preview, true);

    for (const mode of [edited, printed]) {
      expect(mode.table.pages).toBe(2);
      for (const prefix of ['ROW-', 'CHECK-']) {
        const bodyRows = mode.table.markers.filter(marker => marker.text.startsWith(prefix));
        expect(bodyRows.map(marker => marker.text)).toEqual(rows.map(id => prefix + id));
        expect(bodyRows.map(marker => [marker.page, marker.inside])).toEqual(expectedRows.map(p => [p, true]));
      }
      for (const head of heads)
        expect(mode.table.markers.filter(marker => marker.text === head).map(marker => marker.page).sort())
          .toEqual(guard.allowed ? [1, 2] : [2]);
      expect(mode.owners.map(owner => [owner.text, owner.page, owner.inside])).toEqual(
        ['MERGE-A', 'MERGE-C', 'C-LINE-02'].map(label => [label, guard.allowed ? 1 : 2, true]),
      );
      expect(mode.continuation.shape).toEqual([{ bodyParagraphs: 1, hardBreaks: guard.single ? 0 : 1 }]);
      expect(mode.continuation.lines.map(line => [line.text, line.page, line.inside]))
        .toEqual(expectedLines.map(label => [label, 2, true]));
      if (name === 'fixed')
        expect(mode.continuation.lines[1].y - mode.continuation.lines[0].y).toBeCloseTo(20, 1);
    }
    for (const [index, line] of edited.continuation.lines.entries())
      expect(Math.abs(line.y - printed.continuation.lines[index].y), line.text).toBeLessThan(0.5);
    for (const [index, owner] of edited.owners.entries())
      expect(Math.abs(owner.y - printed.owners[index].y), owner.text).toBeLessThan(0.5);
    const printedRows = new Map(printed.table.markers.filter(marker => /^(ROW|CHECK)-/.test(marker.text))
      .map(marker => [marker.text, marker.y]));
    for (const mark of edited.table.markers.filter(marker => /^(ROW|CHECK)-/.test(marker.text)))
      expect(Math.abs(mark.y - printedRows.get(mark.text)!), mark.text).toBeLessThan(0.5);

    if (guard.allowed) {
      await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const output = info.outputPath(`textbox-save-protected-hwp-guard-${name}.docx`);
      await (await pending).saveAs(output);
      const pkg = await JSZip.loadAsync(await readFile(output));
      const xml = await pkg.file('word/document.xml')!.async('string');
      expect(await exportedHwpGuard(page, xml)).toEqual([{
        policy: guard.policy, policyLine: guard.line, standardLine: guard.line, standardRule: 'exact',
        ignorable: true, preserved: true,
        labels: guard.single ? ['ROW-03'] : ['ROW-03', 'R3-LINE-02'],
        breaks: guard.single ? 0 : 1,
      }]);
    }
  });
}

function compactOwnerBox(active: Locator) {
  return active.evaluate(async paragraph => {
    const { documentScale } = await import('/src/document-scale.ts');
    const cell = paragraph.closest('td')!, sheet = paragraph.closest<HTMLElement>('.print-page');
    const paper = sheet ?? paragraph.closest<HTMLElement>('.tiptap[contenteditable=true]')!;
    const box = cell.getBoundingClientRect(), origin = paper.getBoundingClientRect();
    const scale = documentScale(paper);
    const css = getComputedStyle(cell);
    const border = getComputedStyle(cell.closest('table')!).borderCollapse === 'collapse' ? .5 : 1;
    const inset = (padding: string, width: string) => ((parseFloat(padding) || 0) + (parseFloat(width) || 0) * border) * scale;
    const visibleBody = sheet?.querySelector('.print-content')?.getBoundingClientRect();
    const left = Math.max(box.left + inset(css.paddingLeft, css.borderLeftWidth), visibleBody?.left ?? -Infinity);
    const top = Math.max(box.top + inset(css.paddingTop, css.borderTopWidth), visibleBody?.top ?? -Infinity);
    const right = Math.min(box.right - inset(css.paddingRight, css.borderRightWidth), visibleBody?.right ?? Infinity);
    const bottom = Math.min(box.bottom - inset(css.paddingBottom, css.borderBottomWidth), visibleBody?.bottom ?? Infinity);
    const range = document.createRange();
    range.selectNodeContents(paragraph);
    const ink = range.getBoundingClientRect();
    return {
      rowSpan: cell.rowSpan, breaks: paragraph.querySelectorAll('br:not(.ProseMirror-trailingBreak)').length,
      terminal: (paragraph as HTMLElement).dataset.hwpCellTerminal === 'true',
      prefix: !!paragraph.querySelector('[data-hwp-cell-prefix]'),
      prefixText: [...paragraph.querySelectorAll('[data-hwp-cell-prefix]')]
        .filter(node => !node.parentElement?.closest('[data-hwp-cell-prefix]')).map(node => node.textContent).join(''),
      advance: Number((paragraph as HTMLElement).dataset.hwpCellAdvance || 0),
      height: paragraph.getBoundingClientRect().height / scale,
      x: (ink.left - origin.left) / scale,
      visible: ink.width > 0 && ink.height > 0 && ink.left >= left - .5 && ink.right <= right + .5 &&
        ink.top >= top - .5 && ink.bottom <= bottom + .5,
    };
  });
}

for (const name of ['compact-owner', 'terminal-owner'] as const) {
  test(`protected merged HWP ${name} preserves the physical owner ink`, async ({ page }) => {
    const terminal = name === 'terminal-owner';
    const expectedRows = terminal ? [2, 2, 2, 2, 2] : [1, 1, 2, 2, 2];
    const ownerPaper = terminal ? 2 : 1;
    const rowPages = async (root: Locator, print: boolean) => (await snapshot(root, print)).markers
      .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page);
    await openFixture(page, 'center', 'none', 2, 'thick-ordinary', name);
    const stack = page.locator('.document-stack');
    await expect.poll(() => rowPages(stack, false)).toEqual(expectedRows);
    const read = async (root: Locator, print: boolean) => {
      const active = root.locator('td > p[data-hwp-cell-line]');
      await expect(active).toHaveCount(1);
      await expect(active).toHaveAttribute('data-hwp-cell-line', hwpGuards[name].policy);
      const lines = await ownerLinePositions(root, print);
      expect(lines.map(line => [line.text, line.page, line.inside])).toEqual(
        ['MERGE-A', 'MERGE-C', 'C-LINE-02'].map(label => [label, ownerPaper, true]));
      const box = await compactOwnerBox(active);
      expect(box).toMatchObject({ breaks: terminal ? 1 : 0, terminal, prefix: terminal, visible: true });
      expect(Math.abs(box.height - (terminal ? 34.6 : 13.3))).toBeLessThan(.1);
      expect(lines[2].y - lines[1].y).toBeCloseTo(terminal ? 21.3 : 20, 1);
      if (terminal) expect(box.advance).toBeCloseTo(21.3, 1);
      return { box, lines };
    };
    const edited = await read(stack, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    expect(await rowPages(preview, true)).toEqual(expectedRows);
    const printed = await read(preview, true);
    if (!terminal) {
      expect(printed.box.rowSpan).toBe(2);
      const tail = preview.locator('.print-page[data-page="2"] [data-print-merge-continuation]');
      await expect(tail).toHaveCount(2);
      await expect(tail).toHaveText(['', '']);
    }
    expect(Math.abs(edited.box.x - printed.box.x)).toBeLessThan(.5);
    // ownerLinePositions normalizes y to the physical paper in both modes.
    for (const [i, line] of edited.lines.entries())
      expect(Math.abs(line.y - printed.lines[i].y), line.text).toBeLessThan(.5);
  });
}

for (const name of ['natural-owner', 'three-line-owner'] as const) {
  test(`protected merged HWP ${name} preserves naturally wrapped terminal ink`, async ({ page }) => {
    const atomic = name === 'three-line-owner';
    const targetRows = atomic ? [2, 2, 2, 2, 2] : [1, 1, 2, 2, 2];
    const ownerLabels = atomic ? ['MERGE-A', 'OWNER-A', 'OWNER-B'] : ['MERGE-A', 'OWNER-A'];
    const allLabels = [...ownerLabels, 'MERGE-C', 'C-LINE-02'];
    const rowPages = async (root: Locator, print: boolean) => (await snapshot(root, print)).markers
      .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page);
    await openFixture(page, 'center', 'none', 2, 'thick-ordinary', name);
    const stack = page.locator('.document-stack');
    await expect.poll(() => rowPages(stack, false)).toEqual(targetRows);
    const read = async (root: Locator, print: boolean) => {
      const active = root.locator('td > p[data-hwp-cell-line]');
      await expect(active).toHaveCount(1);
      await expect(active).toHaveAttribute('data-hwp-cell-line', hwpGuards[name].policy);
      await expect(active).toHaveAttribute('data-word-line-rule', 'exact');
      expect(await active.textContent()).toBe(ownerLabels.join(' '));
      const lines = await ownerLinePositions(root, print, true);
      expect(lines.map(line => [line.text, line.page, line.inside]))
        .toEqual(allLabels.map(text => [text, atomic ? 2 : 1, true]));
      const box = await compactOwnerBox(active);
      expect(box).toMatchObject({
        breaks: 0, terminal: true, prefix: true, visible: true,
        prefixText: atomic ? 'MERGE-A OWNER-A ' : 'MERGE-A ',
      });
      expect(box.advance).toBeCloseTo(21.3, 1);
      expect(box.height).toBeCloseTo((ownerLabels.length - 1) * 21.3 + 13.3, 1);
      for (let i = 1; i < ownerLabels.length; i++)
        expect(lines[i].y - lines[i - 1].y).toBeCloseTo(21.3, 1);
      return { lines, rowSpan: box.rowSpan };
    };
    const edited = await read(stack, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    expect(await rowPages(preview, true)).toEqual(targetRows);
    const printed = await read(preview, true);
    expect(printed.rowSpan).toBe(atomic ? 5 : 2);
    for (const [index, line] of edited.lines.entries()) {
      const other = printed.lines[index];
      expect(Math.abs(line.x! - other.x!), line.text + ' x').toBeLessThan(.5);
      expect(Math.abs(line.y - other.y), line.text + ' y').toBeLessThan(.5);
    }
  });
}

function repeatingHeaderInk(root: Locator) {
  return root.evaluate(async scope => {
    const { documentScale } = await import('/src/document-scale.ts');
    const paper = scope as HTMLElement;
    const editor = paper.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    return [...paper.querySelectorAll<HTMLParagraphElement>('th > p')].map(p => {
      const sheet = p.closest<HTMLElement>('.print-page');
      const origin = (sheet ?? editor ?? paper).getBoundingClientRect();
      const scale = documentScale(sheet ?? editor ?? paper);
      const stride = Number(paper.dataset.pageHeight) + 24;
      const range = document.createRange();
      range.selectNodeContents(p);
      const ink = range.getBoundingClientRect(), cell = p.closest('th')!.getBoundingClientRect();
      const relativeY = (ink.top - origin.top) / scale;
      const page = sheet ? Number(sheet.dataset.page) : 1 + Math.floor((relativeY + .5) / stride);
      const policy = p.dataset.hwpCellLine;
      const sizes = new Set<number>();
      if (policy) {
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) if (walker.currentNode.textContent?.trim())
          sizes.add(parseFloat(getComputedStyle(walker.currentNode.parentElement!).fontSize));
      }
      return {
        text: p.textContent, page, overlay: !!p.closest('.pagination-headers'),
        x: (ink.left - origin.left) / scale, y: sheet ? relativeY : relativeY - (page - 1) * stride,
        policy, rule: p.dataset.wordLineRule, sizes: [...sizes],
        font: parseFloat(getComputedStyle(p).fontSize),
        height: p.getBoundingClientRect().height / scale,
        inside: ink.width > 0 && ink.height > 0 && ink.top >= cell.top - .5 && ink.bottom <= cell.bottom + .5 &&
          ink.left >= cell.left - .5 && ink.right <= cell.right + .5,
      };
    });
  });
}

for (const name of ['compact-header', 'mixed-header'] as const) {
  test(`compact repeated HWP head ${name} uses only measured whole rows`, async ({ page }) => {
    const atomic = name === 'mixed-header';
    const targetRows = atomic ? [2, 2, 2, 2, 2] : [1, 1, 2, 2, 2];
    const targetPages = atomic ? [2] : [1, 2];
    const labels = [atomic ? 'HEAD-A X' : 'HEAD-A', ...heads.slice(1)];
    const bodyRows = (table: Awaited<ReturnType<typeof snapshot>>) => table.markers
      .filter(marker => /^(ROW|CHECK)-/.test(marker.text));
    await openFixture(page, 'center', 'none', 2, 'thick-ordinary', name);
    const stack = page.locator('.document-stack');
    await expect.poll(async () => bodyRows(await snapshot(stack, false))
      .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page)).toEqual(targetRows);
    if (atomic) await expect(stack.locator('.pagination-headers th')).toHaveCount(0);
    else await expect(stack.locator('.pagination-headers th')).toHaveText(labels);
    const read = async (root: Locator, print: boolean) => ({
      head: await repeatingHeaderInk(root), body: bodyRows(await snapshot(root, print)),
      owners: await ownerLinePositions(root, print),
    });
    const edited = await read(stack, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    const printed = await read(preview, true);

    for (const [mode, print] of [[edited, false], [printed, true]] as const) {
      expect(mode.head.map(head => [head.page, head.text])).toEqual(targetPages.flatMap(p => labels.map(label => [p, label])));
      for (const marker of mode.head) {
        expect(marker.inside).toBe(true);
        expect(marker.overlay).toBe(!print && !atomic && marker.page === 2);
        if (marker.text !== labels[0]) { expect(marker.policy).toBeUndefined(); continue; }
        expect(marker.policy).toBe(percentCellLine);
        expect(marker.rule).toBe('exact');
        expect(marker.font).toBeCloseTo(40 / 3, 1);
        expect(marker.sizes).toHaveLength(atomic ? 2 : 1);
        expect(marker.sizes[0]).toBeCloseTo(40 / 3, 1);
        if (atomic) expect(marker.sizes[1]).toBeCloseTo(32 / 3, 1);
        else expect(marker.height).toBeCloseTo(13.3, 1);
      }
      for (const prefix of ['ROW-', 'CHECK-'])
        expect(mode.body.filter(marker => marker.text.startsWith(prefix)).map(marker => [marker.text, marker.page, marker.inside]))
          .toEqual(rows.map((id, index) => [prefix + id, targetRows[index], true]));
      expect(mode.owners.map(owner => [owner.text, owner.page, owner.inside]))
        .toEqual(['MERGE-A', 'MERGE-C', 'C-LINE-02'].map(label => [label, atomic ? 2 : 1, true]));
    }
    for (const [index, header] of printed.head.entries()) {
      expect(Math.abs(header.x - edited.head[index].x), header.text + ' x').toBeLessThan(.5);
      expect(Math.abs(header.y - edited.head[index].y), header.text + ' y').toBeLessThan(.5);
    }
    for (const [index, marker] of printed.body.entries())
      expect(Math.abs(marker.y - edited.body[index].y), marker.text).toBeLessThan(.5);
  });
}


for (const name of ['ordinary-terminal', 'ordinary-terminal-three'] as const) {
  test(`natural HWP ${name} verifies rows at the repeated-head seam`, async ({ page }) => {
    const atomic = name === 'ordinary-terminal-three';
    const expectedPages = atomic ? [2, 2, 2, 2, 2] : [1, 1, 2, 2, 2];
    const labels = atomic ? ['ROW-03', 'WIDEPARTA', 'WIDEPARTB'] : ['ROW-03', 'WIDEPARTA'];
    const body = (table: Awaited<ReturnType<typeof snapshot>>) => table.markers
      .filter(marker => /^(ROW|CHECK)-/.test(marker.text));
    await openFixture(page, 'center', 'none', 2, 'thick-ordinary', name);
    const stack = page.locator('.document-stack');
    await expect.poll(async () => body(await snapshot(stack, false))
      .filter(marker => marker.text.startsWith('ROW-')).map(marker => marker.page)).toEqual(expectedPages);
    const read = async (root: Locator, print: boolean) => {
      const active = root.locator('td > p[data-hwp-cell-line]');
      await expect(active).toHaveCount(1);
      await expect(active).toHaveAttribute('data-hwp-cell-line', hwpGuards[name].policy);
      await expect(active).toHaveAttribute('data-word-line-rule', 'exact');
      expect(await active.textContent()).toBe(labels.join(' '));
      const ink = await compactOwnerBox(active);
      expect(ink).toMatchObject({
        rowSpan: 1, breaks: 0, terminal: true, visible: true, prefix: true,
        prefixText: labels.slice(0, -1).join(' ') + ' ',
      });
      expect(ink.advance).toBeCloseTo(21.3, 1);
      expect(ink.height).toBeCloseTo((labels.length - 1) * 21.3 + 13.3, 1);
      const measured = await continuationLinePositions(root, print, true);
      expect(measured.shape).toEqual([{ bodyParagraphs: 1, hardBreaks: 0 }]);
      expect(measured.lines.map(line => [line.text, line.page, line.inside]))
        .toEqual([...labels, 'CHECK-03'].map(label => [label, 2, true]));
      for (let i = 1; i < labels.length; i++)
        expect(measured.lines[i].y - measured.lines[i - 1].y).toBeCloseTo(21.3, 1);
      const table = await snapshot(root, print);
      expect(table.pages).toBe(2);
      const rowsAndChecks = body(table);
      for (const prefix of ['ROW-', 'CHECK-'])
        expect(rowsAndChecks.filter(marker => marker.text.startsWith(prefix)).map(m => [m.text, m.page, m.inside]))
          .toEqual(rows.map((id, i) => [prefix + id, expectedPages[i], true]));
      const header = await repeatingHeaderInk(root);
      expect(header.map(label => [label.page, label.text, label.inside])).toEqual(
        (atomic ? [2] : [1, 2]).flatMap(page => heads.map(label => [page, label, true])),
      );
      expect(header.filter(label => label.overlay).map(label => label.page))
        .toEqual(print || atomic ? [] : [2, 2, 2, 2]);
      return { lines: measured.lines, rowsAndChecks, header };
    };
    const edited = await read(stack, false);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(2);
    const printed = await read(preview, true);
    for (const field of ['lines', 'header'] as const)
      for (const [i, marker] of edited[field].entries()) {
        expect(Math.abs(marker.x! - printed[field][i].x!), marker.text + ' x').toBeLessThan(.5);
        expect(Math.abs(marker.y - printed[field][i].y), marker.text + ' y').toBeLessThan(.5);
      }
    for (const [i, marker] of edited.rowsAndChecks.entries())
      expect(Math.abs(marker.y - printed.rowsAndChecks[i].y), marker.text).toBeLessThan(.5);
  });
}

test('live print paper rolls back a planned ordinary terminal split if its geometry changes', async ({ page }) => {
  await openFixture(page, 'center', 'none', 2, 'thick-ordinary', 'ordinary-terminal');
  const preview = page.locator('.print-layout');
  const seam = async (expected: number[]) => {
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    await expect(preview.locator('.print-page')).toHaveCount(2);
    const table = await snapshot(preview, true);
    expect(table.markers.filter(m => /^ROW-/.test(m.text)).map(m => m.page)).toEqual(expected);
    expect((await repeatingHeaderInk(preview)).map(h => [h.page, h.text])).toEqual(
      (expected[0] === 1 ? [1,2] : [2]).flatMap(n => heads.map(label => [n, label])),
    );
  };
  await seam([1,1,2,2,2]);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  const force = await page.addStyleTag({ content: '.print-page[data-page="2"] td > p[data-hwp-cell-line] { max-width:86px !important }' });
  await seam([2,2,2,2,2]);
  await expect(preview.locator('.print-page[data-page="1"] table')).toHaveCount(0);
  await expect(preview.locator('[data-print-merge-continuation]')).toHaveCount(0);
  await expect(preview.locator('td > p[data-hwp-cell-line]')).toHaveText('ROW-03 WIDEPARTA');
  const compact = await compactOwnerBox(preview.locator('td > p[data-hwp-cell-line]'));
  expect(compact.prefixText).toBe('ROW-03 ');
  expect(compact.height).toBeCloseTo(34.6, 1);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  await force.evaluate(el => el.remove());
  await seam([1,1,2,2,2]);
});
