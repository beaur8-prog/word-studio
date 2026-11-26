import { expect, test, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import { fenceFixture, headLabels, rowLabels } from './protected-hwp-fence-fixture';

const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationship = 'http://schemas.openxmlformats.org/package/2006/relationships';
const shortHead = ['QA RUNNING HEADER'];
const shortFoot = ['QA RUNNING FOOTER'];
const longFoot = Array.from({ length: 6 }, (_, i) => 'QA FOOTER LINE ' + (i + 1));
const topMargin = 1417 / 15, bottomMargin = 1134 / 15;
type Scenario = 'empty' | 'short' | 'oversized-footer' | 'page-number';

async function storyFixture(scenario: Scenario) {
  const zip = await JSZip.loadAsync(await fenceFixture('absent'));
  const document = await zip.file('word/document.xml')!.async('string');
  const types = await zip.file('[Content_Types].xml')!.async('string');
  const references = ['header', 'footer'].map(kind =>
    `<w:${kind}Reference w:type="default" r:id="r-${kind}"/>`).join('');
  zip.file('word/document.xml', document.replace('<w:sectPr>', '<w:sectPr>' + references));
  zip.file('[Content_Types].xml', types.replace('</Types>', ['header', 'footer'].map(kind =>
    `<Override PartName="/word/${kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`).join('') + '</Types>'));
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="${packageRelationship}">
    <Relationship Id="r-header" Type="${relationship}/header" Target="header1.xml"/>
    <Relationship Id="r-footer" Type="${relationship}/footer" Target="footer1.xml"/>
  </Relationships>`);
  const paragraphs = (labels: string[], line = 180, size = 16) => labels.length ? labels.map(label =>
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="exact"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size}"/></w:rPr><w:t>${label}</w:t></w:r></w:p>`).join('') : '<w:p/>';
  zip.file('word/header1.xml', `<w:hdr xmlns:w="${word}">${paragraphs(scenario === 'empty' ? [] : shortHead)}</w:hdr>`);
  const footer = scenario === 'oversized-footer' ? paragraphs(longFoot, 320, 24)
    : scenario === 'page-number'
      ? paragraphs(shortFoot).replace('</w:p>', '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>')
      : paragraphs(scenario === 'empty' ? [] : shortFoot);
  zip.file('word/footer1.xml', `<w:ftr xmlns:w="${word}">${footer}</w:ftr>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

function storyLayout(root: Locator, print: boolean) {
  return root.evaluate(async (scope, { print, rowLabels, headLabels }) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const editor = scope.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    const editableCss = editor && getComputedStyle(editor);
    const sheets = [...scope.querySelectorAll<HTMLElement>(print ? '.print-page' : '.paper-backgrounds > .paper')];
    const bodyMarkers = [...(editor ?? scope).querySelectorAll<HTMLParagraphElement>(print ? '.print-content tr p' : 'tr p')];
    const tableHeads = [...scope.querySelectorAll<HTMLParagraphElement>(print ? '.print-content th p' : 'th p')];
    const ink = (p: HTMLElement) => {
      const range = document.createRange();
      range.selectNodeContents(p);
      return range.getBoundingClientRect();
    };
    return sheets.map((sheet, index) => {
      const bounds = sheet.getBoundingClientRect(), scale = documentScale(sheet);
      const number = index + 1, height = bounds.height / scale;
      const header = sheet.querySelector<HTMLElement>(print ? '.print-header' : '.paper-header')!;
      const footer = sheet.querySelector<HTMLElement>(print ? '.print-footer' : '.paper-footer')!;
      const physicalBody = print ? sheet.querySelector('.print-content')!.getBoundingClientRect() : null;
      // Edit draws one continuous body over these paper boxes. Its real CSS
      // padding supplies the inset at every physical separation.
      const body = physicalBody ? {
        top: (physicalBody.top - bounds.top) / scale, bottom: (physicalBody.bottom - bounds.top) / scale,
      } : {
        top: parseFloat(editableCss!.paddingTop), bottom: height - parseFloat(editableCss!.paddingBottom),
      };
      const story = (container: HTMLElement) => {
        const box = container.getBoundingClientRect();
        const visible = getComputedStyle(container).visibility === 'visible' && getComputedStyle(container).display !== 'none';
        const paragraphs = [...container.querySelectorAll<HTMLParagraphElement>('p')].filter(p => p.textContent?.trim());
        return {
          top: (box.top - bounds.top) / scale, bottom: (box.bottom - bounds.top) / scale,
          text: paragraphs.map(p => p.textContent),
          painted: paragraphs.every(p => {
            const rect = ink(p);
            return visible && rect.width > 0 && rect.height > 0 && rect.top >= bounds.top - .5 &&
              rect.bottom <= bounds.bottom + .5 && rect.left >= bounds.left - .5 && rect.right <= bounds.right + .5;
          }),
        };
      };
      const measureBody = (p: HTMLParagraphElement) => {
        const rect = ink(p), row = p.closest('tr')!.getBoundingClientRect();
        return { label: p.textContent!, x: (rect.left - bounds.left) / scale, top: (rect.top - bounds.top) / scale,
          bottom: (rect.bottom - bounds.top) / scale, rowBottom: (row.bottom - bounds.top) / scale };
      };
      const belongsHere = (p: HTMLParagraphElement) => {
        if (print) return p.closest('.print-page') === sheet;
        const rect = ink(p);
        return rect.top >= bounds.top - .5 && rect.top < bounds.bottom;
      };
      return {
        number, height, body, header: story(header), footer: story(footer),
        rows: bodyMarkers.filter(p => rowLabels.includes(p.textContent!) && belongsHere(p)).map(measureBody),
        tableHeads: tableHeads.filter(p => headLabels.includes(p.textContent!) && belongsHere(p)).map(measureBody),
      };
    });
  }, { print, rowLabels, headLabels });
}

test('protected HWP whole rows fit within the measured body of default header and footer stories', async ({ page }) => {
  test.setTimeout(90_000);
  let roomyFirstRowBottom: number | undefined;
  for (const scenario of ['empty', 'short', 'oversized-footer', 'page-number'] as const) await test.step(scenario, async () => {
    const oversized = scenario === 'oversized-footer';
    const keepTogether = oversized || scenario === 'page-number';
    const rowPapers = keepTogether ? [2, 2, 2, 2, 2, 2] : [1, 1, 1, 2, 2, 2];
    const headPapers = keepTogether ? [2] : [1, 2];
    const name = 'protected-default-stories-' + scenario;
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await storyFixture(scenario),
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => document.fonts.ready);
    const stack = page.locator('.document-stack');
    const rowPages = (pages: Awaited<ReturnType<typeof storyLayout>>) => pages.flatMap(p => p.rows.map(() => p.number));
    await expect.poll(async () => rowPages(await storyLayout(stack, false))).toEqual(rowPapers);
    const edited = await storyLayout(stack, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(3);
    const printed = await storyLayout(preview, true);
    for (const mode of [edited, printed]) {
      expect(mode).toHaveLength(3);
      expect(rowPages(mode)).toEqual(rowPapers);
      expect(mode.flatMap(p => p.rows.map(row => row.label))).toEqual(rowLabels);
      expect(mode.flatMap(p => p.tableHeads.map(head => [p.number, head.label]))).toEqual(
        headPapers.flatMap(p => headLabels.map(head => [p, head])),
      );
      for (const paper of mode) {
        const empty = scenario === 'empty';
        expect(paper.header.text).toEqual(empty ? [] : shortHead);
        expect(paper.footer.text).toEqual(empty ? [] : oversized ? longFoot :
          scenario === 'page-number' ? [shortFoot[0] + paper.number] : shortFoot);
        expect(paper.header.painted && paper.footer.painted).toBe(true);
        // Compare the actual positioned story boxes, not a guessed line count.
        const expectedTop = Math.max(topMargin, empty ? 0 : paper.header.bottom + 2);
        const expectedBottom = Math.max(bottomMargin, empty ? 0 : paper.height - paper.footer.top + 2);
        expect(Math.abs(paper.body.top - expectedTop)).toBeLessThan(.5);
        expect(Math.abs(paper.height - paper.body.bottom - expectedBottom)).toBeLessThan(.5);
        if (oversized) expect(expectedBottom - bottomMargin).toBeGreaterThan(80);
        for (const row of [...paper.rows, ...paper.tableHeads]) {
          expect(row.top).toBeGreaterThanOrEqual(paper.body.top - .5);
          expect(row.bottom).toBeLessThanOrEqual(paper.body.bottom + .5);
        }
      }
    }
    if (scenario === 'short') roomyFirstRowBottom = printed[0].rows[0].rowBottom;
    if (oversized) expect(printed[0].body.bottom).toBeLessThan(roomyFirstRowBottom!);
    for (const [i, paper] of printed.entries()) {
      expect(Math.abs(paper.body.top - edited[i].body.top)).toBeLessThan(.5);
      expect(Math.abs(paper.body.bottom - edited[i].body.bottom)).toBeLessThan(.5);
      for (const [j, row] of paper.rows.entries()) {
        expect(Math.abs(row.x - edited[i].rows[j].x), row.label + ' x').toBeLessThan(.5);
        expect(Math.abs(row.top - edited[i].rows[j].top), row.label + ' y').toBeLessThan(.5);
      }
    }
  });
});

type GridScenario = 'literal-grid' | 'tall-grid' | 'auto-grid';
const gridLabels = ['QA GRID HEADER', 'QA GRID FOOTER'];
const gridTwips = [[5295, 1440, 2850], [3210, 3210, 3210]];
const gridWidths = gridTwips.map(columns => columns.map(width => width / 15));

async function gridStoryFixture(scenario: GridScenario) {
  const zip = await JSZip.loadAsync(await storyFixture('short'));
  for (const index of [0, 1]) {
    const header = index === 0, tall = !header && scenario === 'tall-grid';
    const layout = !header && scenario === 'auto-grid' ? 'autofit' : 'fixed';
    const labels = header ? [gridLabels[0], '', ''] : ['', '', gridLabels[1]];
    const widths = gridTwips[index];
    const cells = widths.map((width, i) => {
      const indent = header && labels[i] ? '<w:ind w:left="213" w:hanging="213"/>' : '';
      const run = labels[i] ? `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/></w:rPr>
        ${'<w:br/>'.repeat(tall ? 2 : 1)}<w:t>${labels[i]}</w:t></w:r>` : '';
      return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
        <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/>${indent}
        <w:jc w:val="${header ? 'left' : 'right'}"/></w:pPr>${run}</w:p></w:tc>`;
    }).join('');
    const padding = ['top', 'left', 'bottom', 'right'].map(side =>
      `<w:${side} w:w="${side === 'top' || side === 'bottom' ? 28 : 102}" w:type="dxa"/>`).join('');
    const story = header ? 'hdr' : 'ftr', edge = header ? 'bottom' : 'top';
    zip.file(`word/${header ? 'header' : 'footer'}1.xml`, `<w:${story} xmlns:w="${word}">
      <w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((sum, width) => sum + width, 0)}" w:type="dxa"/>
        <w:tblBorders><w:${edge} w:val="single" w:sz="3" w:color="000000"/></w:tblBorders>
        <w:tblLayout w:type="${layout}"/><w:tblCellMar>${padding}</w:tblCellMar></w:tblPr>
        <w:tblGrid>${widths.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
        <w:tr><w:trPr><w:trHeight w:val="${header ? 363 : tall ? 650 : 370}" w:hRule="atLeast"/></w:trPr>
        ${cells}</w:tr></w:tbl></w:${story}>`);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

function storyGridLayout(root: Locator, print: boolean) {
  return root.evaluate(async (scope, print) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const papers = scope.querySelectorAll<HTMLElement>(print ? '.print-page' : '.paper-backgrounds > .paper');
    return [...papers].map(paper => {
      const reference = paper.getBoundingClientRect(), scale = documentScale(paper);
      const measure = (element: Element | DOMRect) => {
        const rect = element instanceof Element ? element.getBoundingClientRect() : element;
        return [(rect.left - reference.left) / scale, (rect.top - reference.top) / scale,
          rect.width / scale, rect.height / scale];
      };
      return [...paper.querySelectorAll<HTMLTableElement>(print ? '.print-header table,.print-footer table' :
        '.paper-header table,.paper-footer table')].map(table => {
        const row = table.rows[0], p = [...row.cells].map(cell => cell.querySelector('p')!).find(p => p.textContent?.trim())!;
        const range = document.createRange();
        range.selectNodeContents(p);
        return {
          layout: table.dataset.wordTableLayout, text: p.textContent!,
          rows: table.rows.length, breaks: p.querySelectorAll('br').length,
          authored: [...row.cells].map(cell => Number(cell.getAttribute('colwidth'))),
          box: measure(table), cols: [...table.querySelectorAll(':scope > colgroup > col')].map(measure),
          cells: [...row.cells].map(measure), ink: measure(range.getBoundingClientRect()),
        };
      });
    });
  }, print);
}

for (const scenario of ['literal-grid', 'tall-grid', 'auto-grid'] as const) {
  test(`default story ${scenario} preserves its geometry and protected table placement`, async ({ page }) => {
    const split = scenario === 'literal-grid', tall = scenario === 'tall-grid';
    const expectedRows = split ? [1, 1, 1, 2, 2, 2] : [2, 2, 2, 2, 2, 2];
    const rowPages = (mode: Awaited<ReturnType<typeof storyLayout>>) => mode.flatMap(paper => paper.rows.map(() => paper.number));
    const name = 'protected-furniture-' + scenario;
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await gridStoryFixture(scenario),
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => document.fonts.ready);
    const stack = page.locator('.document-stack');
    await expect.poll(async () => rowPages(await storyLayout(stack, false))).toEqual(expectedRows);
    const edit = await storyLayout(stack, false);
    const editGrids = await storyGridLayout(stack, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(3);
    const print = await storyLayout(preview, true);
    const printGrids = await storyGridLayout(preview, true);

    for (const mode of [edit, print]) {
      expect(mode).toHaveLength(3);
      expect(rowPages(mode)).toEqual(expectedRows);
      expect(mode.flatMap(paper => paper.rows.map(row => row.label))).toEqual(rowLabels);
      expect(mode.flatMap(paper => paper.tableHeads.map(head => [paper.number, head.label]))).toEqual(
        (split ? [1, 2] : [2]).flatMap(paper => headLabels.map(head => [paper, head])),
      );
      for (const paper of mode) {
        expect(paper.header.text).toEqual([gridLabels[0]]);
        expect(paper.footer.text).toEqual([gridLabels[1]]);
        expect(paper.header.painted && paper.footer.painted).toBe(true);
        expect(paper.body.top).toBeCloseTo(topMargin, 1);
        expect(Math.abs(paper.height - paper.body.bottom -
          Math.max(bottomMargin, paper.height - paper.footer.top + 2))).toBeLessThan(.5);
        if (tall) expect(paper.height - paper.body.bottom - bottomMargin).toBeGreaterThan(10);
        else expect(paper.height - paper.body.bottom).toBeCloseTo(bottomMargin, 1);
        for (const row of [...paper.rows, ...paper.tableHeads]) {
          expect(row.top).toBeGreaterThanOrEqual(paper.body.top - .5);
          expect(row.bottom).toBeLessThanOrEqual(paper.body.bottom + .5);
        }
      }
    }
    const nearBox = (edit: number[], print: number[]) => {
      expect(edit).toHaveLength(print.length);
      for (const [i, value] of edit.entries()) expect(Math.abs(value - print[i])).toBeLessThan(.5);
    };
    for (const [pageIndex, grids] of printGrids.entries()) {
      expect(grids.map(grid => grid.text)).toEqual(gridLabels);
      for (const [i, grid] of grids.entries()) {
        const isAuto = i === 1 && scenario === 'auto-grid';
        const original = editGrids[pageIndex][i];
        expect(grid).toMatchObject({ rows: 1, layout: isAuto ? 'autofit' : 'fixed', breaks: i === 1 && tall ? 2 : 1 });
        expect(original).toMatchObject({ rows: grid.rows, layout: grid.layout, breaks: grid.breaks, text: grid.text });
        expect(grid.cells).toHaveLength(3);
        expect(original.cells).toHaveLength(3);
        expect(original.cols).toHaveLength(grid.cols.length);
        if (isAuto) {
          expect(grid.cols).toHaveLength(0);
          expect(grid.box[2]).toBeLessThan(gridWidths[i].reduce((sum, width) => sum + width, 0));
        } else {
          expect(grid.authored).toEqual(gridWidths[i]);
          expect(original.authored).toEqual(gridWidths[i]);
          expect(grid.cols).toHaveLength(3);
          expect(grid.box[2]).toBeCloseTo(gridWidths[i].reduce((sum, width) => sum + width, 0), 1);
          for (const [n, width] of gridWidths[i].entries()) {
            expect(grid.cols[n][2]).toBeCloseTo(width, 1);
            expect(grid.cells[n][2]).toBeCloseTo(width, 1);
          }
        }
        nearBox(original.box, grid.box);
        nearBox(original.ink, grid.ink);
        for (const [n, cell] of grid.cells.entries()) nearBox(original.cells[n], cell);
        for (const [n, col] of grid.cols.entries()) nearBox(original.cols[n], col);
      }
      nearBox([edit[pageIndex].body.top, edit[pageIndex].body.bottom],
        [print[pageIndex].body.top, print[pageIndex].body.bottom]);
      for (const [i, row] of print[pageIndex].rows.entries()) {
        const other = edit[pageIndex].rows[i];
        nearBox([other.x, other.top, other.bottom], [row.x, row.top, row.bottom]);
      }
    }
  });
}

async function framedGridStoryFixture() {
  const zip = await JSZip.loadAsync(await gridStoryFixture('literal-grid'));
  const document = await zip.file('word/document.xml')!.async('string');
  const sections = [...document.matchAll(/<w:sectPr>[\s\S]*?<\/w:sectPr>/g)];
  if (sections.length !== 1) throw new Error('Expected the canonical single A4 section');
  const a4 = sections[0][0];
  const withBreak = (type: 'nextPage' | 'continuous') =>
    a4.replace('<w:pgSz', `<w:type w:val="${type}"/><w:pgSz`);
  const letter = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  const paragraph = (text: string, section = '') =>
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/>${section}</w:pPr>
      <w:r><w:t>${text}</w:t></w:r></w:p>`;
  // The paragraph properties end the current section. Both A4 sections
  // reference the original same default header and footer table parts.
  zip.file('word/document.xml', document.replace(a4, withBreak('continuous'))
    .replace('<w:body>', '<w:body>' + paragraph('QA LETTER FIRST', letter))
    .replace('</w:tbl>', '</w:tbl>' + paragraph('QA A4 BEFORE CONTINUOUS', withBreak('nextPage')) +
      paragraph('QA A4 AFTER CONTINUOUS')));
  return zip.generateAsync({ type: 'nodebuffer' });
}

function framedGridPlacement(root: Locator, print: boolean) {
  return root.evaluate(async (scope, { print, rowLabels }) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const sheets = [...scope.querySelectorAll<HTMLElement>(print ? '.print-page' : '.section-paper')];
    const labels = [...rowLabels, 'QA LETTER FIRST', 'QA A4 BEFORE CONTINUOUS', 'QA A4 AFTER CONTINUOUS', 'TAIL'];
    const selector = print ? '.print-content p' : '.section-editable p';
    const paragraphs = [...scope.querySelectorAll<HTMLParagraphElement>(selector)]
      .filter(p => labels.includes(p.textContent!));
    return sheets.map((sheet, i) => {
      const bounds = sheet.getBoundingClientRect(), scale = documentScale(sheet);
      const markers = paragraphs.flatMap(p => {
        const range = document.createRange(); range.selectNodeContents(p);
        const ink = range.getBoundingClientRect();
        const here = print ? p.closest('.print-page') === sheet : ink.top >= bounds.top - .5 && ink.top < bounds.bottom;
        return here ? [{ text: p.textContent!, page: i + 1, x: (ink.x - bounds.x) / scale, y: (ink.y - bounds.y) / scale }] : [];
      });
      return {
        height: bounds.height / scale, markers,
        header: sheet.querySelector(print ? '.print-header' : '.section-header')?.textContent?.trim(),
        footer: sheet.querySelector(print ? '.print-footer' : '.section-footer')?.textContent?.trim(),
      };
    });
  }, { print, rowLabels });
}

test('an internal continuous section fences an earlier protected table in the same A4 frame', async ({ page }) => {
  const name = 'protected-grid-framed-continuous';
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await framedGridStoryFixture(),
  });
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
  await expect(page.locator('.busy-indicator')).toBeHidden();
  await page.evaluate(() => document.fonts.ready);
  const stack = page.locator('.document-stack'), frames = stack.locator('.document-section-content');
  await expect(frames).toHaveCount(2);
  const letter = frames.nth(0), a4 = frames.nth(1);
  await expect(letter.locator('.section-editable')).toContainText('QA LETTER FIRST');
  await expect(a4.locator('.section-editable [data-section-page]')).toHaveCount(1);
  const settings = await a4.evaluate(frame => {
    const { breakType: firstType, ...first } = JSON.parse((frame as HTMLElement).dataset.sectionContent!);
    const marker = frame.querySelector<HTMLElement>('.section-editable [data-section-page]')!;
    const { breakType: nextType, ...next } = JSON.parse(marker.dataset.sectionPage!);
    const table = frame.querySelector('.section-editable table')!;
    return { firstType, nextType, first, next,
      tablePrecedesMarker: !!(table.compareDocumentPosition(marker) & Node.DOCUMENT_POSITION_FOLLOWING) };
  });
  expect(settings).toMatchObject({ firstType: 'nextPage', nextType: 'continuous', tablePrecedesMarker: true });
  expect(settings.next).toEqual(settings.first);
  const rowsOnPaper = (layout: Awaited<ReturnType<typeof framedGridPlacement>>) => layout.flatMap(sheet =>
    sheet.markers.filter(mark => rowLabels.includes(mark.text)).map(mark => [mark.text, mark.page]));
  const expectedRows = rowLabels.map(text => [text, 3]);
  await expect.poll(async () => rowsOnPaper(await framedGridPlacement(stack, false))).toEqual(expectedRows);
  const edited = await framedGridPlacement(stack, false);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const preview = page.locator('.print-layout');
  await expect(preview.locator('.print-page')).toHaveCount(4);
  const printed = await framedGridPlacement(preview, true);
  for (const mode of [edited, printed]) {
    expect(mode).toHaveLength(4);
    expect(mode[0].height).toBeCloseTo(1056, 1);
    expect(mode[0].markers.map(mark => mark.text)).toEqual(['QA LETTER FIRST']);
    expect(mode[0].header ?? '').toBe('');
    expect(mode[0].footer ?? '').toBe('');
    for (const sheet of mode.slice(1)) {
      expect(sheet.height).toBeCloseTo(1122.52, 1);
      expect(sheet.header).toBe('QA GRID HEADER');
      expect(sheet.footer).toBe('QA GRID FOOTER');
    }
    expect(rowsOnPaper(mode)).toEqual(expectedRows);
    expect(mode[2].markers.slice(-2).map(mark => mark.text)).toEqual(['QA A4 BEFORE CONTINUOUS', 'QA A4 AFTER CONTINUOUS']);
    expect(mode[3].markers.map(mark => mark.text)).toEqual(['TAIL']);
  }
  for (const [i, mark] of edited[2].markers.entries()) {
    expect(Math.abs(mark.x - printed[2].markers[i].x), mark.text + ' x').toBeLessThan(.75);
    expect(Math.abs(mark.y - printed[2].markers[i].y), mark.text + ' y').toBeLessThan(.75);
  }
});
