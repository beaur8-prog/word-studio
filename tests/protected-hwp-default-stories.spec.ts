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
