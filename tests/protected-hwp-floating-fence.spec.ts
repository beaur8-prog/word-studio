import JSZip from 'jszip';
import { expect, test, type Locator } from '@playwright/test';
import { fenceFixture, floatLabel, headLabels, localLabel, remoteLabel, rowLabels } from './protected-hwp-fence-fixture';

function fenceSnapshot(root: Locator, print: boolean, extraProbes: string[] = []) {
  return root.evaluate(async (scope, { print, rowLabels, headLabels, localLabel, remoteLabel, extraProbes }) => {
    const { documentScale } = await import('/src/document-scale.ts');
    const editor = scope.querySelector<HTMLElement>('.tiptap[contenteditable=true]');
    const stride = Number((scope as HTMLElement).dataset.pageHeight) + 24;
    const position = (element: Element, box: DOMRect = element.getBoundingClientRect()) => {
      const sheet = element.closest<HTMLElement>('.print-page');
      const reference = sheet ?? editor!, origin = reference.getBoundingClientRect();
      const scale = documentScale(reference);
      const localY = (box.top - origin.top) / scale;
      const page = sheet ? Number(sheet.dataset.page) : Math.max(1, 1 + Math.floor((localY + .5) / stride));
      return { page, x: (box.left - origin.left) / scale, y: sheet ? localY : localY - (page - 1) * stride,
        width: box.width / scale, height: box.height / scale };
    };
    const textPosition = (p: HTMLParagraphElement) => {
      const range = document.createRange(); range.selectNodeContents(p);
      return { text: p.textContent!, ...position(p, range.getBoundingClientRect()) };
    };
    const paragraphs = [...scope.querySelectorAll<HTMLParagraphElement>('p')];
    const image = [...scope.querySelectorAll<HTMLImageElement>('img[data-image-floating]')].map(img => {
      const anchor = img.closest('p')!;
      const box = position(img);
      const paper = img.closest('.print-page')?.getBoundingClientRect();
      const ink = img.getBoundingClientRect();
      return { label: img.alt, anchor: anchor.textContent!, anchorPage: position(anchor).page, box,
        rendered: img.complete && img.naturalWidth > 0 && box.width > 0 && box.height > 0 &&
          (!paper || ink.left >= paper.left - .5 && ink.right <= paper.right + .5 &&
            ink.top >= paper.top - .5 && ink.bottom <= paper.bottom + .5) };
    });
    return {
      pages: print ? scope.querySelectorAll('.print-page').length : Number((scope as HTMLElement).dataset.pages),
      rows: paragraphs.filter(p => rowLabels.includes(p.textContent!)).map(textPosition),
      headers: paragraphs.filter(p => headLabels.includes(p.textContent!)).map(textPosition),
      probes: paragraphs.filter(p => [localLabel, remoteLabel, 'AFTER-FLOAT', 'TAIL', ...extraProbes]
        .some(label => p.textContent?.endsWith(label))).map(textPosition),
      tables: print ? [...scope.querySelectorAll('.print-content table')].map(table => position(table)) : [],
      image,
    };
  }, { print, rowLabels, headLabels, localLabel, remoteLabel, extraProbes });
}

test('a later square drawing keeps its authored paper without blocking protected table flow', async ({ page }) => {
  type Measurement = Awaited<ReturnType<typeof fenceSnapshot>>;
  let unobstructed: Measurement | undefined;
  for (const scenario of ['absent', 'remote', 'local', 'outside-body'] as const) await test.step(scenario, async () => {
    const maySplit = scenario === 'absent' || scenario === 'remote';
    const rowPapers = maySplit ? [1, 1, 1, 2, 2, 2] : [2, 2, 2, 2, 2, 2];
    const headPapers = maySplit ? [1, 2] : [2];
    const name = `protected-rows-float-${scenario}`;
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await fenceFixture(scenario),
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => document.fonts.ready);
    const stack = page.locator('.document-stack');
    await expect(stack.locator('.tiptap[contenteditable=true] [data-page-break]')).toHaveCount(1);
    await expect.poll(async () => (await fenceSnapshot(stack, false)).rows.map(r => r.page)).toEqual(rowPapers);
    if (scenario !== 'absent') await expect.poll(async () => (await fenceSnapshot(stack, false)).image[0]?.rendered).toBe(true);
    const edited = await fenceSnapshot(stack, false);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(3);
    const printed = await fenceSnapshot(preview, true);

    for (const view of [edited, printed]) {
      expect(view.pages).toBe(3);
      expect(view.rows.map(r => [r.text, r.page])).toEqual(rowLabels.map((text, i) => [text, rowPapers[i]]));
      expect(view.headers.map(h => [h.text, h.page])).toEqual(headPapers.flatMap(p => headLabels.map(text => [text, p])));
      expect(view.probes.map(({ text, page }) => [text.endsWith(localLabel) ? localLabel : text, page]))
        .toEqual([[localLabel, 1], [remoteLabel, 3], ['AFTER-FLOAT', 3], ['TAIL', 3]]);
      expect(view.image).toHaveLength(scenario === 'absent' ? 0 : 1);
      if (scenario !== 'absent') {
        const drawingPaper = scenario === 'local' ? 1 : 3;
        const drawing = view.image[0];
        expect(drawing.label).toBe(floatLabel);
        expect(drawing.anchor).toContain(scenario === 'local' ? localLabel : remoteLabel);
        expect(drawing.anchorPage).toBe(drawingPaper);
        expect(drawing.box.page).toBe(drawingPaper);
        expect(drawing.rendered).toBe(true);
        for (const [key, expected] of Object.entries({ x: 88, y: scenario === 'outside-body' ? 974.92 : 940, width: 68, height: 72 }) as
          [keyof typeof drawing.box, number][])
          expect(Math.abs(drawing.box[key] - expected)).toBeLessThan(.5);
      }
    }
    expect(printed.tables.map(table => table.page)).toEqual(headPapers);
    for (const [i, row] of edited.rows.entries()) {
      expect(Math.abs(row.x - printed.rows[i].x), row.text + ' x').toBeLessThan(.5);
      expect(Math.abs(row.y - printed.rows[i].y), row.text + ' y').toBeLessThan(.5);
    }
    if (scenario === 'absent') unobstructed = printed;
    else if (scenario === 'remote') {
      for (const [i, row] of printed.rows.entries()) {
        expect(Math.abs(row.x - unobstructed!.rows[i].x), row.text + ' remote x').toBeLessThan(.5);
        expect(Math.abs(row.y - unobstructed!.rows[i].y), row.text + ' remote y').toBeLessThan(.5);
      }
    } else if (scenario === 'local') {
      // Compare against the actual retained table on the same paper with
      // no drawing. The fallback has already moved that table to paper 2.
      const table = unobstructed!.tables[0], drawing = printed.image[0].box;
      expect(table.page).toBe(drawing.page);
      expect(Math.min(table.x + table.width, drawing.x + drawing.width) - Math.max(table.x, drawing.x)).toBeGreaterThan(1);
      expect(Math.min(table.y + table.height, drawing.y + drawing.height) - Math.max(table.y, drawing.y)).toBeGreaterThan(1);
    }
  });
});


type InterveningCase = 'literal' | 'two-paragraphs' | 'keep-next' | 'no-break';
const intervening = 'QA BODY AFTER THE TABLE';
const secondIntervening = 'QA SECOND BODY BEFORE THE BREAK';

async function interveningFenceFixture(scenario: InterveningCase) {
  const zip = await JSZip.loadAsync(await fenceFixture('remote'));
  const document = await zip.file('word/document.xml')!.async('string');
  const breakParagraph = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  if ((document.match(/<\/w:tbl>/g) ?? []).length !== 1 || !document.includes(breakParagraph))
    throw new Error('Expected one protected table and the later drawing fence');
  const paragraph = (text: string, keep = false) => `<w:p><w:pPr><w:keepNext w:val="${keep ? 'on' : 'off'}"/>
    <w:keepLines w:val="off"/><w:widowControl w:val="off"/><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/>
    </w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const between = paragraph(intervening, scenario === 'keep-next') +
    (scenario === 'two-paragraphs' ? paragraph(secondIntervening) : '');
  const withBody = document.replace('</w:tbl>', '</w:tbl>' + between);
  zip.file('word/document.xml', scenario === 'no-break' ? withBody.replace(breakParagraph, '') : withBody);
  return zip.generateAsync({ type: 'nodebuffer' });
}

for (const scenario of ['literal', 'two-paragraphs', 'keep-next', 'no-break'] as const) {
  test(`late square drawing with ${scenario} intervening body preserves protected table placement`, async ({ page }) => {
    const split = scenario === 'literal', hasBreak = scenario !== 'no-break';
    const rowsOn = split ? [1, 1, 1, 2, 2, 2] : [2, 2, 2, 2, 2, 2];
    const headsOn = split ? [1, 2] : [2];
    const drawingPaper = hasBreak ? 3 : 2;
    const blockLabels = scenario === 'two-paragraphs' ? [intervening, secondIntervening] : [intervening];
    const name = 'protected-fence-intervening-' + scenario;
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await interveningFenceFixture(scenario),
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
    await expect(page.locator('.busy-indicator')).toBeHidden();
    await page.evaluate(() => document.fonts.ready);
    const stack = page.locator('.document-stack');
    const editor = stack.locator('.tiptap[contenteditable=true]');
    await expect(editor.locator('[data-page-break]')).toHaveCount(hasBreak ? 1 : 0);
    if (scenario === 'keep-next')
      await expect(editor.locator('p').filter({ hasText: intervening })).toHaveAttribute('data-word-keep-next', 'true');
    await expect.poll(async () => (await fenceSnapshot(stack, false, blockLabels)).rows.map(row => row.page)).toEqual(rowsOn);
    await expect.poll(async () => (await fenceSnapshot(stack, false, blockLabels)).image[0]?.rendered).toBe(true);
    const edited = await fenceSnapshot(stack, false, blockLabels);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const preview = page.locator('.print-layout');
    await expect(preview.locator('.print-page')).toHaveCount(drawingPaper);
    const printed = await fenceSnapshot(preview, true, blockLabels);
    for (const mode of [edited, printed]) {
      expect(mode.pages).toBe(drawingPaper);
      expect(mode.rows.map(row => [row.text, row.page])).toEqual(rowLabels.map((label, i) => [label, rowsOn[i]]));
      expect(mode.headers.map(head => [head.text, head.page])).toEqual(headsOn.flatMap(paper => headLabels.map(label => [label, paper])));
      expect(mode.probes.filter(probe => blockLabels.includes(probe.text)).map(probe => [probe.text, probe.page]))
        .toEqual(blockLabels.map(label => [label, 2]));
      expect(mode.probes.filter(probe => [remoteLabel, 'AFTER-FLOAT', 'TAIL'].includes(probe.text)).map(probe => [probe.text, probe.page]))
        .toEqual([remoteLabel, 'AFTER-FLOAT', 'TAIL'].map(label => [label, drawingPaper]));
      expect(mode.image).toHaveLength(1);
      const [figure] = mode.image;
      expect(figure).toMatchObject({ label: floatLabel, anchor: remoteLabel, anchorPage: drawingPaper, rendered: true });
      expect(figure.box.page).toBe(drawingPaper);
      for (const [key, value] of [['x', 88], ['y', 940], ['width', 68], ['height', 72]] as const)
        expect(Math.abs(figure.box[key] - value), key).toBeLessThan(.5);
    }
    expect(printed.tables.map(table => table.page)).toEqual(headsOn);
    for (const [i, row] of printed.rows.entries()) {
      expect(Math.abs(row.x - edited.rows[i].x), row.text + ' x').toBeLessThan(.5);
      expect(Math.abs(row.y - edited.rows[i].y), row.text + ' y').toBeLessThan(.5);
    }
  });
}
