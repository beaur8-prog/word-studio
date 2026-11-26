import { test, expect } from '@playwright/test';

test('an observed inline table moves intact when it fits the next body and splits when it grows beyond it', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const settings = { ...defaultPage, size: 'Custom', width: 150, height: 100, margin: 10 };
    const stages = [];
    for (const [rows, witnessPage] of [[4,2],[12,2],[4,3]]) {
      const table = '<table data-hwp-kept-table-page="' + witnessPage + '" style="width:300px">' +
        Array.from({length:rows}, (_, index)=>'<tr style="height:45px"><td><p>ROW ' + index + '</p></td></tr>').join('') + '</table>';
      const load = await importFile(new File(['<p style="line-height:230px;margin:0">LEAD</p>' + table], 'inline.html'));
      let doc = { ...load.document, page: settings, content: parseDocumentHtml(load.html!) };
      const pages = [];
      for (let stage=0;stage<3;stage++) {
        const layout = await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
        try {
          pages.push([...layout.element.querySelectorAll('.print-content')].map(body=>({
            lead: body.textContent?.includes('LEAD'),
            rows: [...body.querySelectorAll('tr')].map(row=>row.textContent),
            inline: [...body.querySelectorAll('table')].every(t=>t.getAttribute('data-hwp-kept-table-page')===String(witnessPage)),
          })));
        } finally { layout.dispose(); }
        if (stage===2) break;
        const reopened = await importFile(new File([await exportDocx(doc)],'inline.docx'));
        doc = { ...reopened.document, page: reopened.document.page, content: parseDocumentHtml(reopened.html!) };
      }
      stages.push(pages);
    }
    return stages;
  });
  for (const pages of result[0]) {
    expect(pages.map(page=>page.rows.length)).toEqual([0,4]);
    expect(pages[0].lead).toBeTruthy();
    expect(pages[1].inline).toBeTruthy();
  }
  for (const pages of result[2]) {
    expect(pages[0].rows.length).toBeGreaterThan(0);
    expect(pages.flatMap(page=>page.rows)).toEqual(['ROW 0','ROW 1','ROW 2','ROW 3']);
  }
  for (const pages of result[1]) {
    expect(pages.length).toBeGreaterThan(1);
    expect(pages[0].rows.length).toBeGreaterThan(0);
    expect(pages).not.toContainEqual({ lead:false, rows:[], inline:true });
    expect(pages.flatMap(page=>page.rows)).toEqual(Array.from({length:12}, (_,index)=>'ROW '+index));
  }
});

test('one authored empty line can witness a whole-table boundary without taking its place', async ({ page }) => {
  await page.goto('/');
  const decisions = await page.evaluate(async () => {
    const { sourceKeptInlineTable } = await import('/src/hwp/paragraph-layout.ts');
    const common = { controlCount: 1, stored: true };
    const props = { treatAsChar: true, pageBreak: 1, hasCaption: false, textWrap: 'TopAndBottom',
      horzRelTo: 'Column', vertRelTo: 'Para', vertAlign: 'Top', outerTop: 150, outerBottom: 150 };
    const blankY = 856;
    const trial = (blank: { text?: string; top?: number; breakBefore?: boolean; pageLeading?: boolean }, laterHeight = 250) => {
      const pageLeading = Boolean(blank.pageLeading);
      const source = {
        getTextRange: (_: number, p: number) => p === 1 ? blank.text || '' : '',
        getParagraphLength: (_: number, p: number) => p === 1 ? blank.text?.length || 0 : 0,
        getParaPropertiesAt: () => JSON.stringify({ lineSpacingType: 'Percent', pageBreakBefore: !!blank.breakBefore }),
        getParagraphLineMetrics: (_: number, p: number) => JSON.stringify(p === 1
          ? { stored: true, controlCount: 0, lines: [{ verticalPosition: pageLeading ? 716 : 816, height: 16, spacing: 8 }] }
          : { ...common, lines: [p === 0 ? { verticalPosition: pageLeading ? 0 : 100, height: 700, spacing: 16 }
            : { verticalPosition: 0, height: laterHeight + 4, spacing: 8 }] }),
        getTableProperties: () => JSON.stringify(props),
        getTableDimensions: () => JSON.stringify({ cellCount: 1 }),
        getTableCellBboxes: (_: number, p: number) => JSON.stringify([
          { pageIndex: p === 0 ? 0 : 1, cellIdx: 0, x: 40, y: p === 0 && !pageLeading ? 142 : 42, w: 480, h: p === 0 ? 700 : laterHeight },
        ]),
        getPageInfo: () => JSON.stringify({ columns: [{}], marginTop: 32, marginHeader: 8, footerArea: { y: 960 } }),
      } as unknown as Parameters<typeof sourceKeptInlineTable>[0];
      const nativeInk = new Map<string, number>([['page:0:1', 0], ['top:0:1', blank.top ?? blankY - (pageLeading ? 100 : 0)]]);
      return sourceKeptInlineTable(source, 0, 2, 0, 2, nativeInk);
    };
    return [trial({}), trial({ text: 'edited' }), trial({ top: 854 }),
      trial({ breakBefore: true }), trial({}, 60), trial({ pageLeading: true })];
  });
  expect(decisions).toEqual([2, undefined, undefined, undefined, undefined, undefined]);
});

test('an automatic HWP row can take its height from intact authored blank lines', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { sourceBlankAutoCellFits } = await import('/src/hwp/paragraph-layout.ts');
    const verify = (change: {
      text?: string; controlCount?: number; stored?: boolean; top?: number; width?: number; innerHeight?: number;
    }) => {
      const source = {
        getCellParagraphCount: () => 3,
        getTextInCellByPath: (_section: number, _paragraph: number, path: string) =>
          JSON.parse(path)[0].cellParaIndex === 1 ? change.text ?? '' : '',
        getCellParaPropertiesAt: () => JSON.stringify({
          lineSpacingType: 'Percent', spacingBefore: 0, spacingAfter: 0, headType: 'None',
        }),
        getParagraphLineMetrics: (_section: number, _paragraph: number, path: string) => {
          const index = JSON.parse(path)[0].cellParaIndex;
          const height = [10, 12, 8][index];
          return JSON.stringify({
            stored: change.stored ?? true, controlCount: index === 1 ? change.controlCount ?? 0 : 0,
            lines: [{ height, textHeight: height, spacing: 2, textStart: 0,
              width: change.width ?? 150, verticalPosition: [0, change.top ?? 12, 26][index] }],
          });
        },
      } as unknown as Parameters<typeof sourceBlankAutoCellFits>[0];
      return sourceBlankAutoCellFits(source, [0, 0, 0, 0], 150, change.innerHeight ?? 34);
    };
    return [verify({}), verify({ text: ' ' }), verify({ controlCount: 1 }), verify({ stored: false }),
      verify({ top: 11 }), verify({ width: 149 }), verify({ innerHeight: 36 })];
  });
  expect(result).toEqual([true, false, false, false, false, false, false]);
});

test('a terminal source blank witnesses a whole next-page table after authored body text', async ({ page }) => {
  await page.goto('/');
  const decisions = await page.evaluate(async () => {
    const { sourceKeptInlineTableAfterBody } = await import('/src/hwp/paragraph-layout.ts');
    const trial = (change: { blankText?: boolean; blankOffset?: number; rows?: number; bottom?: number; bodyPage?: number }) => {
      const doc = {
        getParagraphLength: (_: number, p: number) => p === 0 ? 12 : p === 1 && change.blankText ? 1 : 0,
        getTextRange: (_: number, p: number) => p === 0 ? 'Source text' : '',
        getPageOfPosition: (_: number, p: number) => JSON.stringify({ ok: true, page: p === 0 ? change.bodyPage ?? 0 : 0 }),
        getParaPropertiesAt: (_: number, p: number) => JSON.stringify({ lineSpacingType: 'Percent', spacingAfter: p === 0 ? 8 : 0 }),
        getPageInfo: () => JSON.stringify({ columns: [{}], marginTop: 32, marginHeader: 8, footerArea: { y: change.bottom ?? 930 } }),
        getParagraphLineMetrics: (_: number, p: number) => JSON.stringify({
          stored: true, controlCount: p === 2 ? 1 : 0,
          lines: [{
            textStart: 0, width: 480, verticalPosition: [830, change.blankOffset ?? 858, 0][p],
            height: [18, 16, 34][p], spacing: [6, 8, 5][p],
          }],
        }),
        getTableProperties: () => JSON.stringify({ treatAsChar: true, pageBreak: 2, textWrap: 'TopAndBottom',
          horzRelTo: 'Para', vertRelTo: 'Para', vertAlign: 'Top', outerTop: 0, outerBottom: 0 }),
        getTableDimensions: () => JSON.stringify({ rowCount: change.rows ?? 1, cellCount: 2 }),
        getTableCellBboxes: () => JSON.stringify([0, 1].map(cellIdx => ({ pageIndex: 1, cellIdx, x: 40 + 180 * cellIdx, y: 40, w: 180, h: 34 }))),
      } as unknown as Parameters<typeof sourceKeptInlineTableAfterBody>[0];
      return sourceKeptInlineTableAfterBody(doc, 0, 2, 0);
    };
    return [trial({}), trial({ blankText: true }), trial({ blankOffset: 856 }), trial({ rows: 2 }),
      trial({ bottom: 960 }), trial({ bodyPage: 1 })];
  });
  expect(decisions).toEqual([2, undefined, undefined, undefined, undefined, undefined]);
});

test('restore HWP table-boundary lines only when the next source text also starts too early', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { defaultPage } = await import('/src/types.ts');
    const settings = { ...defaultPage, size: 'Custom', width: 100, height: 100, margin: 10 };
    const literal = 'AAAA BBBB CCCC DDDD', following = 'NEXT BODY WORDS';
    const trials = [];
    for (const nextBlank of [3, 22]) {
      const template = document.createElement('div');
      template.innerHTML = '<p style="margin:0;font:12px/176px monospace">Lead</p>' +
        '<p data-qa-preceding="true" style="position:relative;left:4px;margin:0 0 0 30px;' +
        'text-indent:-30px;font:16px/24px monospace;text-align:justify;min-height:0">' + literal + '</p>' +
        '<p style="margin:0;line-height:10px;min-height:0"></p>' +
        '<table data-hwp-kept-table-page="2" style="width:220px;margin:0">' +
        '<tbody><tr style="height:68px"><td><p>TABLE</p></td></tr></tbody></table>' +
        '<p style="margin:0;min-height:0;line-height:' + nextBlank + 'px"></p>' +
        '<p data-qa-following="true" style="margin:0 0 0 10px;text-indent:-10px;font:12px/12px monospace">' + following + '</p>';
      const preceding = template.querySelector<HTMLElement>('[data-qa-preceding]')!;
      const after = template.querySelector<HTMLElement>('[data-qa-following]')!;
      const measure = (root: HTMLElement) => {
        const sheet = (node: Element | null) => Number(node?.closest<HTMLElement>('.print-page')?.dataset.page);
        return [sheet(root.querySelector('table')), sheet(root.querySelector('[data-qa-following]'))];
      };
      const baseline = await createPrintLayout(template.innerHTML, settings);
      let original: number[];
      try {
        original = measure(baseline.element);
        const p = baseline.element.querySelector<HTMLElement>('[data-qa-preceding]')!;
        const body = p.closest<HTMLElement>('.print-content')!.getBoundingClientRect();
        const paper = p.closest<HTMLElement>('.print-page')!.getBoundingClientRect();
        const box = p.getBoundingClientRect();
        preceding.dataset.hwpSourceWrap = JSON.stringify({
          fingerprint: hwpWrapFingerprint(literal), ends: [8, 16],
          authoredBodyLines: { origin: { page: 1, x: box.left - paper.left - 30, y: box.top - paper.top },
            width: body.width, hanging: 30, size: 16, pitch: 24 },
        });
        after.dataset.hwpSourceWrap = JSON.stringify({
          fingerprint: hwpWrapFingerprint(following), ends: [8, 13],
          authoredBodyLines: { origin: { page: 2, x: body.left - paper.left, y: 0 },
            width: body.width, hanging: 10, size: 12, pitch: 12 },
        });
      } finally { baseline.dispose(); }
      const actual = await createPrintLayout(template.innerHTML, settings);
      try {
        trials.push({ original, projected: measure(actual.element),
          text: actual.element.querySelector('[data-qa-preceding]')!.textContent,
          restored: actual.element.querySelector('[data-qa-preceding] br') !== null });
      } finally { actual.dispose(); }
    }
    return trials;
  });
  expect(result).toEqual([
    { original: [1, 1], projected: [2, 2], text: 'AAAA BBBB CCCC DDDD', restored: true },
    { original: [1, 2], projected: [1, 2], text: 'AAAA BBBB CCCC DDDD', restored: false },
  ]);
});
