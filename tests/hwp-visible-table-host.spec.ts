import { expect, test } from '@playwright/test';

test('a floating HWP table follows its confirmed native host line at the first row', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { sourceHostBeforeFloatingTable } = await import('/src/hwp/paragraph-layout.ts');
    const source = { marginTop: 40, marginHeader: 10, columns: [{}] };
    const style = { textWrap: 'TopAndBottom', vertRelTo: 'Para', vertAlign: 'Top', vertOffset: 1500 };
    const line = { baseline: 14, height: 17, verticalPosition: 50 };
    const metrics = { stored: true, controlCount: 1, lines: [line], tables: [{ inline: false }] };
    const boxes = [{ pageIndex: 2, y: 115 }, { pageIndex: 3, y: 50 }];
    const pitches = new Map<string, number>([['page:0:1', 2], ['top:0:1', 100], ['lines:0:1', 1]]);
    const check = (input = {}) => {
      const data = { props: style, metrics, boxes, source, pitches, ...input };
      const doc = {
        getTableProperties: () => JSON.stringify(data.props),
        getParagraphLineMetrics: () => JSON.stringify(data.metrics),
        getTableCellBboxes: () => JSON.stringify(data.boxes),
        getPageInfo: () => JSON.stringify(data.source),
      };
      return sourceHostBeforeFloatingTable(doc, 0, 1, 0, data.pitches);
    };
    return {
      confirmed: check(),
      inline: check({ props: { ...style, treatAsChar: true } }),
      caption: check({ props: { ...style, hasCaption: true } }),
      separatePage: check({ boxes: [{ pageIndex: 3, y: 50 }, { pageIndex: 4, y: 50 }] }),
      samePageOnly: check({ boxes: [boxes[0]] }),
      beforeTextBaseline: check({ boxes: [{ ...boxes[0], y: 111 }, boxes[1]] }),
      beyondTextLine: check({ boxes: [{ ...boxes[0], y: 120 }, boxes[1]] }),
      recalcOnly: check({ metrics: { ...metrics, stored: false } }),
      columns: check({ source: { ...source, columns: [{}, {}] } }),
      inconsistentCoordinates: check({ metrics: { ...metrics, lines: [{ ...line, verticalPosition: 36 }] } }),
    };
  });
  expect(result).toEqual({
    confirmed: true,
    inline: false,
    caption: false,
    separatePage: false,
    samePageOnly: false,
    beforeTextBaseline: false,
    beyondTextLine: false,
    recalcOnly: false,
    columns: false,
    inconsistentCoordinates: false,
  });
});

test('a floating HWP table can follow a complete host on the preceding sheet', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { sourceHostOnPriorPage } = await import('/src/hwp/paragraph-layout.ts');
    const source = { marginTop: 40, marginHeader: 10, marginBottom: 20, marginFooter: 10,
      height: 500, columns: [{ x: 75, width: 320 }] };
    const table = { textWrap: 'TopAndBottom', vertRelTo: 'Para', horzRelTo: 'Para', vertAlign: 'Top',
      vertOffset: 2700, outerTop: 150, outerBottom: 225 };
    const host = { stored: true, controlCount: 1, lines: [{ height: 20, verticalPosition: 50 }, { height: 20, verticalPosition: 80 }],
      tables: [{ inline: false }] };
    const empty = { stored: true, controlCount: 0, lines: [{ verticalPosition: 205, height: 20, spacing: 7 }] };
    const next = { stored: true, controlCount: 0, lines: [{ verticalPosition: 232 }] };
    const box = { pageIndex: 3, x: 90, y: 52, w: 250, h: 200 };
    const pitches = new Map<string, number>([['page:0:1', 2], ['top:0:1', 100], ['lines:0:1', 2], ['page:0:3', 3]]);
    const check = (options: { first?: typeof box; blank?: string; blankOffset?: number; changedHost?: typeof host } = {}) => {
      const doc = {
        getTableProperties: () => JSON.stringify(table),
        getParagraphLineMetrics: (_: number, paragraph: number) => JSON.stringify(paragraph === 1 ? options.changedHost ?? host
          : paragraph === 2 ? { ...empty, lines: [{ ...empty.lines[0], verticalPosition: options.blankOffset ?? 205 }] } : next),
        getTableCellBboxes: () => JSON.stringify([options.first ?? box]),
        getPageInfo: () => JSON.stringify(source), getParagraphCount: () => 4,
        getTextRange: () => options.blank ?? '', getParaPropertiesAt: () => JSON.stringify({ spacingBefore: 20, spacingAfter: 0 }),
      };
      return sourceHostOnPriorPage(doc, 0, 1, 0, pitches);
    };
    return { confirmed: check(), movedTable: check({ first: { ...box, y: 80 } }),
      samePage: check({ first: { ...box, pageIndex: 2 } }), nonEmpty: check({ blank: 'keep this text' }),
      blankGap: check({ blankOffset: 215 }), clippedHost: check({ changedHost: { ...host, lines: [host.lines[0], { height: 20, verticalPosition: 450 }] } }) };
  });
  expect(result.confirmed).toEqual({ x: 15, top: 2, bottom: 3, blankUsesTableGap: true });
  expect(result.nonEmpty?.blankUsesTableGap).toBe(false);
  expect(result.blankGap?.blankUsesTableGap).toBe(false);
  expect(result.movedTable).toBeUndefined();
  expect(result.samePage).toBeUndefined();
  expect(result.clippedHost).toBeUndefined();
});
