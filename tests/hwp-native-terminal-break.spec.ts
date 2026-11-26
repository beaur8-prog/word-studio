import { expect, test } from '@playwright/test';

test('terminal export accepts unchanged authored wraps and leaves edited or resized wraps measured', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { measuredAuthoredTerminalWrap, measuredTerminalPageBreak } = await import('/src/hwp/export-terminal-page-break');
    const { hiddenEmptyLineClass } = await import('/src/page-empty-lines');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances');
    const frame = document.createElement('div');
    frame.style.cssText = 'position:absolute;left:10px;top:10px;width:160px;height:130px;';
    const empty = () => {
      const p = document.createElement('p');
      p.dataset.exactLineHeight = 'true';
      p.style.cssText = 'font:10px/14px monospace;margin:0;height:14px;';
      return p;
    };
    const terminal = document.createElement('p');
    terminal.style.cssText = 'font:20px/32px "Studio Text";margin:0 0 0 20px;text-indent:-20px;white-space:break-spaces;';
    const literal = 'First word second third';
    terminal.textContent = literal;
    await document.fonts.load('20px "Studio Text"', literal);
    terminal.dataset.hwpBodyLine = '3 160 20 32';
    const proof = {
      fingerprint: hwpWrapFingerprint(literal), ends: [5,9,20],
      authoredBodyLines: { origin: { page: 1, x: 10, y: 24 }, width: 160, hanging: 20, size: 20, pitch: 32 },
    };
    const trailing = empty();
    frame.append(empty(), terminal, trailing);
    document.body.append(frame);
    try {
      const measured = measuredTerminalPageBreak(frame, 0);
      terminal.dataset.hwpSourceWrap = JSON.stringify(proof);
      const authored = measuredTerminalPageBreak(frame, 0);
      terminal.dataset.hwpSourceWrap = JSON.stringify({ ...proof, fingerprint: '0000000000000000' });
      const edited = measuredTerminalPageBreak(frame, 0);
      terminal.dataset.hwpSourceWrap = JSON.stringify(proof);
      frame.style.width = '162px';
      const resized = measuredTerminalPageBreak(frame, 0);
      frame.style.width = '160px';
      trailing.style.height = '0px';
      const unclassified = measuredAuthoredTerminalWrap(frame, 0);
      trailing.classList.add(hiddenEmptyLineClass);
      const edge = measuredAuthoredTerminalWrap(frame, 0);
      terminal.dataset.hwpSourceWrap = JSON.stringify({ ...proof, fingerprint: '0000000000000000' });
      const edgeEdited = measuredAuthoredTerminalWrap(frame, 0);
      return { measured, authored, edited, resized, literal, unclassified, edge, edgeEdited };
    } finally { frame.remove(); }
  });

  expect(result.measured?.starts).toHaveLength(3);
  expect(result.measured?.starts[1]).toBeGreaterThan(6);
  expect(result.authored?.starts).toEqual([0,6,11]);
  expect(result.authored?.literal).toBe(result.literal);
  expect(result.authored?.origin).toBe(result.measured?.origin);
  expect(result.edited?.starts).toEqual(result.measured?.starts);
  expect(result.resized?.starts).toEqual(result.measured?.starts);
  expect(result.unclassified).toBeUndefined();
  expect(result.edge?.starts).toEqual([0,6,11]);
  expect(result.edge?.literal).toBe(result.literal);
  expect(result.edgeEdited).toBeUndefined();
});

test('a short printed table row keeps the size of its authored empty cell lines', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { measuredEmptyTableRows } = await import('/src/hwp/export-empty-table-rows');
    const body = document.createElement('div');
    body.style.cssText = 'position:absolute;left:10px;top:10px;width:160px;';
    body.innerHTML = '<p style="margin:0">An editable line before the table</p>' +
      '<table data-word-table-layout="fixed" style="width:160px;border-collapse:collapse">' +
      '<tr><td>Letter</td><td>Number</td></tr>' +
      '<tr data-row-split="avoid" style="height:7px"><td></td><td></td></tr>' +
      '<tr><td>A</td><td>1</td></tr></table>';
    const table = body.querySelector('table')!;
    const row = table.rows[1];
    for (const cell of row.cells) {
      cell.style.padding = '0';
      cell.innerHTML = '<p data-word-line-rule="exact" data-hwp-cell-authored-size="6.66667px" ' +
        'data-hwp-cell-line="1 Percent 6.6667 10 6.7 6.7" style="font:5pt/6.7px sans-serif;margin:0"></p>';
    }
    document.body.append(body);
    try {
      const top = body.getBoundingClientRect().top;
      const bottom = table.getBoundingClientRect().bottom;
      body.style.height = `${bottom-top+6}px`;
      const nearEdge = measuredEmptyTableRows(body);
      const empty = row.cells[0].firstElementChild as HTMLElement;
      empty.textContent = ' ';
      const retainedSpace = measuredEmptyTableRows(body);
      empty.textContent = '';
      empty.style.fontSize = '8pt';
      const modified = measuredEmptyTableRows(body);
      empty.style.fontSize = '5pt';
      body.style.height = `${bottom-top+50}px`;
      const notAtEdge = measuredEmptyTableRows(body);
      return { nearEdge, retainedSpace, modified, notAtEdge };
    } finally { body.remove(); }
  });

  expect(result.nearEdge).toHaveLength(1);
  expect(result.nearEdge[0].rows).toEqual([{ row: 1, font: 500, height: 7 }]);
  expect(result.nearEdge[0].cells).toEqual([['Letter','Number'],['',''],['A','1']]);
  expect(result.retainedSpace).toEqual([]);
  expect(result.modified).toEqual([]);
  expect(result.notAtEdge).toEqual([]);
});


test('an authored single-line proof requires stored line metrics and preserves the typography guard', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { nativeLineTypography, preserveHwpBodyLineEnds } = await import('/src/hwp/script-fonts');
    const { measuredAuthoredSingleLines } = await import('/src/hwp/export-terminal-page-break');
    const literal = '  one authored line';
    let bold = false;
    const native = {
      getCharShapeRuns: () => JSON.stringify([{ startOffset: 0, endOffset: 5 }, { startOffset: 5, endOffset: literal.length }]),
      getCharPropertiesAt: (_section: number, _paragraph: number, offset: number) =>
        JSON.stringify({ fontFamily: 'monospace', fontSize: 1500, bold: offset >= 5 && bold }),
    };
    const typography = nativeLineTypography(native, 0, 0, literal.length);
    bold = true;
    const modifiedTypography = nativeLineTypography(native, 0, 0, literal.length);
    const html = `<p style="font:20px/32px monospace;text-indent:-16px;text-align:justify;margin:0">${literal}</p>`;
    const line = { textStart: 0, verticalPosition: 64, width: 240, height: 20, textHeight: 20, spacing: 12, baseline: 17 };
    const witness = { origin: { page: 1, x: 10, y: 74 }, hanging: 16, nativeSingle: typography, painted: [{ y: 74, charStart: 0 }] };
    const source = (stored: boolean) => JSON.stringify({ stored, controlCount: 0, lines: [line] });
    const body = document.createElement('div');
    body.style.width = '240px';
    body.innerHTML = preserveHwpBodyLineEnds(html, literal, () => source(true), witness);
    document.body.append(body);
    try {
      const current = measuredAuthoredSingleLines(body, 0);
      const paragraph = body.querySelector('p')!;
      paragraph.style.fontSize = '21px';
      const resized = measuredAuthoredSingleLines(body, 0);
      paragraph.style.fontSize = '20px';
      paragraph.textContent = literal + ' edit';
      const edited = measuredAuthoredSingleLines(body, 0);
      const inferred = preserveHwpBodyLineEnds(html, literal, () => source(false), witness);
      return { current, resized, edited, inferred, html, typography, modifiedTypography };
    } finally { body.remove(); }
  });

  expect(result.current).toHaveLength(1);
  expect(result.current[0].typography).toBe(result.typography);
  expect(result.modifiedTypography).not.toBe(result.typography);
  expect(result.resized).toEqual([]);
  expect(result.edited).toEqual([]);
  expect(result.inferred).toBe(result.html);
});
