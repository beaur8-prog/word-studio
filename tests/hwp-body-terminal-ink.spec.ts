import { expect, test } from '@playwright/test';

test('a witnessed page-end line counts the language-shaped ink and rejects a changed flow', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { clearSourceBodyEndPaint, layoutHwpBodyLines, sourceBodyEndBottom } = await import('/src/hwp-body-line-metrics');
    const { encodeHwpBodyLine, parseHwpBodyLine } = await import('/src/hwp-body-line');
    const { measuredHwpLineInk } = await import('/src/hwp-line-ink');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances');

    const body = document.createElement('div');
    body.className = 'print-content tiptap';
    body.style.cssText = 'position:absolute;left:0;top:0;width:400px;height:51.2px;min-height:0;';
    body.lang = 'ko';
    function makeLine(text: string) {
      const p = document.createElement('p');
      p.style.cssText = 'font:700 20px/31px "Noto Serif KR";margin:0;height:31px;min-height:0;text-align:justify;';
      p.dataset.hwpBodyLine = '3 155 20 31';
      p.textContent = text;
      return p;
    }
    const prior = makeLine('가나다라마바사'), last = makeLine('(가나다) 문장 바닥');
    await document.fonts.load('700 20px "Noto Serif KR"', prior.textContent + last.textContent);
    const policy = { percent: 155, font: 20, base: 31, maxFont: 20, align: 'mixed-top' as const,
      bodyEndFit: { baseline: 17, remainder: 20.2, width: 400,
        fingerprint: hwpWrapFingerprint(last.textContent), preceding: hwpWrapFingerprint(prior.textContent) } };
    last.dataset.hwpBodyLine = encodeHwpBodyLine(policy);
    body.append(prior, last);
    document.body.append(body);
    try {
      layoutHwpBodyLines(body, true);
      const box = last.getBoundingClientRect(), actualBottom = body.getBoundingClientRect().bottom;
      const fullRoom = actualBottom - box.top;
      const korean = measuredHwpLineInk(last, 1, 20);
      last.lang = 'en';
      const english = measuredHwpLineInk(last, 1, 20);
      const englishPlacement = sourceBodyEndBottom(last, body);
      last.lang = 'ko';
      const placement = sourceBodyEndBottom(last, body);
      const projected = last.dataset.hwpProjectedBodyEnd === 'true';
      const unbrokenLine = last.textContent;
      clearSourceBodyEndPaint(last);
      body.style.width = '402px';
      const changedWidth = sourceBodyEndBottom(last, body);
      body.style.width = '400px';
      prior.querySelector('[data-hwp-body-run]')!.textContent += ' 바뀜';
      const changedPredecessor = sourceBodyEndBottom(last, body);
      return {
        fullRoom, originalHeight: box.height, localizedBottom: korean && korean.bottom - korean.baseline,
        unlocalizedBottom: english && english.bottom - english.baseline, projected,
        respectedBound: placement !== undefined && placement <= actualBottom,
        englishAccepted: englishPlacement !== undefined, widthAccepted: changedWidth !== undefined,
        changedPredecessorAccepted: changedPredecessor !== undefined, unbrokenLine,
        policyRoundTrip: JSON.stringify(parseHwpBodyLine(last.dataset.hwpBodyLine)?.bodyEndFit) === JSON.stringify(policy.bodyEndFit),
      };
    } finally { body.remove(); }
  });

  expect(result.fullRoom).toBeCloseTo(20.2, 1);
  expect(result.originalHeight).toBe(31);
  expect(result.localizedBottom).not.toBeNull();
  expect(result.unlocalizedBottom).toBeGreaterThan(result.localizedBottom! + 1);
  expect(result.projected).toBe(true);
  expect(result.respectedBound).toBe(true);
  expect(result.englishAccepted).toBe(false);
  expect(result.widthAccepted).toBe(false);
  expect(result.changedPredecessorAccepted).toBe(false);
  expect(result.unbrokenLine).toBe('(가나다) 문장 바닥');
  expect(result.policyRoundTrip).toBe(true);
});

test('the live source line has the same ink proof across an old page spacer', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpBodyLines } = await import('/src/hwp-body-line-metrics');
    const { editorSourceBodyEnd } = await import('/src/hwp-body-end-editor');
    const { encodeHwpBodyLine } = await import('/src/hwp-body-line');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances');
    const body = document.createElement('div');
    body.className = 'tiptap';
    body.style.cssText = 'position:absolute;left:0;top:0;width:400px;min-height:0;';
    body.lang = 'ko';

    function paragraph(text: string) {
      const p = document.createElement('p');
      p.style.cssText = 'font:700 20px/31px "Noto Serif KR";margin:0;height:31px;min-height:0;text-align:justify;';
      p.dataset.hwpBodyLine = '3 155 20 31';
      p.textContent = text;
      return p;
    }

    const former = paragraph('가나다라마바사'), last = paragraph('(가나다) 문장 바닥');
    await document.fonts.load('700 20px "Noto Serif KR"', former.textContent + last.textContent);
    last.dataset.hwpBodyLine = encodeHwpBodyLine({ percent: 155, font: 20, base: 31, maxFont: 20, align: 'mixed-top',
      bodyEndFit: { baseline: 17, remainder: 20.2, width: 400,
        fingerprint: hwpWrapFingerprint(last.textContent), preceding: hwpWrapFingerprint(former.textContent) } });
    body.append(former, last);
    document.body.append(body);
    try {
      layoutHwpBodyLines(body);
      const before = editorSourceBodyEnd(last, former, 1, 20.2);
      const spacer = document.createElement('div');
      spacer.className = 'ProseMirror-widget pagination-gap';
      spacer.style.height = '232px';
      former.after(spacer);
      const displaced = editorSourceBodyEnd(last, former, 1, 20.2);
      const tooLittleRoom = editorSourceBodyEnd(last, former, 1, 19.5);
      const tooMuchRoom = editorSourceBodyEnd(last, former, 1, 21.5);
      if (displaced) last.style.cssText += ';' + displaced.style;
      const active = editorSourceBodyEnd(last, former, 1, 20.2);
      spacer.classList.remove('pagination-gap');
      const authoredBreak = editorSourceBodyEnd(last, former, 1, 20.2);
      spacer.classList.add('pagination-gap');
      former.querySelector('[data-hwp-body-run]')!.textContent += ' 바뀜';
      const modified = editorSourceBodyEnd(last, former, 1, 20.2);
      const intact = last.textContent;
      return { before, displaced, active, tooLittleRoom, tooMuchRoom, authoredBreak, modified, intact };
    } finally { body.remove(); }
  });

  expect(result.before?.bottom).toBeLessThanOrEqual(20.2);
  expect(result.displaced).toEqual(result.before);
  expect(result.active).toEqual(result.before);
  expect(result.tooLittleRoom).toBeUndefined();
  expect(result.tooMuchRoom).toBeUndefined();
  expect(result.authoredBreak).toBeUndefined();
  expect(result.modified).toBeUndefined();
  expect(result.intact).toBe('(가나다) 문장 바닥');
});

test('equal native HWP pitches fit two unedited lines without moving their ink', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { layoutHwpBodyLines } = await import('/src/hwp-body-line-metrics');
    const { encodeHwpBodyLine, parseHwpBodyLine } = await import('/src/hwp-body-line');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments');
    const body = document.createElement('div');
    body.className = 'tiptap print-content';
    body.style.cssText = 'position:absolute;inset:0 auto auto 0;width:214px;min-height:0';
    const paragraph = document.createElement('p');
    paragraph.style.cssText = 'font:20px/32px monospace;margin:0;min-height:0;text-align:justify';
    const raised = document.createElement('span');
    raised.style.verticalAlign = '1px';
    raised.textContent = 'aaa';
    const tail = document.createElement('span');
    tail.textContent = 'fff';
    paragraph.append(raised, ' bbbb ccccc dddddd eeee ', tail);
    const literal = paragraph.textContent!;
    const policy = { percent: 160, font: 20, base: 32, storedUniformPair: true as const };
    paragraph.dataset.hwpBodyLine = encodeHwpBodyLine(policy);
    paragraph.dataset.hwpSourceWrap = JSON.stringify({
      fingerprint: hwpWrapFingerprint(literal), ends: [10, normalizeHwpWrapText(literal).length], bodyWidth: 214,
    });
    body.append(paragraph);
    document.body.append(body);
    const measure = () => {
      const range = document.createRange();
      range.selectNodeContents(paragraph);
      const bounds = paragraph.getBoundingClientRect();
      return { height: bounds.height, active: paragraph.dataset.hwpStoredMixedPitch === 'true',
        breaks: paragraphLineOffsets(paragraph), text: paragraph.textContent,
        inkBoxes: [...range.getClientRects()].filter(box => box.width && box.height)
          .map(box => [box.left - bounds.left, box.top - bounds.top, box.width, box.height]) };
    };
    try {
      layoutHwpBodyLines(body);
      const before = measure();
      layoutHwpBodyLines(body, true);
      const applied = measure();
      body.style.width = '216px';
      layoutHwpBodyLines(body, true);
      const changedWidth = measure();
      body.style.width = '214px';
      tail.style.position = 'relative';
      tail.style.top = '22px';
      layoutHwpBodyLines(body, true);
      const displaced = measure();
      tail.style.top = '';
      paragraph.firstElementChild!.querySelector('[data-hwp-body-run]')!.textContent += 'a';
      layoutHwpBodyLines(body, true);
      const edited = measure();
      const fallback = parseHwpBodyLine(encodeHwpBodyLine({ ...policy, percent: 210, base: 42, align: 'mixed-top' }));
      return { before, applied, changedWidth, displaced, edited, parsed: parseHwpBodyLine(paragraph.dataset.hwpBodyLine), fallback };
    } finally { body.remove(); }
  });
  expect(result.before.breaks).toHaveLength(1);
  expect(result.before.height).toBeGreaterThan(64);
  expect(result.applied).toEqual({ ...result.before, height: 64, active: true });
  expect(result.changedWidth.active).toBe(false);
  expect(result.displaced.active).toBe(false);
  expect(result.edited.active).toBe(false);
  expect(result.parsed?.storedUniformPair).toBe(true);
  expect(result.fallback).toEqual({ percent: 210, font: 20, base: 42, align: 'mixed-top' });
});

test('native page-end measurement preserves the original body paragraph', async ({ page }) => {
  await page.goto('/');
  const measured = await page.evaluate(async () => {
    const { clearSourceBodyEndPaint, sourceForcedBodyEndBottom } = await import('/src/hwp-body-line-metrics');
    const { measuredHwpLineInk } = await import('/src/hwp-line-ink');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments');
    const body = document.createElement('div');
    body.className = 'tiptap print-content';
    body.style.cssText = 'position:absolute;left:10px;top:14px;width:340px;min-height:0;';
    const previous = document.createElement('p'), terminal = document.createElement('p');
    previous.style.cssText = 'font:20px/32px "Studio Text";margin:0;min-height:0;';
    previous.textContent = 'aaaa bbbb cccc dddd eeee ffff gggg hhhh iiii';
    terminal.style.cssText = 'font:20px/32px "Studio Text";margin:12px 0 0 35px;min-height:0;text-indent:-35px;';
    terminal.textContent = 'Last complete line.';
    terminal.dataset.hwpPrintBeforeExplicitBreak = 'true';
    body.append(previous, terminal);
    await document.fonts.load('20px "Studio Text"', previous.textContent + terminal.textContent);
    document.body.append(body);
    try {
      const top = terminal.getBoundingClientRect().top;
      const precedingLines = paragraphLineOffsets(previous).length + 1;
      const ink = measuredHwpLineInk(terminal, 1, 20, true);
      if (!ink) throw Error('The fixture must provide measured visible ink');
      const relativeTop = top - body.getBoundingClientRect().top;
      const fittingRoom = 26;
      const overflowingRoom = 18;
      const proof = { fingerprint: hwpWrapFingerprint(terminal.textContent),
        ends: [normalizeHwpWrapText(terminal.textContent).length],
        forcedBodyEnd: { baseline: Math.min(19.5, ink.baseline - top - 4), room: fittingRoom,
          width: 340, spaceBefore: 12, pitch: 32, font: 20, hanging: 35,
          preceding: hwpWrapFingerprint(previous.textContent), precedingLines } };
      const setAvailableRoom = (room: number) => {
        body.style.height = (relativeTop + room) + 'px';
        terminal.dataset.hwpSourceWrap = JSON.stringify({ ...proof, forcedBodyEnd: { ...proof.forcedBodyEnd, room } });
      };
      setAvailableRoom(overflowingRoom);
      const overflowingFit = sourceForcedBodyEndBottom(terminal, body, false);
      setAvailableRoom(fittingRoom);
      const original = terminal.outerHTML;
      const nativeFit = sourceForcedBodyEndBottom(terminal, body, false);
      const unchanged = terminal.outerHTML === original && terminal.getBoundingClientRect().top === top;
      const printFit = sourceForcedBodyEndBottom(terminal, body);
      const projected = terminal.dataset.hwpProjectedBodyEnd === 'true' && terminal.getBoundingClientRect().top < top;
      clearSourceBodyEndPaint(terminal);
      previous.textContent += ' changed';
      const afterEdit = sourceForcedBodyEndBottom(terminal, body, false);
      return { precedingLines, nativeFit, unchanged, printFit, projected, overflowingFit, afterEdit,
        overflowBy: proof.forcedBodyEnd.baseline + ink.bottom - ink.baseline - overflowingRoom,
        bottom: body.getBoundingClientRect().bottom, originalBottom: top + 32,
        breaks: terminal.querySelectorAll('br,wbr').length, text: terminal.textContent };
    } finally { body.remove(); }
  });
  expect(measured.precedingLines).toBe(2);
  expect(measured.nativeFit).toBeDefined();
  expect(measured.nativeFit).toBeLessThanOrEqual(measured.bottom);
  expect(measured.originalBottom).toBeGreaterThan(measured.bottom);
  expect(measured.unchanged).toBe(true);
  expect(measured.printFit).toBe(measured.nativeFit);
  expect(measured.projected).toBe(true);
  expect(measured.overflowBy).toBeGreaterThan(0);
  expect(measured.overflowingFit).toBeUndefined();
  expect(measured.afterEdit).toBeUndefined();
  expect(measured.breaks).toBe(0);
  expect(measured.text).toBe('Last complete line.');
});
