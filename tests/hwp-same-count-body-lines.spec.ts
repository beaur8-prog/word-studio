import { expect, test } from '@playwright/test';

test('authored HWP line ends keep the literal and following position without adding a hard break', async ({ page }) => {
  await page.goto('/');
  const measured = await page.evaluate(async () => {
    const { layoutHwpAuthoredBodyLines } = await import('/src/hwp-body-line-metrics.ts');
    const { hwpWrapFingerprint, normalizeHwpWrapText } = await import('/src/hwp-text-advances.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const literal = 'AAAA BBBB CCCC DDDD EEEE FFFF';
    const paper = document.createElement('div');
    paper.className = 'print-page'; paper.dataset.page = '1';
    paper.style.cssText = 'position:fixed;top:0;left:0;width:400px;height:300px';
    paper.innerHTML = '<div class="tiptap print-content" style="position:absolute;top:10px;left:10px;width:222px;height:280px;padding:0 1px;min-height:0;box-sizing:border-box"></div>';
    document.body.append(paper);
    try {
      const body = paper.firstElementChild!;
      return ['source', 'edited space', 'changed width', 'different sheet', 'ink outside'].map(mode => {
        (body as HTMLElement).style.width = mode === 'ink outside' ? '220px' : '222px';
        (body as HTMLElement).style.paddingInline = mode === 'ink outside' ? '0' : '1px';
        body.innerHTML = '<p style="font-family:monospace;font-size:16px;line-height:24px;min-height:0;margin:24px 0 0 30px;text-indent:-30px;text-align:justify"></p><p>FOLLOWING</p>';
        const p = body.querySelector('p')!;
        // Keep the negative case outside the body even with narrower system fonts.
        if (mode === 'ink outside') p.style.cssText += ';position:relative;left:-4px';
        const following = p.nextElementSibling!;
        p.textContent = mode === 'edited space' ? literal.replace('CCCC DDDD', 'CCCC  DDDD') : literal;
        const originalText = p.textContent;
        const box = p.getBoundingClientRect();
        const ordinary = paragraphLineOffsets(p).map(offset => normalizeHwpWrapText(originalText.slice(0, offset)).length);
        p.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(literal), ends: [12, 24],
          authoredBodyLines: { width: mode === 'changed width' ? 222 : 220, hanging: 30, size: 16, pitch: 24,
            origin: { page: mode === 'different sheet' ? 2 : 1, x: box.left - 30, y: box.top } } });
        const before = following.getBoundingClientRect().top;
        layoutHwpAuthoredBodyLines(paper);
        const final = paragraphLineOffsets(p).map(offset => normalizeHwpWrapText(p.textContent!.slice(0, offset)).length);
        return { mode, ordinary, final, sameText: p.textContent === originalText,
          followingMoves: following.getBoundingClientRect().top !== before, hardBreaks: p.querySelectorAll('br').length };
      });
    } finally { paper.remove(); }
  });
  const { ordinary, ...source } = measured[0];
  expect(ordinary).toHaveLength(1);
  expect(ordinary[0]).toBeGreaterThan(12);
  expect(source).toEqual({ mode: 'source', final: [12], sameText: true, followingMoves: false, hardBreaks: 0 });
  for (const rejected of measured.slice(1)) {
    expect(rejected.final).toEqual(rejected.ordinary);
    expect(rejected.sameText).toBe(true);
    expect(rejected.followingMoves).toBe(false);
    expect(rejected.hardBreaks).toBe(0);
  }
});
