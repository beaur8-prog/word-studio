import { expect, test } from '@playwright/test';

test('a stored two-line HWP tail can keep its last line inside the body without trimming its strut', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances');
    const { sourceAuthoredTailInkBottom } = await import('/src/hwp-body-line-metrics');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments');
    const paper = document.createElement('div');
    paper.className = 'print-page';
    paper.dataset.page = '7';
    paper.style.cssText = 'position:absolute;left:10px;top:10px;width:210px;height:128px;';
    const body = document.createElement('div');
    body.style.cssText = 'position:absolute;left:0;top:0;width:190px;';
    const paragraph = document.createElement('p');
    paragraph.style.cssText = 'font:16px/26px monospace;margin:45px 10px 0 12px;text-indent:-12px;text-align:justify;';
    paragraph.textContent = 'First line fitting and second';
    paper.append(body);
    body.append(paragraph);
    document.body.append(paper);
    try {
      const literal = paragraph.textContent;
      const first = paragraph.getBoundingClientRect(), frame = paper.getBoundingClientRect();
      const starts = paragraphLineOffsets(paragraph);
      const end = literal.replace(/\s/g, '').length;
      paragraph.dataset.hwpSourceWrap = JSON.stringify({
        fingerprint: hwpWrapFingerprint(literal), ends: [literal.slice(0, starts[0]).replace(/\s/g, '').length, end],
        authoredTailInk: { origin: { page: 7, x: 0, y: first.top - frame.top }, width: first.width + 12,
          hanging: 12, rightMargin: 10, size: 16, pitch: 26 },
      });
      const fit = (bottom: number) => {
        body.style.height = (bottom - body.getBoundingClientRect().top) + 'px';
        return sourceAuthoredTailInkBottom(paragraph, body);
      };
      const valid = fit(first.bottom - .85);
      const clipped = fit(first.bottom - 9);
      paper.dataset.page = '8';
      const moved = fit(first.bottom - .85);
      paper.dataset.page = '7';
      paragraph.textContent += '!';
      const edited = fit(first.bottom - .85);
      return { lines: starts.length + 1, valid, clipped, moved, edited, originalHeight: first.height };
    } finally { paper.remove(); }
  });
  expect(result.lines).toBe(2);
  expect(result.originalHeight).toBe(52);
  expect(result.valid).toBeGreaterThan(0);
  expect(result.clipped).toBeUndefined();
  expect(result.moved).toBeUndefined();
  expect(result.edited).toBeUndefined();
});
