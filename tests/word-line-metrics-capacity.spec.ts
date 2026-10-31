import { test, expect } from '@playwright/test';

// Synthetic capacity coverage, separate from the fixed Word-authored corpus.
test('unchanged long documents reuse measured lines beyond the shared cache capacity', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { wordLineHeight } = await import('/src/word-line-metrics.ts');
    const host = document.createElement('section');
    host.style.cssText = 'position:absolute;left:-3000px;width:700px';
    document.body.append(host);
    const paragraphs = Array.from({ length: 3000 }, (_, i) => {
      const p = document.createElement('p');
      p.style.cssText = 'font:400 11pt Verdana;margin:0';
      p.dataset.wordLineRule = 'auto'; p.dataset.wordLine = '1.15';
      p.textContent = `Unique paragraph ${i}: complete English text.`;
      host.append(p); return p;
    });
    const original = Element.prototype.getBoundingClientRect;
    let layoutReads = 0;
    Element.prototype.getBoundingClientRect = function () {
      layoutReads++; return original.call(this);
    };
    try {
      const cold = paragraphs.map(p => wordLineHeight(p));
      const coldReads = layoutReads; layoutReads = 0;
      const warm = paragraphs.map(p => wordLineHeight(p));
      const warmReads = layoutReads;
      // Editing an existing text node and changing inherited font properties
      // must invalidate its measurement even when the element itself is reused.
      paragraphs[0].firstChild!.textContent = '한글과 日本語';
      const scriptChanged = wordLineHeight(paragraphs[0]);
      const probe = document.createElement('span');
      probe.style.cssText = 'position:fixed;font:400 11pt Verdana;white-space:pre;line-height:normal';
      probe.textContent = 'Ag한글과 日本語'; host.append(probe);
      const scriptExpected = original.call(probe).height * 1.15;
      paragraphs[0].style.fontSize = '24pt';
      const changed = wordLineHeight(paragraphs[0]);
      probe.style.fontSize = '24pt';
      const expected = original.call(probe).height * 1.15;
      return { coldReads, warmReads, stable: cold.every((height, i) => height === warm[i]), scriptChanged, scriptExpected, changed, expected };
    } finally { Element.prototype.getBoundingClientRect = original; host.remove(); }
  });
  expect(result.coldReads).toBeGreaterThan(2500);
  expect(result.warmReads).toBeLessThan(10);
  expect(result.stable).toBe(true);
  expect(result.scriptChanged).toBeCloseTo(result.scriptExpected, 5);
  expect(result.changed).toBeCloseTo(result.expected, 5);
});
