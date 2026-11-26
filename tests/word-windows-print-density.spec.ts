import { expect, test } from '@playwright/test';

test.use({ deviceScaleFactor: 1.25 });

test('Windows print uses reference-density automatic Word line metrics without changing exact spacing', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const original = Element.prototype.getBoundingClientRect;
    const initialBridge = Object.getOwnPropertyDescriptor(window, 'wordDesktop');
    const targetFont = 'monospace';
    try {
      Object.defineProperty(window, 'wordDesktop', { configurable: true, value: { platform: 'win32' } });
      Element.prototype.getBoundingClientRect = function () {
        const box = original.call(this);
        if (!(this instanceof HTMLSpanElement) || this.style.lineHeight !== 'normal' ||
            this.style.left !== '-10000px' || this.style.fontFamily !== targetFont || this.style.fontSize !== '16px') return box;
        const zoom = this.currentCSSZoom;
        const normal = Math.abs(devicePixelRatio * zoom - 1) < .000001 ? 19 : 20;
        return new DOMRect(box.x, box.y, box.width, normal * zoom);
      };
      const heading = '<p data-word-line-rule="auto" data-word-line="1.2" style="font:16px monospace">Heading</p>';
      const table = '<table data-word-table-layout="fixed"><tbody><tr><td><p data-word-line-rule="auto" data-word-line="1.2" style="font:16px monospace">Cell</p></td></tr></tbody></table>';
      const fixed = '<p data-word-line-rule="exact" data-word-line="15pt" style="font:16px/20px monospace">Explicit spacing</p>';
      async function measure(nativeDesktop: boolean, withTable = true) {
        const layout = await createPrintLayout(heading + (withTable ? table : '') + fixed, defaultPage, undefined, undefined, { nativeDesktop });
        try {
          const paragraphs = [...layout.element.querySelectorAll('.print-content p[data-word-line-rule]')];
          return paragraphs.map(paragraph => [paragraph.textContent, parseFloat(getComputedStyle(paragraph).lineHeight)]);
        } finally { layout.dispose(); }
      }
      return { dpr: devicePixelRatio, browser: await measure(false), windows: await measure(true), flow: await measure(true, false) };
    } finally {
      Element.prototype.getBoundingClientRect = original;
      if (initialBridge) Object.defineProperty(window, 'wordDesktop', initialBridge);
      else Reflect.deleteProperty(window, 'wordDesktop');
    }
  });

  expect(result.dpr).toBe(1.25);
  expect(result.browser).toEqual([['Heading', 24], ['Cell', 24], ['Explicit spacing', 20]]);
  expect(result.windows).toEqual([['Heading', 22.8], ['Cell', 22.8], ['Explicit spacing', 20]]);
  expect(result.flow).toEqual([['Heading', 24], ['Explicit spacing', 20]]);
});
