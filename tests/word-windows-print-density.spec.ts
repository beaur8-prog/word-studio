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

test('a regular strut reserves its measured space around a short bold tabbed line', async ({page}) => {
  await page.goto('/');
  await page.locator('#root .tiptap[role="textbox"]').waitFor();
  const result = await page.evaluate(async () => {
    const { reserveWordTabbedStruts } = await import('/src/word-native-tabbed-strut.ts');
    const host = document.createElement('div');
    host.style.cssText = 'width:400px;height:500px;position:absolute;left:20px;top:20px';
    const item = (line: string, type = 'auto', spill = '1') =>
      '<p data-test-gap="' + spill + '" data-word-line-rule="' + type +
      '" style="font:400 14.6667px/20.4492px Arial;margin:0">' + line + '</p>';
    const bold = '<u style="font-weight:700">10<span data-docx-tab style="display:inline-block;width:30px">\t</span>Detailed outcome</u>';
    host.innerHTML = [
      item(bold), item(bold, 'exact'), item(bold + ' ordinary'),
      item('<span data-docx-tab>\t</span> regular'), item(bold,'auto','2.5'),
    ].join('');
    document.body.append(host);
    const bounds = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function () {
      const raw = bounds.call(this);
      if (!(this instanceof HTMLElement) || !this.dataset.testGap) return raw;
      const scale = this.currentCSSZoom;
      const height = 20.4492 + (scale !== 1 ? Number(this.dataset.testGap) : 0);
      return new DOMRect(raw.x, raw.y, raw.width, height * scale);
    };
    try {
      const visited = new WeakSet<HTMLElement>();
      reserveWordTabbedStruts(host, visited);
      const paragraphs = [...host.querySelectorAll('p')];
      const reserved = paragraphs.map(p => parseFloat(p.style.minHeight) || 0);
      const original = paragraphs[0].textContent;
      reserveWordTabbedStruts(host, visited);
      return { reserved, unchanged: paragraphs[0].textContent === original,
        weight: getComputedStyle(paragraphs[0].querySelector('u')!).fontWeight, dpr: devicePixelRatio };
    } finally {
      Element.prototype.getBoundingClientRect = bounds;
      host.remove();
    }
  });
  expect(result.dpr).toBe(1.25);
  expect(result.reserved[0]).toBeCloseTo(21.4492, 3);
  expect(result.reserved.slice(1)).toEqual([0, 0, 0, 0]);
  expect(result.unchanged).toBe(true);
  expect(result.weight).toBe('700');
});
