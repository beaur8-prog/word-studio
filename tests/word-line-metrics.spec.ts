import { test, expect } from '@playwright/test';

// Browser font metrics are the oracle here, not fixed pixels tied to one OS.
// These synthetic paragraphs complement the frozen, Word-authored DOCX corpus.
test('auto spacing measures the actual script and does not cache a hypothetical fallback', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { wordLineHeight } = await import('/src/word-line-metrics.ts');
    const root = document.createElement('div'); document.body.append(root);
    const normal = (font: string, text: string) => {
      const probe = document.createElement('span');
      Object.assign(probe.style, { position: 'fixed', font, whiteSpace: 'pre', lineHeight: 'normal' });
      probe.textContent = text; root.append(probe);
      const height = probe.getBoundingClientRect().height; probe.remove(); return height;
    };
    const measure = (html: string) => {
      const p = document.createElement('p');
      p.style.cssText = 'font:400 11pt Verdana; margin:0';
      p.dataset.wordLineRule = 'auto'; p.dataset.wordLine = '1.15'; p.innerHTML = html; root.append(p);
      const base = normal(getComputedStyle(p).font, 'Ag');
      let expected = base;
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        expected = Math.max(expected, normal(getComputedStyle(node.parentElement!).font, node.data.replace(/[\t\r\n\u2028\u2029]/g, ' ')));
      }
      const result = { html, actual: wordLineHeight(p), expected: expected * 1.15 };
      p.remove(); return result;
    };
    try { return [
      measure('English body only'), measure('한글과 日本語'), measure('English body only'),
      measure('Latin <span style="font-size:6pt">한글</span>'),
      measure('Latin <strong><span style="font-size:18pt">English</span></strong>'),
      measure('Arabic العربية Hindi हिन्दी'), measure('Family 👨‍👩‍👧‍👦 and 🧑🏽‍💻 © ∑'),
      measure('<a style="font-size:24pt">한글 link</a>'),
      measure('<span style="font-size:24pt"> </span>'), measure('<br>'),
      measure('A\nB\tC\u2028D'),
    ]; } finally { root.remove(); }
  });
  for (const state of states) expect(state.actual, state.html).toBeCloseTo(state.expected, 5);
});

test('legacy Korean profiles, minimum bounds and exact source spacing remain distinct', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async () => {
    const { wordLineHeight, layoutWordLines } = await import('/src/word-line-metrics.ts');
    const root = document.createElement('div'); document.body.append(root);
    const p = document.createElement('p'); p.style.cssText = 'font:12px DotumChe;line-height:17px'; root.append(p);
    p.dataset.wordLineRule = 'auto'; p.dataset.wordLine = '2';
    p.innerHTML = '<span>한글</span>'; const korean = wordLineHeight(p);
    p.innerHTML = '<span>･</span>'; const fallback = wordLineHeight(p);
    p.textContent = '･'; const directFallback = wordLineHeight(p);
    p.dataset.wordLineRule = 'atLeast'; p.dataset.wordLine = '20pt'; const minimum = wordLineHeight(p);
    p.dataset.wordLine = '2pt'; const smallerMinimum = wordLineHeight(p);
    p.dataset.wordLineRule = 'exact'; p.dataset.wordLine = '9pt'; layoutWordLines(root);
    const exact = { result: wordLineHeight(p), css: p.style.lineHeight };
    p.dataset.wordLineRule = 'auto'; p.dataset.wordLine = 'invalid'; const invalid = wordLineHeight(p);
    root.remove(); return { korean, fallback, directFallback, minimum, smallerMinimum, exact, invalid };
  });
  expect(states.korean).toBeCloseTo(12 * 1.3 * 2, 5);
  expect(states.fallback).toBeCloseTo(12 * 1.55 * 2, 5);
  expect(states.directFallback).toBeCloseTo(states.fallback!, 5);
  expect(states.minimum).toBeCloseTo(20 * 4 / 3, 5);
  expect(states.smallerMinimum).toBeCloseTo(12 * 1.55, 5);
  expect(states.exact).toEqual({ result: null, css: '17px' });
  expect(states.invalid).toBeNull();
});

test('editor remeasures after script edits, undo and late font loading without saving pixel leading', async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    return { format: 'word-studio', version: 1, id: 'line-metrics', title: 'Line metrics', updatedAt: new Date().toISOString(), page: defaultPage, comments: [],
      content: parseDocumentHtml('<p data-word-line-rule="auto" style="font-family:LateMetric,Verdana;font-size:11pt;line-height:1.15;margin:0">Latin only</p>') };
  });
  await page.locator('input[type=file]').first().setInputFiles({ name: 'metrics.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  const paragraph = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('p').first();
  await expect(paragraph).toHaveText('Latin only');
  const error = () => paragraph.evaluate(p => {
    const probe = document.createElement('span');
    const style = getComputedStyle(p);
    Object.assign(probe.style, { position: 'fixed', fontFamily: style.fontFamily, fontSize: style.fontSize,
      fontWeight: style.fontWeight, fontStyle: style.fontStyle, lineHeight: 'normal', whiteSpace: 'pre' });
    probe.textContent = 'Ag' + p.textContent; document.body.append(probe);
    const expected = probe.getBoundingClientRect().height * 1.15; probe.remove();
    return Math.abs(parseFloat(getComputedStyle(p).lineHeight) - expected);
  });
  await expect.poll(error).toBeLessThan(.01);
  const latin = await paragraph.evaluate(p => getComputedStyle(p).lineHeight);
  await paragraph.fill('한글과 日本語'); await expect.poll(error).toBeLessThan(.01);
  await page.keyboard.press('ControlOrMeta+z'); await expect(paragraph).toHaveText('Latin only');
  await expect(paragraph).toHaveCSS('line-height', latin);
  await page.evaluate(async () => {
    const face = new FontFace('LateMetric', 'url(/src/assets/fonts/NotoSansMono.ttf)', { ascentOverride: '150%', descentOverride: '40%', lineGapOverride: '20%' });
    document.fonts.add(face); await face.load(); await document.fonts.ready;
  });
  await expect.poll(error).toBeLessThan(.01);
  await expect.poll(() => paragraph.evaluate(p => getComputedStyle(p).lineHeight)).not.toBe(latin);
  const pending = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const file = await pending;
  const fs = await import('node:fs/promises');
  const saved = JSON.parse(await fs.readFile((await file.path())!, 'utf8'));
  expect(saved.content.content[0].attrs.wordLineRule).toBe('auto');
  expect(saved.content.content[0].attrs.lineSpacing).toBe('1.15');
  expect(saved.content.content[0].content[0].text).toBe('Latin only');
});
