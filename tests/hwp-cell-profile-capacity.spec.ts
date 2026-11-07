import { test, expect } from '@playwright/test';

test('unchanged HWP terminal lines beyond the shared cache capacity reuse their measurements and invalidate on font changes', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const result = await page.evaluate(async () => {
    const { hwpCellTerminalLayout } = await import('/src/hwp-cell-terminal.ts');
    const host = document.createElement('div'); host.className = 'tiptap';
    host.style.cssText = 'position:fixed;left:-100000px;top:0;width:180px;visibility:hidden;pointer-events:none';
    const paragraphs = Array.from({ length: 600 }, (_, i) => {
      const p = document.createElement('p'); p.style.cssText = 'margin:0;padding:0;font:18px "Noto Serif KR";line-height:18px';
      p.textContent = String.fromCodePoint(0xac00 + i) + ' 검증 문단의 마지막 줄 높이를 확인합니다. '.repeat(2); host.append(p); return p;
    });
    document.body.append(host);
    await document.fonts.load('18px "Noto Serif KR"', paragraphs.map(p => p.textContent).join('')); await document.fonts.ready;
    let probes = 0;
    const observer = new MutationObserver(records => { for (const record of records) for (const n of record.addedNodes)
      if (n instanceof HTMLElement && n.style.position === 'fixed' && n.style.whiteSpace === 'pre') probes++; });
    observer.observe(document.body, { childList: true });
    const measure = async (advance: number) => {
      const rows = paragraphs.map(p => { const result = hwpCellTerminalLayout(p, 18, advance, 18); return result && { offset: result.offset, shift: result.shift }; });
      await Promise.resolve(); const count = probes; probes = 0; return { probes: count, rows };
    };
    try {
      const cold = await measure(24), warm = await measure(24), changed = await measure(26);
      document.fonts.dispatchEvent(new Event('loadingdone'));
      const invalidated = await measure(26);
      return { cold, warm, changed, invalidated, text: paragraphs.map(p => p.textContent) };
    } finally { observer.disconnect(); host.remove(); }
  });
  expect(result.cold.rows.filter(Boolean)).toHaveLength(600);
  expect(result.cold.probes).toBeGreaterThanOrEqual(600);
  expect(result.warm.rows).toEqual(result.cold.rows);
  expect(result.warm.probes).toBe(0);
  expect(result.changed.rows).not.toEqual(result.cold.rows);
  expect(result.changed.probes).toBeGreaterThanOrEqual(600);
  expect(result.invalidated.rows).toEqual(result.changed.rows);
  expect(result.invalidated.probes).toBeGreaterThanOrEqual(600);
  expect(result.text[0]).toContain('검증 문단의 마지막 줄');
});

test('a cached first letter does not replace a different fallback-glyph sample in the same paragraph', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { hwpCellTerminalLayout } = await import('/src/hwp-cell-terminal.ts');
    const p = document.createElement('p'); p.style.cssText = 'position:fixed;left:-100000px;top:0;width:150px;margin:0;padding:0;visibility:hidden;font:18px "Times New Roman";line-height:18px'; document.body.append(p);
    const samples = ['Aabcdefg', 'A😀🙂😀', 'A日本語書式', 'Aअक्षरदेवनागरी'];
    await document.fonts.load('18px "Times New Roman"', samples.join('')); await document.fonts.ready;
    const measure = (text: string) => { p.textContent = (text + ' ').repeat(12); const r = hwpCellTerminalLayout(p, 18, 24, 18); return r && { shift: r.shift, offset: r.offset }; };
    try {
      const cold = samples.map(text => { document.fonts.dispatchEvent(new Event('loadingdone')); return measure(text); });
      document.fonts.dispatchEvent(new Event('loadingdone'));
      const warm = samples.map(measure);
      return { cold, warm };
    } finally { p.remove(); }
  });
  expect(result.warm).toEqual(result.cold);
});
