import { test, expect } from '@playwright/test';

test('measure Word normal line metrics when a window changes device density', async ({ page }) => {
  await page.goto('/');
  const cdp = await page.context().newCDPSession(page);
  await page.evaluate(async () => {
    await document.fonts.ready;
    const { wordLineHeight } = await import('/src/word-line-metrics.ts');
    const records = [];
    for (const family of ['Arial', 'Calibri', 'Aptos', 'Times New Roman', 'Georgia', 'Cambria', 'Garamond', 'Malgun Gothic', 'system-ui', 'monospace'])
      for (const fontSize of [10, 11, 13, 14, 15, 16, 19, 21]) {
        const p = document.createElement('p');
        p.dataset.wordLineRule = 'auto'; p.dataset.wordLine = '1';
        p.style.cssText = `position:fixed;visibility:hidden;font-family:${family};font-size:${fontSize}px`;
        p.textContent = 'Ag test.';
        document.body.append(p);
        records.push({p,family,fontSize});
      }
    (window as any).densityWitness = () => records.map(({p,family,fontSize}) => {
      const font = getComputedStyle(p);
      const probe = document.createElement('span');
      probe.style.cssText = `position:fixed;left:-10000px;visibility:hidden;font:${font.fontStyle} ${font.fontWeight} ${font.fontSize} ${font.fontFamily};line-height:normal;white-space:pre;padding:0;border:0`;
      probe.textContent = p.textContent; document.body.append(probe);
      const native = probe.getBoundingClientRect().height;
      probe.remove();
      return { family,fontSize,native,word:wordLineHeight(p) };
    });
  });
  const result = [];
  for (const scale of [1,1.25,1.5,2,1]) {
    await cdp.send('Emulation.setDeviceMetricsOverride',{ width:1280,height:980,deviceScaleFactor:scale,mobile:false });
    await expect.poll(() => page.evaluate(() => devicePixelRatio)).toBe(scale);
    result.push({scale, measurements:await page.evaluate(() => (window as any).densityWitness())});
  }
  const base = result[0].measurements;
  for (const {scale, measurements} of result) {
    const changed=measurements.map((m,i)=>[base[i],m]).filter(([a,b])=>a.native!==b.native);
    console.log(JSON.stringify({scale,changed:changed.length,stale:changed.filter(([a,b])=>a.word===b.word).length,sample:changed.slice(0,5)}));
  }
});
