'use strict';
// Numeric diagnostics for repository-licensed faces only. No Word documents,
// account state, local installed font binaries or private PDF fonts are read.
const fs = require('node:fs/promises'), path = require('node:path');
const { createRequire } = require('node:module');
const root = path.resolve(__dirname, '../../..'), req = createRequire(root + '/package.json');
(async () => {
  const { chromium } = req('@playwright/test');
  const channel = process.env.METRIC_BROWSER_CHANNEL;
  const browser = await chromium.launch({ headless: true, ...(channel ? { channel } : {}) });
  const faces = [];
  for (const [style, weight, slope, ascent] of [
    ['regular', '400', 'normal', 1989], ['bold', '700', 'normal', 2032],
    ['italic', '400', 'italic', 1989], ['bolditalic', '700', 'italic', 2060],
  ]) faces.push({ style, weight, slope, ascent, data: (await fs.readFile(root + '/src/assets/fonts/texgyreadventor-' + style + '.otf')).toString('base64') });
  const rows = [];
  try {
    for (const dpr of [1, 2]) {
      const page = await browser.newPage({ deviceScaleFactor: dpr });
      await page.goto('about:blank');
      rows.push(...await page.evaluate(async ({ faces, dpr }) => {
        const rows = [];
        for (const face of faces) {
          const family = 'Diagnostic_' + face.style;
          const font = new FontFace(family, 'url(data:font/otf;base64,' + face.data + ')', {
            weight: face.weight, style: face.slope,
            ascentOverride: face.ascent / 2048 * 100 + '%', descentOverride: 451 / 2048 * 100 + '%', lineGapOverride: '0%',
          });
          await font.load(); document.fonts.add(font);
          for (const points of [9, 10, 11, 11.3, 12, 14, 16, 18, 28]) {
            const span = document.createElement('span');
            span.style.cssText = `position:absolute;display:inline-block;white-space:pre;line-height:normal;font-family:${family};font-size:${points}pt;font-weight:${face.weight};font-style:${face.slope}`;
            span.textContent = 'Typography AV To 123'; document.body.append(span);
            const normal = span.getBoundingClientRect().height, normalAdvance = span.getBoundingClientRect().width;
            const design = points * 96 / 72 * (face.ascent + 451) / 2048;
            const pitches = {};
            for (const [mode, value] of [['browserNormal', normal], ['fractionalDesign', design]]) {
              const p = document.createElement('p'); p.style.cssText = span.style.cssText + `;margin:0;line-height:${value * 1.15}px`;
              p.innerHTML = 'FIRST<br>SECOND<br>THIRD'; document.body.append(p);
              const lines = [...p.childNodes].filter(n => n.nodeType === Node.TEXT_NODE).map(n => {
                const range = document.createRange(); range.selectNodeContents(n); return range.getBoundingClientRect().y;
              });
              pitches[mode] = { requestedCSSpx: value * 1.15, observedCSSpx: lines.slice(1).map((y, i) => y - lines[i]), paragraphHeightCSSpx: p.getBoundingClientRect().height };
              p.remove();
            }
            rows.push({ style: face.style, points, dpr, normalCSSpx: normal, normalAdvanceCSSpx: normalAdvance, fractionalDesignCSSpx: design, pitches });
            span.remove();
          }
        }
        return rows;
      }, { faces, dpr }));
      await page.close();
    }
    console.log(JSON.stringify({ method: 'Actual Chromium DOM font struts and three-line pitches. Licensed fallback bytes with the existing public Word-profile overrides. Numeric diagnostic only; neither an original Microsoft font measurement nor a product correction coefficient.', platform: process.platform, browser: browser.version(), rows }));
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
