'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');

for (const extension of ['hwp', 'hwpx']) {
  test(`landscape ${extension} paginates alternating table rows without clipping or losing text`, async () => {
    const { default: init, HwpDocument } = await import('@rhwp/core');
    await init({ module_or_path: await fs.readFile(path.join(path.dirname(require.resolve('@rhwp/core')), 'rhwp_bg.wasm')) });
    const document = new HwpDocument(await fs.readFile(path.join(__dirname, 'fixtures/hwp/landscape-pagination.' + extension)));
    try {
      // The header + 24 alternating 80/90 px rows cannot fit in three
      // physical frames. Upstream 0.8.7 returned three and clipped rows 15/16.
      assert.equal(document.pageCount(), 4);
      const allText = [];
      for (let page = 0; page < document.pageCount(); page++) {
        const svg = document.renderPageSvg(page);
        const height = Number(svg.match(/<svg[^>]*\bheight="([\d.]+)"/)[1]);
        const width = Number(svg.match(/<svg[^>]*\bwidth="([\d.]+)"/)[1]);
        const { runs } = JSON.parse(document.getPageTextLayout(page));
        for (const run of runs.filter(run => run.text.trim())) {
          assert.ok(run.y >= -0.5 && run.y + run.h <= height + 0.5, `Page ${page + 1}: ${run.text} outside physical height`);
          assert.ok(run.x >= -0.5 && run.x + run.w <= width + 0.5, `Page ${page + 1}: ${run.text} outside physical width`);
          allText.push(run.text);
        }
        assert.ok(runs.some(run => run.text === 'COLUMN-1'), 'Repeat the table header on every table fragment');
      }
      for (let row = 1; row <= 24; row++) for (let column = 1; column <= 2; column++) {
        const label = `ROW-${String(row).padStart(2, '0')}-COLUMN-${column}`;
        assert.equal(allText.filter(text => text === label).length, 1, `${label} must appear exactly once`);
      }
      assert.ok(JSON.parse(document.getPageTextLayout(3)).runs.some(run => run.text.includes('Word Studio')), 'Following text belongs after the last table fragment');
    } finally { document.free(); }
  });
}
