import { expect, test } from '@playwright/test';
import sharp from 'sharp';

test('Windows measured hanging source line witnesses across installed and bundled faces', async ({ page }) => {
  await page.goto('/');
  const results = await page.evaluate(async () => {
    const { layoutHwpBodyLines, restoreHwpAuthoredBodyLines } = await import('/src/hwp-body-line-metrics.ts');
    const { horizontalInkFits, sourceLineBoundaries } = await import('/src/hwp-source-cell-lines.ts');
    const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
    const { encodeHwpBodyLine } = await import('/src/hwp-body-line.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    await document.fonts.load('16px "Studio Text"', 'AAAA');
    const literal = 'AAAA BBBB CCCC DDDD', sheet = document.createElement('div');
    sheet.className = 'print-page';
    sheet.style.cssText = 'position:fixed;top:0;left:0;width:400px;height:400px';
    sheet.dataset.page = '1';
    sheet.innerHTML = '<div class="tiptap print-content" style="position:absolute;top:10px;left:10px;width:240px;height:360px;padding:0;min-height:0;box-sizing:border-box"></div>';
    document.body.append(sheet);
    try {
      const body = sheet.querySelector<HTMLElement>('.print-content')!, paper = sheet.getBoundingClientRect();
      const output = [];
      for (const face of ['monospace', 'Consolas', 'Courier New', 'Arial', '"Studio Text"']) for (const small of [16, 14, 12]) {
        const p = document.createElement('p');
        p.style.cssText = 'font-size:16px;line-height:24px;min-height:0;margin:0 0 0 30px;text-indent:-30px;text-align:justify;position:relative;left:4px';
        p.style.fontFamily = face;
        p.innerHTML = '<span style="font-size:16px">AAAA BBBB </span><span style="font-size:' + small + 'px">CCCC DDDD</span>';
        if (small !== 16) p.dataset.hwpBodyLine = encodeHwpBodyLine({ percent: 150, font: 16, base: 24 });
        body.replaceChildren(p);
        layoutHwpBodyLines(body, true);
        const box = p.getBoundingClientRect();
        const before = { height: box.height, offsets: paragraphLineOffsets(p), lineHeight: getComputedStyle(p).lineHeight };
        p.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(literal), ends:[8,16],
          authoredBodyLines: { hanging: 30, size: 16, pitch: 24, width: 240,
            origin: { page: 1, x: box.left - paper.left - 30, y: box.top - paper.top } } });
        const undo = restoreHwpAuthoredBodyLines(p);
        const accepted = Boolean(undo);
        undo?.();
        const boundaries = sourceLineBoundaries(p, [8]);
        const range = document.createRange();
        const at = boundaries?.[1];
        if (!at) throw new Error('Test boundary unavailable');
        range.setStart(at.node, at.at); range.collapse(true); range.insertNode(document.createElement('br'));
        if (small !== 16) layoutHwpBodyLines(p, true);
        const after = { height: p.getBoundingClientRect().height, offsets: paragraphLineOffsets(p),
          inkFits: horizontalInkFits(p, range, body.getBoundingClientRect()) };
        output.push({ face, small, before, after, accepted, retained: p.textContent === literal });
      }
      return output;
    } finally { sheet.remove(); }
  });
  expect(results).toHaveLength(15);
  expect(results.every(item => item.retained)).toBe(true);
  console.log('WINDOWS_HANGING_WITNESS ' + JSON.stringify(results));
});

test('Windows print image CSS dimensions, raster clip and authored red ink', async ({ browser, baseURL }) => {
  for (const dpr of [1, 1.25, 2]) {
    const context = await browser.newContext({ baseURL, viewport: { width: 1440, height: 1000 }, deviceScaleFactor: dpr });
    try {
      const page = await context.newPage();
      await page.goto('/');
      for (const authored of [true, false]) {
        await page.evaluate(async authored => {
          document.querySelector('.print-layout')?.remove();
          const { createPrintLayout } = await import('/src/print-layout.ts'), { defaultPage } = await import('/src/types.ts');
          const canvas = document.createElement('canvas');
          canvas.width = canvas.height = 1;
          const context = canvas.getContext('2d')!;
          context.fillStyle = '#ed2010'; context.fillRect(0, 0, 1, 1);
          const layout = await createPrintLayout('<p><img ' + (authored ? 'data-inline-image="true"' : '') +
            ' src="' + canvas.toDataURL() + '" width="120" height="30" style="width:120px;height:30px" alt="FRAME"></p>',
          { ...defaultPage, pageNumbers: false });
          layout.element.classList.replace('measuring', 'preview');
          document.body.classList.add('studio-printing');
        }, authored);
        await page.emulateMedia({ media: 'print' });
        const target = page.locator('.print-content img[alt="FRAME"]');
        const extent = await target.evaluate(element => {
          const box = element.getBoundingClientRect();
          return { x: box.x, y: box.y, width: box.width, height: box.height, fit: getComputedStyle(element).objectFit };
        });
        const screenshot = await target.screenshot({ scale: 'css' });
        const { data, info } = await sharp(screenshot).removeAlpha().raw().toBuffer({ resolveWithObject: true });
        let left = info.width, right = -1, top = info.height, bottom = -1;
        for (let y = 0; y < info.height; y++) for (let x = 0; x < info.width; x++) {
          const offset = (y * info.width + x) * info.channels;
          if (data[offset] <= 150 || data[offset + 1] >= 100 || data[offset + 2] >= 100) continue;
          left = Math.min(left, x); right = Math.max(right, x);
          top = Math.min(top, y); bottom = Math.max(bottom, y);
        }
        const ink = [right - left + 1, bottom - top + 1], frame = [info.width, info.height];
        console.log('WINDOWS_PRINT_IMAGE_WITNESS ' + JSON.stringify({ dpr, authored, extent, frame, ink }));
        expect([extent.width, extent.height]).toEqual([120, 30]);
        await page.emulateMedia({ media: 'screen' });
      }
    } finally { await context.close(); }
  }
});
