import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

for (const authored of [true, false]) test('print ' + (authored ? 'Word inline' : 'ordinary block') + ' artwork uses its intended image frame', async ({ page }, info) => {
  await page.goto('/');
  await page.evaluate(async authored => {
    const { createPrintLayout } = await import('/src/print-layout.ts'), { defaultPage } = await import('/src/types.ts');
    const canvas = document.createElement('canvas');
    canvas.width = canvas.height = 1;
    const context = canvas.getContext('2d')!;
    context.fillStyle = '#ed2010';
    context.fillRect(0, 0, 1, 1);
    const inline = authored ? 'data-inline-image="true"' : '';
    const layout = await createPrintLayout('<p><img ' + inline + ' src="' + canvas.toDataURL() +
      '" width="120" height="30" style="width:120px;height:30px" alt="FRAME"></p>',
    { ...defaultPage, margin: authored ? defaultPage.margin : 25.5, pageNumbers: false });
    layout.element.classList.replace('measuring', 'preview');
    document.body.classList.add('studio-printing');
  }, authored);

  await page.emulateMedia({ media: 'print' });
  const target = page.locator('.print-content img[alt="FRAME"]');
  await target.scrollIntoViewIfNeeded();
  const box = await target.evaluate(element => {
    const { x, y, width, height } = element.getBoundingClientRect();
    return { x: x + window.scrollX, y: y + window.scrollY, width, height };
  });
  expect([box.width, box.height]).toEqual([120, 30]);

  const png = await target.screenshot({ scale: 'css' });
  await writeFile(info.outputPath('printed-frame.png'), png);
  const ink = await page.evaluate(async encoded => {
    const image = new Image();
    image.src = 'data:image/png;base64,' + encoded;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;
    const context = canvas.getContext('2d')!;
    context.drawImage(image, 0, 0);
    const data = context.getImageData(0, 0, canvas.width, canvas.height).data;
    let left = canvas.width, right = -1, top = canvas.height, bottom = -1;
    for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) {
      const i = (y * canvas.width + x) * 4;
      if (data[i] <= 150 || data[i + 1] >= 100 || data[i + 2] >= 100) continue;
      left = Math.min(left, x); right = Math.max(right, x);
      top = Math.min(top, y); bottom = Math.max(bottom, y);
    }
    return { width: right - left + 1, height: bottom - top + 1, frame: [canvas.width, canvas.height] };
  }, png.toString('base64'));
  await writeFile(info.outputPath('printed-frame-ink.json'), JSON.stringify({ box, ...ink }));
  // Element screenshots enclose a fractional document rectangle in whole
  // CSS pixels. The exact frame measured above should still be 120 by 30.
  const enclosing = (start: number, size: number) => Math.ceil(start + size - .001) - Math.floor(start + .001);
  expect(ink.frame).toEqual([enclosing(box.x, box.width), enclosing(box.y, box.height)]);
  expect(ink.height).toBeGreaterThanOrEqual(29);
  expect(ink.width).toBeGreaterThanOrEqual(authored ? 119 : 29);
  if (!authored) expect(ink.width).toBeLessThanOrEqual(31);
  await page.pdf({ path: info.outputPath('printed-frame.pdf'), preferCSSPageSize: true, printBackground: true });
});
