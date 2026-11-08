import { expect, type Page, type Locator } from '@playwright/test';
import { Document, Packer, Paragraph, Table, TableRow, TableCell } from 'docx';
import { readFile } from 'node:fs/promises';
export async function fixture() {
  return Packer.toBuffer(new Document({ sections: [{ children: Array.from({ length: 5 }, (_, p) => [
    new Paragraph({ text: `ZOOM PAGE ${p + 1}`, pageBreakBefore: p > 0 }),
    ...Array.from({ length: 10 }, (_, n) => new Paragraph(`PAGE ${p + 1} LINE ${n + 1} Korean 한글 content stays editable.`)),
    new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph(`TABLE PAGE ${p + 1}`)] }),new TableCell({ children: [new Paragraph('SECOND CELL')] })] })] }),
  ]).flat() }] }));
}
export async function zoomScale(viewport: Locator) {
  return viewport.evaluate(el => {
    if ((el as HTMLElement).dataset.scale) return Number((el as HTMLElement).dataset.scale);
    if ((el as HTMLElement).dataset.displayScale) return Number((el as HTMLElement).dataset.displayScale);
    return new DOMMatrix(getComputedStyle(el.querySelector('.scaled-document')!).transform).a;
  });
}
export async function wheel(page: Page, viewport: Locator, delta = -120, modified = true) {
  const r = (await viewport.boundingBox())!;
  await page.mouse.move(r.x + r.width / 2, r.y + r.height / 2);
  if (modified) await page.keyboard.down('Control');
  try { await page.mouse.wheel(0, delta); } finally { if (modified) await page.keyboard.up('Control'); }
}
export async function trackpad(page: Page, viewport: Locator, factor = 1.3) {
  const r = (await viewport.boundingBox())!, client = await page.context().newCDPSession(page);
  try { await client.send('Input.synthesizePinchGesture', { x: r.x + r.width / 2, y: r.y + r.height / 2, scaleFactor: factor, gestureSourceType: 'mouse' }); }
  finally { await client.detach(); }
}
export async function webkitPinch(viewport: Locator, factor = 1.3) {
  return viewport.evaluate((el, factor) => {
    const r = el.getBoundingClientRect(), cancelled = [];
    for (const [type, scale] of [['gesturestart', 1], ['gesturechange', factor], ['gestureend', factor]] as const) {
      const event = new Event(type, { bubbles: true, cancelable: true });
      Object.assign(event, { scale, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 });
      cancelled.push(!el.dispatchEvent(event));
    }
    return cancelled;
  }, factor);
}
export async function assertInputs(page: Page, viewport: Locator) {
  const before = await zoomScale(viewport);
  await wheel(page, viewport); await expect.poll(() => zoomScale(viewport)).toBeGreaterThan(before * 1.1);
  const enlarged = await zoomScale(viewport);
  await wheel(page, viewport, 120); await expect.poll(() => zoomScale(viewport)).toBeLessThan(enlarged * .95);
  const base = await zoomScale(viewport);
  await trackpad(page, viewport); await expect.poll(() => zoomScale(viewport)).toBeGreaterThan(base * 1.1);
  const pinchBase = await zoomScale(viewport);
  expect(await webkitPinch(viewport)).toEqual([true, true, true]);
  await expect.poll(() => zoomScale(viewport)).toBeCloseTo(pinchBase * 1.3, 2);
  const scrollBase = await zoomScale(viewport);
  await wheel(page, viewport, 120, false); await expect.poll(() => zoomScale(viewport)).toBeCloseTo(scrollBase, 5);
  expect(await page.evaluate(() => window.visualViewport?.scale)).toBe(1);
}
export async function saveModel(page: Page) {
  const waiting = page.waitForEvent('download'); await page.keyboard.press('Control+s');
  const download = await waiting, value = JSON.parse(await readFile((await download.path())!, 'utf8')); await download.delete();
  return { content: value.content, page: value.page, comments: value.comments };
}
export async function pointOnPage(viewport: Locator, paper: Locator) {
  const element = await paper.elementHandle();
  return viewport.evaluate((view, paper) => {
    const vr = view.getBoundingClientRect(), pr = paper!.getBoundingClientRect(), scale = new DOMMatrix(getComputedStyle(paper!).transform).a;
    return { x: (vr.x + vr.width / 2 - pr.x) / scale, y: (vr.y + vr.height / 2 - pr.y) / scale };
  }, element);
}
