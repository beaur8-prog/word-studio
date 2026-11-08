import { test, expect } from '@playwright/test';
import path from 'node:path';
import { fixture, assertInputs, saveModel, zoomScale, wheel, pointOnPage, trackpad } from './zoom-test-helpers';
async function open(page: import('@playwright/test').Page) {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'zoom.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await fixture() });
  await expect(page.locator('.tiptap')).toContainText('ZOOM PAGE 5');
}
test('PC editor wheel and real trackpad pinch preserve paper/model, cursor anchor and undo history', async ({ page }) => {
  await open(page); const editor = page.locator('.tiptap'), view = page.locator('.document-workspace');
  await editor.locator('p').first().click(); await page.keyboard.press('End'); await page.keyboard.type(' ZOOM EDIT');
  const original = await saveModel(page), count = await page.locator('.document-stack').getAttribute('data-pages');
  await page.getByLabel('확대 축소', { exact: true }).fill('200');
  await expect.poll(() => zoomScale(view)).toBe(2);
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  await view.evaluate(el => { el.scrollTop = 2400; el.scrollLeft = 150; });
  await expect.poll(() => view.evaluate(el => el.scrollTop)).toBe(2400);
  const paper = page.locator('.scaled-document'), before = await pointOnPage(view, paper);
  await wheel(page, view); await expect.poll(() => zoomScale(view)).toBeGreaterThan(2.1);
  const after = await pointOnPage(view, paper); expect(Math.abs(after.x - before.x) * await zoomScale(view)).toBeLessThan(2); expect(Math.abs(after.y - before.y) * await zoomScale(view)).toBeLessThan(2);
  await page.getByRole('button', { name: /%$/, exact: false }).last().click();
  await assertInputs(page, view);
  expect(await page.locator('.document-stack').getAttribute('data-pages')).toBe(count);
  expect(await saveModel(page)).toEqual(original);
  await editor.locator('p').first().click(); await page.keyboard.press('ControlOrMeta+z'); await expect(editor).not.toContainText('ZOOM EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('ZOOM EDIT');
});
test('PC reading mode has document gestures while the UI stays its original size', async ({ page }) => {
  await open(page);
  await page.getByRole('button', { name: '읽기 모드 전환', exact: true }).click();
  await expect(page.locator('.tiptap')).toHaveAttribute('contenteditable', 'false');
  const ribbon = await page.locator('.ribbon').boundingBox();
  await assertInputs(page, page.locator('.document-workspace'));
  expect((await page.locator('.ribbon').boundingBox())?.height).toBe(ribbon?.height);
  await expect(page.locator('.tiptap')).toContainText('TABLE PAGE 5');
});
test('PC HWP original zoom anchors a later virtual page without resetting its position', async ({ page }) => {
  await page.goto('/'); await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/landscape-pagination.hwp'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true }), view = reader.locator('.hwp-viewport');
  await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByLabel('HWP 쪽 번호').fill('2'); await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '2쪽');
  const paper = reader.locator('[data-hwp-page="2"] .hwp-sheet');
  await reader.getByLabel('HWP 확대', { exact: true }).selectOption('125');
  await view.evaluate(el => el.scrollBy(0, 150));
  const before = await pointOnPage(view, paper), base = await zoomScale(view);
  await wheel(page, view); await expect.poll(() => zoomScale(view)).toBeGreaterThan(base * 1.1);
  const after = await pointOnPage(view, paper); expect(Math.abs(after.x - before.x) * await zoomScale(view)).toBeLessThan(2); expect(Math.abs(after.y - before.y) * await zoomScale(view)).toBeLessThan(2);
  await assertInputs(page, view);
  await expect(reader.getByLabel('HWP 확대', { exact: true })).not.toHaveValue('');
  await expect(reader.getByTestId('hwp-page-count')).toHaveText('/ 4');
});
test('print preview gestures keep source pages and print-media geometry unchanged', async ({ page }) => {
  await open(page); const original = await saveModel(page);
  await page.getByRole('button', { name: '보기', exact: true }).click(); await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const view = page.locator('.print-layout.preview'), pages = view.locator(':scope > .print-page'); await expect(view).toBeVisible();
  const dimensions = await pages.evaluateAll(nodes => nodes.map(el => ({ width: (el as HTMLElement).style.width, height: (el as HTMLElement).style.height }))), count = await pages.count();
  await assertInputs(page, view);
  expect(await pages.count()).toBe(count); expect(await pages.evaluateAll(nodes => nodes.map(el => ({ width: (el as HTMLElement).style.width, height: (el as HTMLElement).style.height })))).toEqual(dimensions);
  await page.emulateMedia({ media: 'print' });
  const printed = await pages.evaluateAll(nodes => nodes.map(el => ({ zoom: getComputedStyle(el).zoom, width: el.getBoundingClientRect().width, height: el.getBoundingClientRect().height })));
  printed.forEach((p, i) => { expect(Number(p.zoom)).toBe(1); expect(p.width).toBeCloseTo(parseFloat(dimensions[i].width), 0); expect(p.height).toBeCloseTo(parseFloat(dimensions[i].height), 0); });
  await page.emulateMedia({ media: 'screen' });
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click(); expect(await saveModel(page)).toEqual(original);
});
