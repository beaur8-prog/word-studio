import { test, expect } from '@playwright/test';
test.use({ viewport: { width: 412, height: 840 }, hasTouch: true, userAgent: 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 Chrome/130.0.0.0 Mobile Safari/537.36' });

test('a new Android document has a full-size page and horizontal typing from the first screen', async ({ page }, testInfo) => {
  await page.goto('/'); await page.getByRole('button', { name: '새 문서', exact: true }).click();
  const editor = page.getByLabel('Android 문서 편집 영역'); await expect(editor).toHaveAttribute('contenteditable', 'true');
  await editor.click(); await page.keyboard.type('FIRST LINE MUST STAY HORIZONTAL');
  const metrics = await editor.evaluate(el => {
    const stack = el.closest<HTMLElement>('.document-stack')!;
    const p = el.querySelector('p')!, node = p.firstChild!;
    const first = document.createRange(); first.setStart(node, 0); first.setEnd(node, 1);
    const next = document.createRange(); next.setStart(node, 1); next.setEnd(node, 2);
    const a = first.getBoundingClientRect(), b = next.getBoundingClientRect();
    return { paperWidth: stack.offsetWidth, paperHeight: Number(stack.dataset.pageHeight), paperVisualWidth: stack.getBoundingClientRect().width,
      first: { x: a.x, y: a.y }, second: { x: b.x, y: b.y }, viewportWidth: window.innerWidth };
  });
  await testInfo.attach('new-page-metrics', { body: JSON.stringify(metrics, null, 2), contentType: 'application/json' });
  expect(metrics.paperWidth).toBeGreaterThan(790);
  expect(metrics.paperHeight).toBeGreaterThan(1120);
  expect(Math.abs(metrics.first.y - metrics.second.y)).toBeLessThan(1);
  expect(metrics.second.x).toBeGreaterThan(metrics.first.x);
  expect(metrics.paperVisualWidth).toBeGreaterThan(metrics.viewportWidth - 40);
  expect(metrics.paperVisualWidth).toBeLessThanOrEqual(metrics.viewportWidth);
  await page.screenshot({ path: testInfo.outputPath('new-page-after.png') });
});

import { Document, Packer, Paragraph, Table, TableRow, TableCell } from 'docx';
import path from 'node:path';
import type { Page } from '@playwright/test';
async function openPages(page: Page) {
  await page.goto('/');
  const buffer = await Packer.toBuffer(new Document({ sections: [{ children: [1, 2, 3].map(i => new Paragraph({ text: `PRESENTATION PAGE ${i}`, pageBreakBefore: i > 1 })) }] }));
  await page.locator('input[type=file]').setInputFiles({ name: 'present.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await expect(page.locator('.android-paper-viewport')).toBeVisible();
  await page.getByRole('button', { name: '전체 화면 뷰어', exact: true }).click();
  await expect(page.locator('.android-presentation-page')).toContainText('PRESENTATION PAGE 1');
}
async function swipe(page: Page, from: [number, number], to: [number, number]) {
  const session = await page.context().newCDPSession(page);
  await session.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: from[0], y: from[1], id: 1 }] });
  for (let i = 1; i <= 4; i++) await session.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: from[0] + (to[0] - from[0]) * i / 4, y: from[1] + (to[1] - from[1]) * i / 4, id: 1 }] });
  await session.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] }); await session.detach();
}
async function reveal(page: Page) {
  if (!await page.getByRole('button', { name: '뷰어 확대', exact: true }).isVisible()) await page.locator('.android-presentation-viewport').tap({ position: { x: 200, y: 300 } });
}

test('fullscreen viewer uses the screen, hides toolbars and turns pages in four swipe directions', async ({ page }, testInfo) => {
  await openPages(page);
  const viewer = page.locator('.android-presentation');
  await page.evaluate(() => document.documentElement.style.setProperty('--office-cutout-top', '24px'));
  expect((await page.getByRole('button', { name: '뷰어 닫기', exact: true }).boundingBox())!.y).toBeGreaterThanOrEqual(24);
  expect((await page.locator('.android-presentation-viewport').boundingBox())!.y).toBe(0);
  await expect(page.locator('.android-topbar')).not.toBeVisible();
  const rect = await viewer.boundingBox(); expect(rect?.width).toBe(412); expect(rect?.height).toBeGreaterThanOrEqual(800);
  await swipe(page, [330, 400], [70, 400]); await expect(viewer).toHaveAttribute('data-page', '2');
  await expect(page.locator('.android-presentation-page')).toContainText('PRESENTATION PAGE 2');
  await swipe(page, [70, 400], [330, 400]); await expect(viewer).toHaveAttribute('data-page', '1');
  await swipe(page, [205, 650], [205, 250]); await expect(viewer).toHaveAttribute('data-page', '2');
  await swipe(page, [205, 250], [205, 650]); await expect(viewer).toHaveAttribute('data-page', '1');
  await swipe(page, [70, 400], [330, 400]); await expect(viewer).toHaveAttribute('data-page', '1');
  await page.screenshot({ path: testInfo.outputPath('android-fullscreen-viewer.png') });
  await page.getByRole('button', { name: '뷰어 닫기', exact: true }).click(); await expect(viewer).toHaveCount(0);
  await expect(page.locator('.android-paper-viewport')).toBeVisible();
});

test('pinch, pan and button zoom stay on the current page; zoomed content pans before page navigation', async ({ page }) => {
  await openPages(page);
  const viewer = page.locator('.android-presentation'), paper = page.locator('.android-presentation-page'), viewport = page.locator('.android-presentation-viewport');
  const before = (await paper.boundingBox())!.width;
  const session = await page.context().newCDPSession(page);
  await session.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: 150, y: 350, id: 1 }, { x: 250, y: 450, id: 2 }] });
  await session.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: 85, y: 285, id: 1 }, { x: 315, y: 515, id: 2 }] });
  await expect.poll(async () => (await paper.boundingBox())!.width).toBeGreaterThan(before * 1.8);
  // Release one finger and move the remaining one: this is not a page swipe.
  await session.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [{ x: 315, y: 515, id: 2 }] });
  await session.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: 70, y: 515, id: 2 }] });
  await session.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] }); await session.detach();
  await expect(viewer).toHaveAttribute('data-page', '1');
  const oldScroll = await viewport.evaluate(el => el.scrollLeft);
  await swipe(page, [300, 400], [150, 400]);
  await expect(viewer).toHaveAttribute('data-page', '1');
  expect(await viewport.evaluate(el => el.scrollLeft)).toBeGreaterThan(oldScroll);
  const oldTop = await viewport.evaluate(el => el.scrollTop);
  await swipe(page, [205, 600], [205, 420]);
  await expect(viewer).toHaveAttribute('data-page', '1');
  expect(await viewport.evaluate(el => el.scrollTop)).toBeGreaterThan(oldTop);
  await reveal(page); await page.getByRole('button', { name: '뷰어 폭 맞춤', exact: true }).click();
  await expect.poll(async () => (await paper.boundingBox())!.width).toBeCloseTo(before, 0);
  await page.getByRole('button', { name: '뷰어 확대', exact: true }).click(); await expect.poll(async () => (await paper.boundingBox())!.width).toBeGreaterThan(before);
  await page.getByRole('button', { name: '뷰어 축소', exact: true }).click(); await expect.poll(async () => (await paper.boundingBox())!.width).toBeCloseTo(before, 0);
  await page.getByRole('button', { name: '다음 페이지', exact: true }).click(); await expect(viewer).toHaveAttribute('data-page', '2');
});

test('Android native fullscreen is requested and restored without changing the Word document', async ({ page }) => {
  await page.addInitScript(() => { (window as any).systemBars = []; (window as any).wordAndroid = { setImmersiveMode: async (enabled: boolean) => { (window as any).systemBars.push(enabled); } }; });
  await openPages(page);
  expect(await page.evaluate(() => (window as any).systemBars.includes(true))).toBe(true);
  await page.getByRole('button', { name: '뷰어 닫기', exact: true }).click();
  await page.getByRole('button', { name: '뒤로', exact: true }).click();
  await expect(page.getByText('문서를 열어보세요', { exact: true })).toBeVisible();
  await expect.poll(() => page.evaluate(() => (window as any).systemBars.at(-1))).toBe(false);
});

test('HWP original has the same fullscreen page navigation and zoom', async ({ page }) => {
  await page.goto('/'); await page.locator('input[type=file]').setInputFiles(path.resolve('tests/fixtures/hwp/landscape-pagination.hwp'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true }); await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByRole('button', { name: '한글 전체 화면 뷰어', exact: true }).click();
  const viewer = page.locator('.android-presentation'); await expect(viewer).toHaveAttribute('data-page', '1');
  await expect(page.locator('.android-presentation-page svg')).toBeVisible();
  await swipe(page, [330, 400], [70, 400]); await expect(viewer).toHaveAttribute('data-page', '2');
  await reveal(page); await page.getByRole('button', { name: '뷰어 확대', exact: true }).click();
  await expect(viewer).toHaveAttribute('data-page', '2');
  await page.getByRole('button', { name: '뷰어 닫기', exact: true }).click();
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '2쪽');
});


test('touch typing and table tools target the actual cell in the scaled editor', async ({ page }) => {
  await page.goto('/');
  const buffer = await Packer.toBuffer(new Document({ sections: [{ children: [new Table({ rows: [new TableRow({ children: [
    new TableCell({ children: [new Paragraph('LEFT CELL')] }), new TableCell({ children: [new Paragraph('RIGHT CELL')] }),
  ] })] }), new Paragraph('OUTSIDE')] }] }));
  await page.locator('input[type=file]').setInputFiles({ name: 'cells.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await page.getByRole('button', { name: '편집', exact: true }).click();
  const editor = page.getByLabel('Android 문서 편집 영역'), cell = editor.locator('td').filter({ hasText: 'LEFT CELL' });
  await cell.tap(); await page.keyboard.type(' TOUCH_MARKER');
  await expect(cell).toContainText('TOUCH_MARKER');
  await expect(editor.locator('p').filter({ hasText: 'OUTSIDE' })).not.toContainText('TOUCH_MARKER');
  await cell.tap(); await page.getByRole('button', { name: '표 도구', exact: true }).click();
  await expect(page.getByRole('button', { name: '아래에 행 삽입', exact: true })).toBeVisible();
});
