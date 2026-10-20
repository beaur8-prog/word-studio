import { test, expect, type Page } from '@playwright/test';
import path from 'node:path';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
const { createHwpScrollFixture } = createRequire(import.meta.url)('./hwp-scroll-fixture.cjs');
async function currentPageText(page: Page) { return page.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, '')); }

test('wheel, trackpad direction, keyboard and page controls share one continuous HWP scroll position', async ({ page }, testInfo) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/landscape-pagination.hwp'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true }), viewport = reader.getByRole('main', { name: 'HWP 문서 스크롤' });
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '1쪽');
  await viewport.hover(); await page.mouse.wheel(0, 650);
  await expect(reader.getByLabel('HWP 쪽 번호')).toHaveValue('2');
  await expect.poll(() => viewport.evaluate(element => element.scrollTop)).toBeGreaterThan(500);
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '2쪽');
  const before = await viewport.evaluate(element => element.scrollTop);
  await page.mouse.wheel(0, 80);
  await expect.poll(() => viewport.evaluate(element => element.scrollTop)).toBeGreaterThan(before + 50);
  await page.mouse.wheel(0, -1500);
  await expect(reader.getByLabel('HWP 쪽 번호')).toHaveValue('1');
  await viewport.press('End');
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '4쪽');
  await expect.poll(() => currentPageText(page)).toContain('ROW-24-COLUMN-2');
  await viewport.press('Home'); await expect(reader.getByLabel('HWP 쪽 번호')).toHaveValue('1');
  await viewport.press('PageDown'); await expect(reader.getByLabel('HWP 쪽 번호')).toHaveValue('2');
  await viewport.press('PageUp'); await expect(reader.getByLabel('HWP 쪽 번호')).toHaveValue('1');
  await reader.getByLabel('HWP 쪽 번호').fill('3');
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '3쪽');
  await reader.getByLabel('HWP 확대').selectOption('200');
  await expect(reader.getByLabel('HWP 쪽 번호')).toHaveValue('3');
  await viewport.hover(); await page.mouse.wheel(500, 0);
  await expect.poll(() => viewport.evaluate(element => element.scrollLeft)).toBeGreaterThan(100);
  await reader.getByLabel('HWP 확대').selectOption('85');
  await reader.getByLabel('HWP 문서에서 찾기').fill('ROW-24-COLUMN-2'); await reader.getByRole('button', { name: '다음 찾기' }).click();
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '4쪽');
  await page.screenshot({ path: testInfo.outputPath('continuous-scroll.png') });
  expect(await reader.locator('.hwp-sheet svg').count()).toBeLessThanOrEqual(5);
  expect(errors).toEqual([]);
});

test('a 1000-page document can seek, scroll and search while only nearby pages are mounted', async ({ page }, testInfo) => {
  test.setTimeout(120000);
  await fs.mkdir('.cache', { recursive: true });
  const directory = await fs.mkdtemp(path.resolve('.cache/hwp-scroll-regression-'));
  const file = await createHwpScrollFixture(directory);
  await page.goto('/'); await page.locator('input[type=file]').first().setInputFiles(file);
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true }), viewport = reader.getByRole('main', { name: 'HWP 문서 스크롤' });
  await expect(reader.getByTestId('hwp-page-count')).toHaveText('/ 1000');
  await expect.poll(() => currentPageText(page)).toContain('SCROLL-PAGE-0001');
  await reader.getByLabel('HWP 쪽 번호').fill('500');
  await expect.poll(() => currentPageText(page)).toContain('SCROLL-PAGE-0500');
  expect(await reader.locator('.hwp-sheet svg').count()).toBeLessThanOrEqual(5);
  await viewport.hover(); await page.mouse.wheel(0, 1000);
  await expect(reader.getByLabel('HWP 쪽 번호')).toHaveValue('501');
  await viewport.press('End'); await expect.poll(() => currentPageText(page)).toContain('SCROLL-PAGE-1000');
  await reader.getByLabel('HWP 문서에서 찾기').fill('SCROLL-PAGE-0999'); await reader.getByRole('button', { name: '다음 찾기' }).click();
  await expect.poll(() => currentPageText(page), { timeout: 60000 }).toContain('SCROLL-PAGE-0999');
  expect(await reader.locator('.hwp-sheet svg').count()).toBeLessThanOrEqual(5);
  await page.screenshot({ path: testInfo.outputPath('1000-page-scroll.png') });
});

test('queued reader requests dispatch in order and closing rejects waiting work', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const original = window.Worker; let controlled: any;
    class ControlledWorker {
      onmessage: any; onerror: any; messages: any[] = []; terminated = false;
      constructor() { controlled = this; }
      postMessage(value: any) { this.messages.push(value); }
      terminate() { this.terminated = true; }
    }
    window.Worker = ControlledWorker as any;
    try {
      const reader = new HwpReaderSession(), first = reader.pageText(0), obsolete = reader.page(1).catch(error => error.message), second = reader.page(2).catch(error => error.message);
      reader.protectQueuedPage(2); reader.discardQueuedPages(new Set([0]));
      const before = controlled.messages.length;
      controlled.onmessage({ data: { id: controlled.messages[0].id, result: 'first page' } });
      const firstText = await first, after = controlled.messages.length;
      reader.close(); return { before, after, firstText, obsolete: await obsolete, protectedPage: controlled.messages[1].page, second: await second, terminated: controlled.terminated };
    } finally { window.Worker = original; }
  });
  expect(result).toEqual({ before: 1, after: 2, firstText: 'first page', obsolete: '쪽 읽기 작업이 취소됐습니다.', protectedPage: 2, second: '읽기 작업이 취소됐습니다.', terminated: true });
});

test('OPF-relative HWPX sections open, find text and preserve the current Word document', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.goto('/');
  await page.locator('.tiptap').fill('KEEP AFTER RELATIVE HWPX');
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/manifest-relative.hwpx'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page-count')).toHaveText('/ 1');
  await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('첫번째문단:HWPX상대경로검증');
  await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('두번째문단:WordStudio');
  await reader.getByLabel('HWP 문서에서 찾기').fill('두 번째 문단');
  await reader.getByRole('button', { name: '다음 찾기' }).click();
  await expect(reader.getByRole('status').last()).toContainText('1쪽에서 찾았습니다');
  await reader.getByRole('button', { name: 'HWP 읽기 닫기' }).click();
  await expect(page.locator('.tiptap')).toHaveText('KEEP AFTER RELATIVE HWPX');
  expect(errors).toEqual([]);
});

test('landscape table page breaks keep the final row and preserve the Word document', async ({ page }) => {
  await page.goto('/');
  await page.locator('.tiptap').fill('KEEP AFTER LANDSCAPE PAGINATION');
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/landscape-pagination.hwp'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page-count')).toHaveText('/ 4');
  await reader.getByLabel('HWP 쪽 번호').fill('4');
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '4쪽');
  await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('ROW-24-COLUMN-2');
  await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('WordStudio');
  await reader.getByRole('button', { name: 'HWP 읽기 닫기' }).click();
  await expect(page.locator('.tiptap')).toHaveText('KEEP AFTER LANDSCAPE PAGINATION');
});

for (const extension of ['hwp', 'hwpx']) {
  test(`${extension} opens locally with selectable table text and preserves the Word document`, async ({ page }) => {
    const errors: string[] = [], external: string[] = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('request', request => { if (/^https?:/.test(request.url()) && !request.url().startsWith('http://127.0.0.1:5191/')) external.push(request.url()); });
    await page.goto('/');
    await page.locator('.tiptap').fill('보존할 Word 문서 KEEP');
    await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
    await page.locator('input[type=file]').first().setInputFiles(path.resolve(`tests/fixtures/hwp/reader.${extension}`));
    const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
    await expect(reader).toBeVisible();
    await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('한글읽기검증');
    await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('테이블속한글');
    await expect(reader.getByTestId('hwp-page-count')).toHaveText('/ 1');
    await reader.getByLabel('HWP 확대').selectOption('50');
    await expect(reader.getByTestId('hwp-page')).toHaveCSS('transform', 'matrix(0.5, 0, 0, 0.5, 0, 0)');
    await page.keyboard.press('ControlOrMeta+s');
    await expect(reader.getByRole('status').last()).toContainText('읽기 전용');
    await reader.getByLabel('HWP 문서에서 찾기').fill('테이블 속');
    await reader.getByRole('button', { name: '다음 찾기' }).click();
    await expect(reader.getByRole('status').last()).toContainText('1쪽에서 찾았습니다');
    await reader.getByRole('button', { name: 'HWP 읽기 닫기' }).click();
    await expect(reader).not.toBeVisible();
    await expect(page.locator('.tiptap')).toHaveText('보존할 Word 문서 KEEP');
    await page.reload();
    await expect(page.locator('.tiptap')).toHaveText('보존할 Word 문서 KEEP');
    expect(external).toEqual([]); expect(errors).toEqual([]);
  });
}

test('corrupt, empty and obsolete HWP documents fail without replacing the current document', async ({ page }) => {
  await page.goto('/'); await page.locator('.tiptap').fill('KEEP AFTER FAILED IMPORT');
  for (const buffer of [Buffer.from('HWP Document File - v3'), Buffer.alloc(0), Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])]) {
    await page.locator('input[type=file]').first().setInputFiles({ name: 'corrupt.hwp', mimeType: 'application/x-hwp', buffer });
    const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
    await expect(reader.getByRole('alert')).toContainText('HWP 파일을 열지 못했습니다');
    await reader.getByRole('button', { name: '편집 문서로 돌아가기' }).click();
    await expect(page.locator('.tiptap')).toHaveText('KEEP AFTER FAILED IMPORT');
  }
});

test('SVG import preserves clipping and strips scripts and external resources', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { safeHwpSvg } = await import('/src/hwp/reader.ts');
    return safeHwpSvg('<svg xmlns="http://www.w3.org/2000/svg" width="794" height="1122"><defs><clipPath id="cell"><rect width="30" height="20"/></clipPath></defs><g clip-path="url(#cell)"><text onclick="alert(1)">한글</text></g><script>alert(1)</script><image href="https://example.com/tracker.png"/><rect fill="url(https://example.com/paint)"/><foreignObject><iframe src="https://example.com"/></foreignObject></svg>');
  });
  expect(result.width).toBe(794); expect(result.svg).toMatch(/url\(#hwp-[\w-]+\)/); expect(result.svg).toContain('한글');
  expect(result.svg).not.toMatch(/script|onclick|example\.com|foreignObject|iframe/);
});

test('closing during a delayed parse cannot reopen the viewer or change the Word document', async ({ page }) => {
  await page.route('**/rhwp_bg.wasm*', async route => { await new Promise(resolve => setTimeout(resolve, 800)); await route.continue(); });
  await page.goto('/'); await page.locator('.tiptap').fill('KEEP DURING PARSE');
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/reader.hwp'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await reader.getByRole('button', { name: 'HWP 읽기 닫기' }).click();
  await page.waitForTimeout(1000);
  await expect(reader).not.toBeVisible();
  await expect(page.locator('.tiptap')).toHaveText('KEEP DURING PARSE');
});

test('inline pages have independent clipping resources, including a second print copy', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { safeHwpSvg } = await import('/src/hwp/reader.ts');
    const source = '<svg xmlns="http://www.w3.org/2000/svg" width="794" height="1122"><defs><clipPath id="cell"><rect width="40" height="20"/></clipPath></defs><g clip-path="url(#cell)"><text>표</text></g></svg>';
    const first = safeHwpSvg(source), second = safeHwpSvg(source), printCopy = safeHwpSvg(first.svg);
    const wrapper = document.createElement('div'); wrapper.innerHTML = first.svg + second.svg + printCopy.svg;
    const ids = [...wrapper.querySelectorAll('[id]')].map(el => el.id);
    return { unique: new Set(ids).size === ids.length, references: [...wrapper.querySelectorAll('g')].every(el => {
      const id = el.getAttribute('clip-path')?.slice(5, -1);
      return !![...el.closest('svg')!.querySelectorAll('[id]')].find(resource => resource.id === id);
    }) };
  });
  expect(result).toEqual({ unique: true, references: true });
});

test('opening a regular document from HWP does not hide a recovery failure dialog', async ({ page }) => {
  await page.addInitScript(() => {
    (window as any).diskFail = false;
    (window as any).wordDesktop = { platform: 'test', readRecovery: async () => [], setTheme: () => {}, onMenu: () => () => {},
      saveRecovery: async () => { if ((window as any).diskFail) throw Error('disk full'); return true; } };
  });
  await page.goto('/'); await expect(page.locator('.tiptap')).toBeVisible();
  await page.evaluate(() => { (window as any).diskFail = true; });
  await page.locator('.tiptap').fill('KEEP AFTER RECOVERY FAILURE');
  await expect(page.getByText('자동 저장 실패 · 파일로 저장하세요', { exact: true })).toBeVisible();
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/reader.hwp'));
  await expect(page.getByTestId('hwp-page')).toBeVisible();
  await page.locator('input[type=file]').first().setInputFiles({ name: 'another.txt', mimeType: 'text/plain', buffer: Buffer.from('NEW DOCUMENT') });
  await expect(page.getByRole('dialog', { name: 'HWP 읽기', exact: true })).not.toBeVisible();
  const failure = page.getByRole('dialog', { name: '현재 문서를 저장하지 못했습니다' });
  await expect(failure).toBeVisible();
  await failure.getByRole('button', { name: '취소', exact: true }).click();
  await expect(page.locator('.tiptap')).toHaveText('KEEP AFTER RECOVERY FAILURE');
});
