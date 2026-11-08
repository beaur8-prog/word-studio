import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, ImageRun } from 'docx';
import path from 'node:path';
import fs from 'node:fs/promises';
test.use({ hasTouch: true, viewport: { width: 412, height: 840 }, userAgent: 'Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 Chrome/130.0.0.0 Mobile Safari/537.36' });
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=', 'base64');
async function fixture() {
  return Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('ANDROID FIRST PAGE'), new Table({ rows: [
    new TableRow({ children: [new TableCell({ children: [new Paragraph('Merged title')], columnSpan: 2 })] }),
    new TableRow({ children: [new TableCell({ children: [new Paragraph('Cell A')] }), new TableCell({ children: [new Paragraph('Cell B')] })] }),
  ] }), new Paragraph({ children: [new ImageRun({ type: 'png', data: png, transformation: { width: 50, height: 50 } })] }),
  new Paragraph({ text: 'ANDROID LAST PAGE', pageBreakBefore: true })] }] }));
}
async function start(page: Page) {
  await page.goto('/'); await expect(page.locator('.android-shell')).toBeVisible();
}
async function openFixture(page: Page) {
  await start(page);
  await page.locator('input[type=file]').setInputFiles({ name: 'android.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await fixture() });
  await expect(page.locator('.android-paper-viewport')).toBeVisible();
  await expect(page.getByRole('button', { name: '쪽 1/2', exact: true })).toBeVisible();
}

test('Android startup is empty; desktop, iPad and Electron never opt in because of width or touch', async ({ browser, page }) => {
  await start(page); await expect(page.getByText('문서를 열어보세요', { exact: true })).toBeVisible();
  for (const userAgent of ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0.0.0 Safari/537.36',
    'Mozilla/5.0 (iPad; CPU OS 18_0 like Mac OS X) AppleWebKit/605.1.15 Version/18.0 Mobile/15E148 Safari/604.1']) {
    const context = await browser.newContext({ userAgent, hasTouch: true, viewport: { width: 360, height: 740 } });
    const desktop = await context.newPage(); await desktop.goto(page.url());
    await expect(desktop.locator('.office-shell')).toBeVisible(); await expect(desktop.locator('.android-shell')).toHaveCount(0);
    expect(await desktop.locator('.office-shell').evaluate(el => getComputedStyle(el).minWidth)).toBe('840px');
    await context.close();
  }
  const context = await browser.newContext({ userAgent: 'Android 14' });
  await context.addInitScript(() => { (window as any).wordDesktop = { platform: 'win32', readRecovery: async () => [], onMenu: () => () => {}, setTheme: () => {} }; });
  const desktop = await context.newPage(); await desktop.goto(page.url()); await expect(desktop.locator('.office-shell')).toBeVisible(); await context.close();
});

test('read-only DOCX keeps pages, merged cells and pictures through fit, search, zoom and rotation', async ({ page }, testInfo) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await openFixture(page);
  await expect(page.locator('.android-paper-slot[aria-current=page]')).toContainText('ANDROID FIRST PAGE');
  await expect(page.locator('.android-paper-layer td[colspan="2"]')).toContainText('Merged title');
  await expect(page.locator('.android-paper-layer img')).toHaveCount(1);
  await expect.poll(() => page.locator('.android-paper-layer img').evaluate((img: HTMLImageElement) => img.complete && img.naturalWidth > 0)).toBe(true);
  await expect(page.locator('[contenteditable="true"]')).toHaveCount(0);
  for (const width of [360, 599, 600, 839, 840, 1280, 412]) {
    await page.setViewportSize({ width, height: 840 });
    await expect.poll(() => page.locator('.android-paper-slot').first().evaluate(el => el.getBoundingClientRect().width)).toBeLessThan(width);
    expect(await page.locator('.android-shell').evaluate(el => el.scrollWidth <= el.clientWidth + 1)).toBe(true);
    expect(await page.locator('.android-paper-slot .print-page').first().evaluate(el => el.style.width)).toMatch(/^793/);
  }
  await page.getByRole('button', { name: '찾기', exact: true }).click();
  await page.getByLabel('찾을 내용').fill('ANDROID LAST');
  await expect(page.getByText('1개 쪽에서 찾았습니다')).toBeVisible();
  await page.getByRole('button', { name: '다음 결과로 이동' }).click();
  await expect(page.getByRole('button', { name: '쪽 2/2', exact: true })).toBeVisible();
  await expect(page.locator('.android-paper-slot[aria-current=page]')).toContainText('ANDROID LAST PAGE');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByLabel('문서 확대', { exact: true }).fill('150');
  await page.getByRole('button', { name: '보기 설정 닫기' }).click();
  await expect(page.getByRole('button', { name: '쪽 2/2', exact: true })).toBeVisible();
  await page.screenshot({ path: testInfo.outputPath('android-reader.png') });
  expect(errors).toEqual([]);
});

test('Android toolbar and sheets provide 48px targets without changing paper controls', async ({ page }) => {
  await openFixture(page); await page.getByRole('button', { name: '문서 메뉴', exact: true }).click();
  const small = await page.locator('.android-topbar button, .android-reader-tools button, .android-sheet[open] button').evaluateAll(elements => elements.filter(el => {
    const r = el.getBoundingClientRect(); return r.width < 47.5 || r.height < 47.5;
  }).map(el => ({ text: el.textContent, width: el.getBoundingClientRect().width, height: el.getBoundingClientRect().height })));
  expect(small).toEqual([]);
  await page.getByRole('button', { name: '문서 메뉴 닫기' }).click();
});

test('basic editing, table commands, undo/redo and two DOCX saves preserve content and pictures', async ({ page }) => {
  await openFixture(page);
  await page.getByRole('button', { name: '편집', exact: true }).click();
  const editor = page.getByLabel('Android 문서 편집 영역');
  await expect(editor).toHaveAttribute('contenteditable', 'true');
  await editor.click(); await page.keyboard.press('ControlOrMeta+End'); await page.keyboard.type(' MOBILE EDIT MARKER');
  await page.getByRole('button', { name: '실행 취소', exact: true }).click();
  await expect(editor).not.toContainText('MOBILE EDIT MARKER');
  await page.getByRole('button', { name: '다시 실행', exact: true }).click(); await expect(editor).toContainText('MOBILE EDIT MARKER');
  await page.getByRole('button', { name: '편집 완료', exact: true }).click();
  await expect(page.locator('.android-paper-viewport')).toBeVisible();
  for (let round = 0; round < 2; round++) {
    await page.getByRole('button', { name: '공유·저장', exact: true }).click();
    const download = page.waitForEvent('download'); await page.getByRole('button', { name: 'Word 문서 사본 저장 (.docx)', exact: true }).click();
    const file = await download; expect(file.suggestedFilename()).toBe('android.docx');
    const bytes = await fs.readFile((await file.path())!);
    await page.getByRole('button', { name: '공유·사본 저장 닫기' }).click();
    // A browser download is not acknowledged as a native write; keep a recovery copy before switching.
    await page.locator('input[type=file]').setInputFiles({ name: 'android.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
    if (await page.getByRole('dialog', { name: '변경 내용 보관' }).count()) await page.getByRole('button', { name: '복구본 보관 후 계속' }).click();
    await expect(page.locator('.android-paper-viewport')).toBeVisible();
    await expect(page.locator('.android-paper-layer')).toContainText('MOBILE EDIT MARKER');
    await expect(page.locator('.android-paper-layer td[colspan="2"]')).toContainText('Merged title');
    await expect(page.locator('.android-paper-layer img')).toHaveCount(1);
  }
});

test('native picker cancellation and error release the request; external intent opens the requested file', async ({ page }) => {
  await page.addInitScript(() => {
    let calls = 0;
    (window as any).wordAndroid = { openFile: async () => { calls++; (window as any).pickerCalls = calls; if (calls === 1) return null; if (calls === 2) throw new Error('Provider unavailable'); return { name: 'native.txt', data: [78, 65, 84, 73, 86, 69] }; },
      onOpenFile: (callback: any) => { (window as any).incomingFile = callback; return () => {}; } };
  });
  await start(page);
  await page.getByRole('button', { name: '파일 열기', exact: true }).click(); await expect(page.locator('.android-operation')).toHaveCount(0);
  await page.getByRole('button', { name: '파일 열기', exact: true }).click(); await expect(page.getByText('Provider unavailable', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: '파일 열기', exact: true }).click(); await expect(page.locator('.android-paper-layer')).toContainText('NATIVE');
  expect(await page.evaluate(() => (window as any).pickerCalls)).toBe(3);
  await page.evaluate(() => (window as any).incomingFile({ name: 'external.txt', data: [69, 88, 84, 69, 82, 78, 65, 76] }));
  await expect(page.locator('.android-paper-layer')).toContainText('EXTERNAL');
});

test('HWP original viewer fits the phone and leaves conversion explicit', async ({ page }) => {
  await start(page); await page.locator('input[type=file]').setInputFiles(path.resolve('tests/fixtures/hwp/landscape-pagination.hwp'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader).toHaveAttribute('data-android', 'true'); await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await expect(reader.locator('.hwp-toolbar')).toHaveCount(0);
  await reader.press('Control+f'); await expect(page.getByRole('dialog', { name: '한글 문서에서 찾기', exact: true })).toBeVisible();
  await page.getByRole('button', { name: '한글 문서에서 찾기 닫기', exact: true }).click();
  await expect.poll(() => reader.locator('.hwp-sheet').first().evaluate(el => el.getBoundingClientRect().width)).toBeLessThan(412);
  await reader.getByRole('button', { name: /^쪽 1\// }).click();
  await page.getByLabel('HWP 쪽 번호').fill('4'); await page.getByRole('button', { name: '원본으로 돌아가기' }).click();
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '4쪽');
  await reader.getByRole('button', { name: '한글 원본 보기 닫기' }).click(); await expect(reader).toHaveCount(0);
  await expect(page.getByText('문서를 열어보세요', { exact: true })).toBeVisible();
});

test('table operations and numerical row/column sizing are touch accessible and undoable', async ({ page }) => {
  await openFixture(page); await page.getByRole('button', { name: '편집', exact: true }).click();
  const edit = page.getByLabel('Android 문서 편집 영역');
  await edit.locator('td').filter({ hasText: 'Cell A' }).tap();
  await page.getByRole('button', { name: '표 도구', exact: true }).click();
  const table = edit.locator('table');
  await page.getByRole('button', { name: '아래에 행 삽입', exact: true }).click();
  await expect(table.locator('tr')).toHaveCount(3);
  await page.getByLabel('선택 행 최소 높이 (mm)').fill('18');
  await page.getByRole('button', { name: '행 높이 적용', exact: true }).click();
  await expect.poll(() => table.locator('tr').evaluateAll(rows => rows.some(row => parseFloat(row.style.height) > 67))).toBe(true);
  await page.getByLabel('선택 열 너비 (mm)').fill('40');
  await page.getByRole('button', { name: '열 너비 적용', exact: true }).click();
  await expect.poll(() => table.locator('[colwidth]').evaluateAll(cells => cells.some(cell => cell.getAttribute('colwidth')?.includes('151')))).toBe(true);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  await page.getByRole('button', { name: '실행 취소', exact: true }).click();
  await expect.poll(() => table.locator('[colwidth]').evaluateAll(cells => cells.some(cell => cell.getAttribute('colwidth')?.includes('151')))).toBe(false);
  await page.getByRole('button', { name: '다시 실행', exact: true }).click();
  await expect(table).toContainText('Cell A'); await expect(table).toContainText('Cell B');
  await page.getByRole('button', { name: '편집 완료', exact: true }).click(); await expect(page.locator('.android-paper-layer td[colspan="2"]')).toContainText('Merged title');
});

test('file cancel, invalid input and native save cancel/error never claim success or lose edited text', async ({ page }) => {
  await page.addInitScript(() => {
    let calls = 0;
    (window as any).wordAndroid = { saveFile: async () => { calls++; if (calls === 1) return 'cancelled'; if (calls === 2) throw new Error('Write failed'); return new Promise(resolve => { (window as any).completeSave = () => resolve('completed'); }); } };
  });
  await start(page);
  await page.locator('input[type=file]').dispatchEvent('cancel');
  await expect(page.getByText('파일 선택을 취소했습니다', { exact: true })).toBeVisible();
  await page.locator('input[type=file]').setInputFiles({ name: 'broken.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from('broken zip') });
  await expect(page.locator('.android-operation')).toHaveCount(0); await expect(page.getByText('문서를 열어보세요', { exact: true })).toBeVisible();
  await page.locator('input[type=file]').setInputFiles({ name: 'safe.txt', mimeType: 'text/plain', buffer: Buffer.from('PRESERVE THIS') });
  await expect(page.locator('.android-paper-layer')).toContainText('PRESERVE THIS');
  await page.getByRole('button', { name: '편집', exact: true }).click();
  const edit = page.getByLabel('Android 문서 편집 영역'); await edit.click(); await page.keyboard.press('ControlOrMeta+End'); await page.keyboard.type(' UNSAVED');
  await page.getByRole('button', { name: '공유·저장', exact: true }).click();
  const save = page.getByRole('button', { name: 'Word 문서 사본 저장 (.docx)', exact: true });
  await save.click(); await expect(page.getByText('저장을 취소했습니다', { exact: true })).toBeVisible();
  await save.click(); await expect(page.getByText('Write failed', { exact: true })).toBeVisible();
  await expect(edit).toContainText('UNSAVED');
  await save.click();
  await expect(page.locator('.android-operation')).toContainText('시스템의 사본 저장 작업');
  await expect(page.locator('.android-operation').getByRole('button', { name: '취소', exact: true })).toHaveCount(0);
  await expect.poll(() => page.evaluate(() => typeof (window as any).completeSave)).toBe('function');
  await page.evaluate(() => (window as any).completeSave());
  await expect(page.getByText('Word 문서 사본 저장 완료', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: '공유·사본 저장 닫기' }).click();
  await page.reload(); await expect(page.getByText('문서를 열어보세요', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: '최근 문서·복구본', exact: true }).click();
  await page.getByRole('button', { name: /^safe ·/ }).click(); await expect(page.locator('.android-paper-layer')).toContainText('UNSAVED');
});

test('long DOCX mounts nearby pages and keeps source geometry while rotating and jumping', async ({ page }) => {
  await start(page);
  const buffer = await Packer.toBuffer(new Document({ sections: [{ children: Array.from({ length: 100 }, (_, i) => new Paragraph({ text: `LONG PAGE ${i + 1}`, pageBreakBefore: i > 0 })) }] }));
  await page.locator('input[type=file]').setInputFiles({ name: 'long.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await expect(page.getByRole('button', { name: '쪽 1/100', exact: true })).toBeVisible();
  expect(await page.locator('.android-paper-slot .print-page').count()).toBeLessThan(7);
  await page.getByRole('button', { name: '쪽 1/100', exact: true }).click(); await page.getByLabel('쪽 번호', { exact: true }).fill('100');
  await page.getByRole('button', { name: '문서로 돌아가기', exact: true }).click();
  await expect(page.locator('.android-paper-slot[aria-current=page]')).toContainText('LONG PAGE 100');
  await page.setViewportSize({ width: 900, height: 412 });
  await expect(page.getByRole('button', { name: '쪽 100/100', exact: true })).toBeVisible();
  expect(await page.locator('.android-paper-slot .print-page').count()).toBeLessThan(7);
  await page.getByRole('button', { name: '쪽 100/100', exact: true }).click(); await page.getByLabel('쪽 번호', { exact: true }).fill('1');
  await page.getByRole('button', { name: '문서로 돌아가기', exact: true }).click();
  await expect(page.locator('.android-paper-slot[aria-current=page]')).toContainText('LONG PAGE 1');
});

test('original sharing keeps exact input bytes, and edited sharing uses a separate prepared DOCX', async ({ page }) => {
  await page.addInitScript(() => { (window as any).wordAndroid = { shareFile: async (file: any) => { (window as any).shared = { name: file.name, bytes: Array.from(file.data) }; (window as any).shareCount = ((window as any).shareCount || 0) + 1; return 'completed'; } }; });
  await start(page); const original = await fixture();
  await page.locator('input[type=file]').setInputFiles({ name: 'original.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: original });
  await expect(page.locator('.android-paper-viewport')).toBeVisible();
  await page.getByRole('button', { name: '편집', exact: true }).click();
  const edit = page.getByLabel('Android 문서 편집 영역'); await edit.click(); await page.keyboard.press('ControlOrMeta+End'); await page.keyboard.type(' SHARED EDIT');
  await page.getByRole('button', { name: '공유·저장', exact: true }).click();
  await page.getByRole('button', { name: '원본 파일 공유', exact: true }).click();
  await expect.poll(() => page.evaluate(() => (window as any).shareCount)).toBe(1);
  const shared = await page.evaluate(() => (window as any).shared);
  expect(Buffer.from(shared.bytes)).toEqual(original);
  await page.getByRole('button', { name: 'Word 편집본 사본 준비', exact: true }).click();
  await expect(page.getByRole('button', { name: '준비된 사본 공유', exact: true })).toBeVisible();
  await page.getByRole('button', { name: '준비된 사본 공유', exact: true }).click();
  await expect.poll(() => page.evaluate(() => (window as any).shareCount)).toBe(2);
  const edited = await page.evaluate(() => (window as any).shared);
  const { default: JSZip } = await import('jszip'); const zip = await JSZip.loadAsync(Buffer.from(edited.bytes));
  expect(await zip.file('word/document.xml')!.async('string')).toContain('SHARED EDIT');
  expect(Buffer.from(edited.bytes).equals(original)).toBe(false);
});

test('same DOCX uses the same fixed paper layout as desktop print preview', async ({ page, browser }) => {
  const buffer = await fixture(); await start(page);
  await page.locator('input[type=file]').setInputFiles({ name: 'same.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await expect(page.locator('.android-paper-layer')).toBeVisible();
  const signature = async (target: Page, selector: string) => target.locator(selector).evaluateAll(sheets => sheets.map(sheet => ({
    width: (sheet as HTMLElement).style.width, height: (sheet as HTMLElement).style.height,
    text: sheet.textContent, cells: [...sheet.querySelectorAll('td,th')].map(cell => ({ text: cell.textContent, cols: cell.getAttribute('colspan'), rows: cell.getAttribute('rowspan') })),
    images: [...sheet.querySelectorAll('img')].map(image => ({ src: image.getAttribute('src'), width: image.getAttribute('width'), height: image.getAttribute('height') })),
  })));
  const android = await signature(page, '.android-paper-slot .print-page');
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0.0.0 Safari/537.36' });
  const desktop = await context.newPage(); const loaded: string[] = []; desktop.on('request', req => loaded.push(req.url()));
  await desktop.goto(page.url()); await desktop.locator('.tiptap').waitFor();
  await desktop.locator('input[type=file]').first().setInputFiles({ name: 'same.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await expect(desktop.locator('.tiptap')).toContainText('ANDROID FIRST PAGE');
  await desktop.getByRole('button', { name: '보기', exact: true }).click(); await desktop.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(desktop.locator('.print-layout.preview')).toBeVisible();
  expect(await signature(desktop, '.print-layout.preview > .print-page')).toEqual(android);
  expect(loaded.some(url => url.includes('/android/AndroidApp') || url.includes('/android/android.css'))).toBe(false);
  await context.close();
});

test('HWP editing remains an explicit converted copy and can return to original reading', async ({ page }) => {
  await start(page); await page.locator('input[type=file]').setInputFiles(path.resolve('tests/fixtures/hwp/reader.hwpx'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true }); await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByRole('button', { name: '편집본', exact: true }).click();
  const notice = reader.getByRole('region', { name: '한글 편집 변환 안내' }); await expect(notice).toBeVisible();
  await notice.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  await expect(page.getByLabel('Android 문서 편집 영역')).toHaveAttribute('contenteditable', 'true');
  await expect(reader).toHaveCount(0);
  await page.getByRole('button', { name: '편집 완료', exact: true }).click(); await expect(page.locator('.android-paper-layer')).toBeVisible();
  await page.getByRole('button', { name: '문서 메뉴', exact: true }).click(); await page.getByRole('button', { name: '한글 원본 보기', exact: true }).click();
  await expect(reader.getByTestId('hwp-page')).toBeVisible();
});

test('external keyboard search and hardware Back use Android panels, and IME composition locks formatting', async ({ page }) => {
  await page.addInitScript(() => { (window as any).wordAndroid = { onBack: (back: any) => { (window as any).hardwareBack = back; return () => {}; } }; });
  await openFixture(page);
  await page.keyboard.press('Control+f'); await expect(page.getByRole('dialog', { name: '문서에서 찾기', exact: true })).toBeVisible();
  await page.getByLabel('찾을 내용').fill('ANDROID');
  await page.evaluate(() => (window as any).hardwareBack());
  await expect(page.getByRole('dialog')).toHaveCount(0); await expect(page.locator('.android-paper-viewport')).toBeVisible();
  await page.getByRole('button', { name: '편집', exact: true }).click();
  const editor = page.getByLabel('Android 문서 편집 영역');
  await editor.dispatchEvent('compositionstart', { data: '한' });
  await expect(page.getByRole('button', { name: '굵게', exact: true })).toBeDisabled();
  await editor.dispatchEvent('compositionend', { data: '한' });
  await expect(page.getByRole('button', { name: '굵게', exact: true })).toBeEnabled();
  await page.keyboard.press('Control+f'); await expect(page.getByRole('dialog', { name: '문서에서 찾기', exact: true })).toBeVisible();
  await page.getByRole('button', { name: '문서에서 찾기 닫기', exact: true }).click(); await expect(page.locator('[contenteditable=true]')).toHaveCount(0);
});
