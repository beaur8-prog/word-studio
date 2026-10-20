import { test, expect, type Page } from '@playwright/test';
import path from 'node:path';
import fs from 'node:fs/promises';

async function editFile(page: Page, file: string) {
  await page.locator('input[type=file]').first().setInputFiles(file);
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  const notice = reader.getByRole('region', { name: '한글 편집 변환 안내' });
  await expect(notice).toBeVisible();
  await notice.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  await expect(reader).not.toBeVisible();
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(file).replace(/\.(hwp|hwpx)$/i, ''));
}
async function prepareExport(page: Page, format: 'hwp' | 'hwpx') {
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.locator('.backstage-nav button').filter({ hasText: format === 'hwp' ? /^한글 문서로 내보내기/ : /^한글 표준 문서로 내보내기/ }).click();
  const dialog = page.getByRole('dialog', { name: '한글 문서로 내보내기', exact: true });
  await expect(dialog).toBeVisible({ timeout: 60000 }); return dialog;
}

for (const format of ['hwp', 'hwpx'] as const) {
  test(`${format}: edit in the existing Word UI, save as Hancom, reopen, convert to DOCX and back`, async ({ page }, testInfo) => {
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.goto('/');
    await editFile(page, path.resolve(`tests/fixtures/hwp/reader.${format}`));
    const editor = page.locator('.tiptap');
    await expect(editor.locator('table')).toHaveCount(1);
    await expect(editor).toContainText('테이블 속 한글');
    await expect(editor.locator('td p').first()).not.toHaveText(/^\s*\n/);
    const cell = editor.locator('td').filter({ hasText: '테이블 속 한글' }).first();
    await cell.locator('p').first().fill('편집한 표 EDITED 2026');
    const before = await editor.innerText();
    // Original view remains available and never replaces the edited copy.
    await page.getByRole('button', { name: '파일', exact: true }).click();
    await page.getByRole('button', { name: '한글 원본 보기', exact: true }).click();
    await expect(page.getByTestId('hwp-page')).toBeVisible();
    await page.getByRole('button', { name: 'HWP 읽기 닫기' }).click();
    expect(await editor.innerText()).toBe(before);
    const dialog = await prepareExport(page, format);
    await expect(dialog).toContainText('표의 병합과 내용');
    const downloaded = page.waitForEvent('download');
    await dialog.getByRole('button', { name: `${format.toUpperCase()} 파일 저장`, exact: true }).click();
    const download = await downloaded, output = testInfo.outputPath(`edited.${format}`);
    await download.saveAs(output);
    const bytes = await fs.readFile(output);
    expect(bytes.subarray(0, format === 'hwp' ? 8 : 2).toString('hex')).toBe(format === 'hwp' ? 'd0cf11e0a1b11ae1' : '504b');
    await editFile(page, output);
    await expect(editor).toContainText('편집한 표 EDITED 2026');
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const docxDownload = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
    const docx = await docxDownload, docxPath = testInfo.outputPath('converted.docx'); await docx.saveAs(docxPath);
    await page.locator('input[type=file]').first().setInputFiles(docxPath);
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue('converted');
    await expect(editor).toContainText('편집한 표 EDITED 2026');
    await expect(editor.locator('table')).toHaveCount(1);
    const back = await prepareExport(page, format);
    await back.getByRole('button', { name: '취소', exact: true }).click();
    await expect(editor).toContainText('편집한 표 EDITED 2026');
    expect(errors).toEqual([]);
  });
}

test('canceling conversion keeps the current document, and recovery failure does not discard it', async ({ page }) => {
  await page.addInitScript(() => {
    (window as any).diskFail = false;
    (window as any).wordDesktop = { platform: 'test', readRecovery: async () => [], setTheme: () => {}, onMenu: () => () => {},
      saveRecovery: async () => { if ((window as any).diskFail) throw Error('disk full'); return true; } };
  });
  await page.goto('/'); await page.locator('.tiptap').fill('KEEP EDITS BEFORE CONVERSION');
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/manifest-relative.hwpx'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  await reader.getByRole('button', { name: '원본 보기로 돌아가기', exact: true }).click();
  await expect(reader.getByRole('region', { name: '한글 편집 변환 안내' })).not.toBeVisible();
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  await page.evaluate(() => { (window as any).diskFail = true; });
  await reader.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  const failure = page.getByRole('dialog', { name: '현재 문서를 저장하지 못했습니다' });
  await expect(failure).toBeVisible();
  await failure.getByRole('button', { name: '취소', exact: true }).click();
  await expect(page.locator('.tiptap')).toHaveText('KEEP EDITS BEFORE CONVERSION');
});

test('native Hancom save cancellation and failure keep the immutable export and edited document', async ({ page }) => {
  await page.addInitScript(() => {
    (window as any).saveMode = 'cancel'; (window as any).writes = [];
    (window as any).wordDesktop = { platform: 'test', readRecovery: async () => [], setTheme: () => {}, onMenu: () => () => {}, saveRecovery: async () => true,
      saveFile: async (file: any) => { (window as any).writes.push({ name: file.name, extension: file.extension, magic: [...file.data.slice(0, 8)] });
        if ((window as any).saveMode === 'fail') throw Error('disk full'); return (window as any).saveMode === 'save'; } };
  });
  await page.goto('/'); await page.locator('.tiptap').fill('NATIVE 한글 SAVE');
  const dialog = await prepareExport(page, 'hwp');
  await dialog.getByRole('button', { name: 'HWP 파일 저장', exact: true }).click();
  await expect(dialog).toBeVisible();
  await page.evaluate(() => { (window as any).saveMode = 'fail'; });
  await dialog.getByRole('button', { name: 'HWP 파일 저장', exact: true }).click();
  await expect(dialog.getByRole('alert')).toContainText('disk full');
  await page.evaluate(() => { (window as any).saveMode = 'save'; });
  await dialog.getByRole('button', { name: 'HWP 파일 저장', exact: true }).click();
  await expect(dialog).not.toBeVisible();
  await expect(page.locator('.tiptap')).toHaveText('NATIVE 한글 SAVE');
  const writes = await page.evaluate(() => (window as any).writes);
  expect(writes).toHaveLength(3); expect(writes[0]).toEqual(writes[2]);
  expect(writes[0].extension).toBe('hwp'); expect(writes[0].magic).toEqual([208,207,17,224,161,177,26,225]);
});

test('DOCX to Hancom bridge preserves heading appearance, image, merged cells, row sizes and explicit page break', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const doc: any = { format: 'word-studio', version: 1, title: '양방향 변환', comments: [], updatedAt: new Date().toISOString(),
      page: { size: 'Letter', orientation: 'landscape', margin: 20, header: '', footer: '', pageNumbers: false },
      content: parseDocumentHtml(`<h2>서식 제목</h2><p style="text-align:center"><strong>굵게 😀</strong> BEFORE<img data-inline-image="true" src="data:image/png;base64,${png}" width="24" height="24">AFTER</p><table><tr style="height:60px"><td colspan="2" colwidth="120,180"><p>병합 셀</p></td></tr><tr style="height:70px"><td colwidth="120"><p>왼쪽</p></td><td colwidth="180"><p>오른쪽</p></td></tr></table><div data-page-break="true"></div><p>둘째 쪽</p>`) };
    const imported = await importFile(new File([await exportDocx(doc)], 'source.docx'));
    const source = imported.html ? { ...imported.document, content: parseDocumentHtml(imported.html) } : imported.document;
    const files: any[] = [];
    for (const format of ['hwp', 'hwpx'] as const) {
      const exported = await exportEditableHwp(source, format);
      const session = new HwpReaderSession();
      try {
        const info = await session.open(new File([exported.blob], 'converted.' + format));
        const back = importEditableHwp(await session.editable(), 'back.' + format);
        files.push({ format, pages: info.count, html: generateHTML(back.content, editorExtensions), page: back.page, warnings: exported.warnings });
      } finally { session.close(); }
    }
    return files;
  });
  for (const file of result) {
    expect(file.pages).toBe(2);
    for (const text of ['서식 제목', '굵게 😀', 'BEFORE', 'AFTER', '병합 셀', '왼쪽', '오른쪽', '둘째 쪽']) expect(file.html).toContain(text);
    expect(file.html).toMatch(/<strong>/); expect(file.html).toMatch(/font-size: 13pt/);
    expect(file.html).not.toContain('&amp;quot');
    expect(file.html).toMatch(/data:image\/png;base64/); expect(file.html).toMatch(/colspan="2"/);
    expect(file.html).toMatch(/colwidth="120,180"/); expect(file.html).toMatch(/height: (?:60|70)px/);
    expect(file.html.indexOf('BEFORE')).toBeLessThan(file.html.indexOf('<img'));
    expect(file.html.indexOf('<img')).toBeLessThan(file.html.indexOf('AFTER'));
    expect(file.page.size).toBe('Letter'); expect(file.page.orientation).toBe('landscape');
    expect(file.warnings.some((warning: string) => warning.includes('쪽 수 비교'))).toBe(true);
  }
});

if (process.env.HWP_EDITING_PUBLIC_SAMPLES === '1') {
  test('public sample editing previews disclose real Word pagination and retain tables and images', async ({ page }, testInfo) => {
    test.setTimeout(240000);
    await page.goto('/');
    const reports: any[] = [];
    for (const name of ['01-employment-contract.hwp', '02-work-rules.hwp', '03-health-duties.hwp', '04-language-plan.hwp', '05-complex-table.hwp', '06-image.hwp', '07-image.hwpx']) {
      await page.locator('input[type=file]').first().setInputFiles(path.resolve('analysis/hwp-validation/samples', name));
      const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
      await expect(reader.getByTestId('hwp-page')).toBeVisible();
      await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
      const panel = reader.getByRole('region', { name: '한글 편집 변환 안내' });
      await expect(panel).toBeVisible({ timeout: 60000 });
      const notice = await panel.innerText();
      await page.screenshot({ path: testInfo.outputPath(name + '.conversion.png') });
      await panel.getByRole('button', { name: '편집본 만들기', exact: true }).click();
      const editor = page.locator('.tiptap');
      const count = { tables: await editor.locator('table').count(), images: await editor.locator('img:not(.ProseMirror-separator)').count(), text: (await editor.innerText()).length };
      await page.screenshot({ path: testInfo.outputPath(name + '.png') });
      const exported = await prepareExport(page, 'hwp');
      await page.screenshot({ path: testInfo.outputPath(name + '.export.png') });
      reports.push({ name, notice, ...count, exportNotice: await exported.innerText() });
      await exported.getByRole('button', { name: '취소', exact: true }).click();
    }
    await fs.writeFile(testInfo.outputPath('corpus-browser.json'), JSON.stringify(reports, null, 2));
  });
}
