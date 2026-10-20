import { test, expect, type Page } from '@playwright/test';
import path from 'node:path';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
const { createLargeHwpFixture } = createRequire(import.meta.url)('./large-hwp-fixture.cjs');
let largeFixture: Promise<any> | undefined;
async function largeDocument() {
  if (!largeFixture) largeFixture = fs.mkdir('.cache', { recursive: true }).then(() => fs.mkdtemp(path.resolve('.cache/hwp-large-regression-'))).then(createLargeHwpFixture);
  return largeFixture;
}

async function editorImageHashes(page: Page) {
  return page.locator('.tiptap').evaluate(async root => {
    const result: string[] = [];
    for (const image of root.querySelectorAll('img:not(.ProseMirror-separator)')) {
      const bytes = await (await fetch(image.getAttribute('src')!)).arrayBuffer();
      result.push([...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))].map(n => n.toString(16).padStart(2, '0')).join(''));
    }
    return result;
  });
}

async function editFile(page: Page, file: string, timeout = 15000) {
  await page.locator('input[type=file]').first().setInputFiles(file);
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page')).toBeVisible({ timeout });
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  const notice = reader.getByRole('region', { name: '한글 편집 변환 안내' });
  await expect(notice).toBeVisible({ timeout });
  await notice.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  await expect(reader).not.toBeVisible();
  await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(file).replace(/\.(hwp|hwpx)$/i, ''), { timeout });
}
async function prepareExport(page: Page, format: 'hwp' | 'hwpx', timeout = 60000) {
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.locator('.backstage-nav button').filter({ hasText: format === 'hwp' ? /^한글 문서로 내보내기/ : /^한글 표준 문서로 내보내기/ }).click();
  const dialog = page.getByRole('dialog', { name: '한글 문서로 내보내기', exact: true });
  await expect(dialog).toBeVisible({ timeout }); return dialog;
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

  test(`${format}: a 21 MB picture file expands past 24 MiB and still edits, saves and reopens without losing pictures`, async ({ page }, testInfo) => {
    test.setTimeout(180000);
    const fixture = await largeDocument(), file = fixture.files[format];
    expect(file.bytes).toBeGreaterThan(20 * 1000 * 1000);
    expect(file.bytes).toBeLessThan(23 * 1000 * 1000);
    expect(fixture.htmlBytes).toBeGreaterThan(24 * 1024 * 1024);
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.goto('/'); await editFile(page, file.file);
    const editor = page.locator('.tiptap');
    for (let i = 1; i <= 3; i++) await expect(editor).toContainText('LARGE-IMAGE-' + i);
    expect(await editorImageHashes(page)).toEqual(fixture.imageSHA256);
    await editor.locator('p').first().fill('EDITED LARGE DOCUMENT');
    const saved = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
    const savedFile = testInfo.outputPath('large-edited.wordstudio'); await (await saved).saveAs(savedFile);
    const backup = JSON.parse(await fs.readFile(savedFile, 'utf8'));
    backup.title = 'REOPENED LARGE BACKUP'; await fs.writeFile(savedFile, JSON.stringify(backup));
    await page.locator('input[type=file]').first().setInputFiles(savedFile);
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue('REOPENED LARGE BACKUP');
    await expect(editor).toContainText('EDITED LARGE DOCUMENT');
    expect(await editorImageHashes(page)).toEqual(fixture.imageSHA256);
    const dialog = await prepareExport(page, format);
    const output = page.waitForEvent('download'); await dialog.getByRole('button', { name: `${format.toUpperCase()} 파일 저장`, exact: true }).click();
    const outputFile = testInfo.outputPath('large-edited.' + format); await (await output).saveAs(outputFile);
    await editFile(page, outputFile);
    await expect(editor).toContainText('EDITED LARGE DOCUMENT');
    expect(await editorImageHashes(page)).toEqual(fixture.imageSHA256);
    await page.screenshot({ path: testInfo.outputPath('large-edited.png') });
    await fs.writeFile(testInfo.outputPath('large-result.json'), JSON.stringify({ format, sourceBytes: file.bytes, expandedHtmlBytes: fixture.htmlBytes, savedBytes: (await fs.stat(outputFile)).size, images: 3, imageSHA256: fixture.imageSHA256, rendererErrors: errors }, null, 2));
    expect(errors).toEqual([]);
  });
}

test('a 21 MB individual picture is retained and a Word Studio backup above 40 MiB reopens intact', async ({ page }) => {
  test.setTimeout(90000);
  const fixture = await largeDocument();
  const data = 'data:image/png;base64,' + (await fs.readFile(fixture.largeImage)).toString('base64');
  await page.goto('/');
  const result = await page.evaluate(async src => {
    const { sanitizeImportedHtml } = await import('/src/document-html.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { validateDocument, importFile } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const clean = sanitizeImportedHtml(`<p>KEEP LARGE PICTURE</p><p><img src="${src}" width="320" height="875"></p>`);
    const doc = validateDocument(importEditableHwp({ html: clean.html, page: defaultPage, warnings: [] }, 'picture.hwp'));
    // Two copies exceed the former 40 MiB native-file cap, with real PNG data.
    const image = doc.content.content!.find(node => node.type === 'image' || node.content?.some(n => n.type === 'inlineImage'))!;
    if (!image) throw Error('Large picture was discarded');
    doc.content.content!.push(structuredClone(image));
    const file = new File([JSON.stringify(doc)], 'large.wordstudio');
    const reopened = await importFile(file);
    let images = 0, intact = true;
    const walk = (node: any) => { if (node.type === 'image' || node.type === 'inlineImage') { images++; intact &&= node.attrs.src === src; } node.content?.forEach(walk); };
    walk(reopened.document.content);
    let compressedLimit = '';
    try { await importFile(new File([new Uint8Array(129 * 1024 * 1024)], 'too-large.docx')); }
    catch (error) { compressedLimit = String(error); }
    return { bytes: file.size, images, intact, warnings: clean.warnings, compressedLimit };
  }, data);
  expect(result.bytes).toBeGreaterThan(40 * 1024 * 1024);
  expect(result.images).toBe(2); expect(result.intact).toBe(true); expect(result.warnings).toEqual([]);
  expect(result.compressedLimit).toContain('128 MiB');
});

test('a source above 40 MiB expands past 100 MiB, edits and saves without dropping pictures', async ({ page }, testInfo) => {
  test.setTimeout(240000);
  await fs.mkdir('.cache', { recursive: true });
  const fixture = await createLargeHwpFixture(await fs.mkdtemp(path.resolve('.cache/hwp-expanded-regression-')), 4);
  expect(fixture.files.hwp.bytes).toBeGreaterThan(40 * 1024 * 1024);
  expect(fixture.htmlBytes).toBeGreaterThan(100 * 1024 * 1024);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.goto('/'); await editFile(page, fixture.files.hwp.file, 120000);
  const editor = page.locator('.tiptap');
  await expect(editor.locator('img:not(.ProseMirror-separator)')).toHaveCount(12);
  expect(await editorImageHashes(page)).toEqual(fixture.imageSHA256);
  await editor.locator('p').first().fill('EXPANDED DOCUMENT EDITED');
  const backupDownload = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const backup = testInfo.outputPath('expanded.wordstudio'); await (await backupDownload).saveAs(backup);
  expect((await fs.stat(backup)).size).toBeGreaterThan(100 * 1024 * 1024);
  const dialog = await prepareExport(page, 'hwp', 120000);
  const downloaded = page.waitForEvent('download'); await dialog.getByRole('button', { name: 'HWP 파일 저장', exact: true }).click();
  const output = testInfo.outputPath('expanded.hwp'); await (await downloaded).saveAs(output);
  expect((await fs.stat(output)).size).toBeGreaterThan(40 * 1024 * 1024);
  await page.locator('input[type=file]').first().setInputFiles(output);
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page-count')).toHaveText('/ 12');
  await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('EXPANDEDDOCUMENTEDITED');
  await reader.getByLabel('HWP 쪽 번호').fill('12');
  await expect.poll(() => reader.getByTestId('hwp-page').textContent().then(text => text?.replace(/\s/g, ''))).toContain('LARGE-IMAGE-12');
  await page.screenshot({ path: testInfo.outputPath('expanded-reader.png') });
  await fs.writeFile(testInfo.outputPath('expanded-result.json'), JSON.stringify({ sourceBytes: fixture.files.hwp.bytes, htmlBytes: fixture.htmlBytes,
    backupBytes: (await fs.stat(backup)).size, outputBytes: (await fs.stat(output)).size, imageSHA256: fixture.imageSHA256, rendererErrors: errors }, null, 2));
  expect(errors).toEqual([]);
});

test('canceling conversion keeps the current document, and recovery failure does not discard it', async ({ page }) => {
  await page.addInitScript(() => {
    (window as any).diskFail = false;
    (window as any).wordDesktop = { platform: 'test', readRecovery: async () => [], setTheme: () => {}, onMenu: () => () => {},
      saveRecovery: async () => { if ((window as any).diskFail) throw Error('disk full'); return true; } };
  });
  await page.goto('/'); await page.locator('.tiptap').fill('PREVIOUS SAVED DOCUMENT');
  await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
  await page.evaluate(() => { (window as any).diskFail = true; });
  await page.locator('.tiptap').fill('KEEP EDITS BEFORE CONVERSION');
  await expect(page.getByText('자동 저장 실패 · 파일로 저장하세요', { exact: true })).toBeVisible();
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/manifest-relative.hwpx'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  await reader.getByRole('button', { name: '원본 보기로 돌아가기', exact: true }).click();
  await expect(reader.getByRole('region', { name: '한글 편집 변환 안내' })).not.toBeVisible();
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
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
