'use strict';

const { _electron: electron, expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const path = require('node:path');

async function main() {
  const inventory = JSON.parse(await fs.readFile(process.env.WINDOWS_PUBLIC_QA_INVENTORY, 'utf8'));
  const directory = await fs.mkdtemp(path.join(process.env.RUNNER_TEMP, 'studio-native-originals-'));
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  delete env.WORD_STUDIO_DEV;
  let failures = 0;
  let measured = 0;

  for (const item of inventory.filter(row => row.verified)) {
    const pdf = path.join(directory, item.id + '.pdf');
    let stage = 'launch';
    let app;
    let close;
    try {
      app = await electron.launch({
        executablePath: path.resolve(process.env.WORD_STUDIO_EXECUTABLE),
        args: ['--user-data-dir=' + path.join(directory, item.id)],
        env,
      });
      close = await require('../native-smoke-lifetime.cjs')(app, 'Public original ' + item.id, 600000);
      const page = await app.firstWindow();
      await page.setViewportSize({ width: 1440, height: 1000 });
      await page.locator('.tiptap').first().waitFor({ state: 'visible' });
      await app.evaluate(({ dialog }, files) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [files.source] });
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: files.pdf });
      }, { source: item.path, pdf });
      stage = 'open';
      await page.keyboard.press('ControlOrMeta+o');
      let readerPages = null;
      if (item.ext !== 'docx') {
        const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
        await reader.getByTestId('hwp-page').waitFor({ state: 'visible', timeout: 75000 });
        readerPages = Number((await reader.getByTestId('hwp-page-count').textContent()).replace(/\D/g, ''));
        if (!Number.isSafeInteger(readerPages) || readerPages < 1) throw new Error('Reader page count unavailable');
        stage = 'convert-analysis';
        const conversionStarted = Date.now();
        await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
        const panel = reader.getByRole('region', { name: '한글 편집 변환 안내' });
        await panel.waitFor({ state: 'visible', timeout: 240000 });
        stage = 'convert-editor';
        await panel.getByRole('button', { name: '편집본 만들기', exact: true }).click({ timeout: 90000 });
        await reader.waitFor({ state: 'hidden', timeout: 120000 });
        console.log('WINDOWS_NATIVE_CONVERT ' + item.id + ' ' + (Date.now() - conversionStarted));
      }
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true }))
        .toHaveValue(item.id, { timeout: 60000 });
      await page.locator('.busy-indicator').waitFor({ state: 'hidden', timeout: 90000 });
      await page.evaluate(() => document.fonts.ready);
      const stack = page.locator('.document-stack');
      await expect(stack).toHaveAttribute('data-pages', /^[1-9]\d*$/, { timeout: 60000 });
      // Use a frame matching the preview opening, after editor pagination settles.
      await page.getByRole('button', { name: '보기', exact: true }).click();
      const editorPages = Number(await stack.getAttribute('data-pages'));
      stage = 'preview-command';
      const previewStarted = Date.now();
      await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click({ timeout: 75000 });
      stage = 'preview-layout';
      const toolbar = page.getByRole('toolbar', { name: '인쇄 미리보기', exact: true });
      await toolbar.waitFor({ state: 'visible', timeout: 120000 });
      const previewPages = await page.locator('.print-layout.preview .print-page').count();
      const previewMilliseconds = Date.now() - previewStarted;
      stage = 'pdf';
      const pdfStarted = Date.now();
      await toolbar.getByRole('button', { name: 'PDF 저장', exact: true }).click({ timeout: 75000 });
      await expect(page.getByText('PDF를 저장했습니다', { exact: true })).toBeVisible({ timeout: 90000 });
      const bytes = await fs.readFile(pdf);
      const pdfPages = (bytes.toString('latin1').match(/\/Type\s*\/Page\b/g) || []).length;
      if (!pdfPages) throw new Error('Native PDF page count unavailable');
      const metrics = { id: item.id, readerPages, editorPages, previewPages, pdfPages,
        previewMilliseconds, pdfMilliseconds: Date.now() - pdfStarted };
      console.log('WINDOWS_NATIVE_ORIGINAL ' + JSON.stringify(metrics));
      measured++;
    } catch (error) {
      // Playwright errors can include source content. Emit only fixed stage labels.
      const reason = error?.name === 'TimeoutError' ? 'timeout' : 'error';
      console.log('WINDOWS_NATIVE_FAILURE ' + item.id + ' ' + stage + ' ' + reason);
      failures++;
    } finally {
      if (close) await close().catch(() => { failures++; });
      else if (app?.process().exitCode === null) app.process().kill();
    }
  }
  console.log('WINDOWS_NATIVE_TOTAL expected ' + inventory.length + ' measured ' + measured + ' failed ' + failures);
  if (failures || measured !== inventory.length) process.exitCode = 1;
}

main().catch(() => { console.error('Native original check failed outside a document stage'); process.exitCode = 1; });
