'use strict';

const { _electron: electron, expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');

async function main() {
  const inventory = JSON.parse(await fs.readFile(process.env.WINDOWS_PUBLIC_QA_INVENTORY, 'utf8'));
  const directory = await fs.mkdtemp(path.join(process.env.RUNNER_TEMP || os.tmpdir(), 'studio-native-preview-'));
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
      close = await require('../native-smoke-lifetime.cjs')(app, 'Public original ' + item.id, 1050000);
      const page = await app.firstWindow();
      await page.setViewportSize({ width: 1440, height: 1000 });
      await page.evaluate(() => {
        const tasks = [];
        window.qaLongTasks = tasks;
        new PerformanceObserver(list => {
          for (const task of list.getEntries()) tasks.push({ at: task.startTime, duration: task.duration });
        }).observe({ type: 'longtask' });
      });
      await page.locator('.tiptap').first().waitFor({ state: 'visible' });
      await app.evaluate(({ dialog }, files) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [files.source] });
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: files.pdf });
      }, { source: item.path, pdf });
      stage = 'open';
      await page.keyboard.press('ControlOrMeta+o');
      let readerPages = null, readyStarted = Date.now(), rendererStarted = 0;
      if (item.ext !== 'docx') {
        const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
        await reader.getByTestId('hwp-page').waitFor({ state: 'visible', timeout: 75000 });
        readerPages = Number((await reader.getByTestId('hwp-page-count').textContent()).replace(/\D/g, ''));
        if (!Number.isSafeInteger(readerPages) || readerPages < 1) throw new Error('Reader page count unavailable');
        console.log('WINDOWS_NATIVE_READER ' + item.id + ' ' + readerPages);
        stage = 'convert-analysis';
        const conversionStarted = Date.now();
        await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
        const panel = reader.getByRole('region', { name: '한글 편집 변환 안내' });
        await panel.waitFor({ state: 'visible', timeout: 240000 });
        const comparison = await panel.locator('strong').first().textContent();
        const previewEstimate = comparison?.match(/Word 인쇄 (\d+)쪽/);
        console.log('WINDOWS_NATIVE_ESTIMATE ' + item.id + ' ' + (previewEstimate?.[1] ?? 'unavailable'));
        stage = 'convert-editor-command';
        readyStarted = Date.now();
        rendererStarted = await page.evaluate(() => performance.now());
        await panel.getByRole('button', { name: '편집본 만들기', exact: true }).click({ timeout: 90000 });
        stage = 'convert-editor-view';
        await reader.waitFor({ state: 'hidden', timeout: 360000 });
        console.log('WINDOWS_NATIVE_CONVERT ' + item.id + ' ' + (Date.now() - conversionStarted));
      }
      stage = 'editor-title';
      await expect(page.locator('.topbar').getByRole('textbox', { name: '문서 제목', exact: true }))
        .toHaveValue(item.id, { timeout: 60000 });
      stage = 'editor-ready';
      await page.locator('.busy-indicator').waitFor({ state: 'hidden', timeout: 90000 });
      await page.evaluate(() => document.fonts.ready);
      const stack = page.locator('.document-stack');
      await expect(stack).toHaveAttribute('data-pages', /^[1-9]\d*$/, { timeout: 60000 });
      const editorPages = Number(await stack.getAttribute('data-pages'));
      const editorMilliseconds = Date.now() - readyStarted;
      const editorLongTasks = await page.evaluate(start => window.qaLongTasks.filter(task => task.at >= start)
        .map(task => Math.round(task.duration)).sort((a,b) => b-a).slice(0,5), rendererStarted);
      console.log('WINDOWS_HWP_EDITOR_READY ' + JSON.stringify({ id: item.id, editorPages, editorMilliseconds, editorLongTasks }));
      stage = 'preview-tab';
      const rendererPreviewStart = await page.evaluate(() => performance.now());
      const previewStarted = Date.now();
      page.on('console', message => {
        if (/^WINDOWS_HWP_(?:TAB|PRINT)_(?:DISPATCHED|DELIVERED|MEASURE|VISIBLE) [0-9]+$/.test(message.text()))
          console.log(message.text());
      });
      // Schedule the real button handlers from the renderer so actionability
      // waiting cannot hide the instant the event reaches the app.
      await page.evaluate(() => {
        const button = [...document.querySelectorAll('nav.tabs button')].find(el => el.textContent.trim() === '보기');
        if (!(button instanceof HTMLButtonElement)) throw new Error('View tab unavailable');
        const mark = (kind) => console.log('WINDOWS_HWP_' + kind + ' ' + Math.round(performance.now()));
        button.addEventListener('click', () => mark('TAB_DELIVERED'), { capture: true, once: true });
        setTimeout(() => { mark('TAB_DISPATCHED'); button.click(); }, 0);
      });
      stage = 'preview-command';
      const control = page.locator('.ribbon button.tool').filter({ hasText: /^인쇄 미리보기$/ });
      await control.waitFor({ state: 'attached', timeout: 180000 });
      await control.evaluate(button => {
        const mark = (kind) => console.log('WINDOWS_HWP_PRINT_' + kind + ' ' + Math.round(performance.now()));
        const inspector = setInterval(() => {
          if (!window.qaMeasuring && document.querySelector('.print-layout.measuring')) {
            window.qaMeasuring = true;
            mark('MEASURE');
          }
          if (document.querySelector('.print-preview-controls')) {
            clearInterval(inspector);
            mark('VISIBLE');
          }
        }, 50);
        button.addEventListener('click', () => mark('DELIVERED'), { capture: true, once: true });
        setTimeout(() => { mark('DISPATCHED'); button.click(); }, 0);
      });
      stage = 'preview-layout';
      const toolbar = page.locator('.print-preview-controls');
      await toolbar.waitFor({ state: 'visible', timeout: 420000 });
      const previewLongTasks = await page.evaluate(start => window.qaLongTasks.filter(task => task.at >= start)
        .map(task => Math.round(task.duration)).sort((a,b) => b-a).slice(0,8), rendererPreviewStart);
      console.log('WINDOWS_HWP_PREVIEW_READY ' + JSON.stringify({ id: item.id, previewLongTasks }));
      const previewPages = await page.locator('.print-layout.preview .print-page').count();
      const previewMilliseconds = Date.now() - previewStarted;
      stage = 'pdf';
      const pdfStarted = Date.now();
      await toolbar.getByRole('button', { name: 'PDF 저장', exact: true }).evaluate(button => setTimeout(() => button.click(), 0));
      await expect.poll(() => fs.stat(pdf).then(item => item.size).catch(() => 0), { timeout: 240000 }).toBeGreaterThan(0);
      const bytes = await fs.readFile(pdf);
      const pdfPages = (bytes.toString('latin1').match(/\/Type\s*\/Page\b/g) || []).length;
      if (!pdfPages) throw new Error('Native PDF page count unavailable');
      const metrics = { id: item.id, readerPages, editorPages, previewPages, pdfPages,
        previewMilliseconds, pdfMilliseconds: Date.now() - pdfStarted };
      console.log('WINDOWS_NATIVE_ORIGINAL ' + JSON.stringify(metrics));
      measured++;
    } catch (error) {
      // Playwright errors can include source content. Emit only fixed stage labels.
      const reason = error?.name === 'TimeoutError'
        ? (/intercepts pointer events/.test(error.message) ? 'pointer-interception' : 'timeout')
        : 'error';
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
