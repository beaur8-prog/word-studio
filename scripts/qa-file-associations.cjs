'use strict';

// Exercise only disposable profiles; never close or modify the user's session.
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { _electron: electron, expect } = require('@playwright/test');
const { Document, Paragraph, Packer } = require('docx');

(async () => {
  const executablePath = process.argv[2], baseline = process.argv.includes('--baseline');
  if (!executablePath) throw new Error('Pass a packaged Word Studio executable.');
  const output = path.resolve(process.argv[3] || '.cache/file-open/native'); await fs.mkdir(output, { recursive: true });
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), 'word-file-open-'));
  const profile = path.join(temporary, 'profile'), documents = path.join(temporary, '한글과 공백 폴더');
  await fs.mkdir(documents);
  const first = path.join(documents, '첫 번째 외부 문서.DOCX'), second = path.join(documents, '두 번째 외부 문서.txt');
  await fs.writeFile(first, await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('더블클릭으로 연 DOCX 본문')] }] })));
  await fs.writeFile(second, '실행 중 더블클릭으로 연 두 번째 본문');
  const broken = path.join(documents, '손상된 문서.docx'); await fs.writeFile(broken, 'not a zip');
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const result = { passed: false, platform: process.platform, baseline, isolatedProfile: true, checks: [], rendererErrors: [] };
  let application, close;
  const launch = async (files = []) => {
    application = await electron.launch({ executablePath, args: ['--user-data-dir=' + profile, ...files], env, timeout: 45000 });
    close = await require('./native-smoke-lifetime.cjs')(application, 'OS file open');
    const page = await application.firstWindow(); page.setDefaultTimeout(15000);
    page.on('pageerror', error => result.rendererErrors.push(error.message));
    return page;
  };
  const emit = (...files) => application.evaluate(({ app }, filenames) => filenames.map(filename => app.emit('open-file', { preventDefault() {} }, filename)), files);
  const shown = async (page, text) => {
    await expect(page.locator('.document-stack .tiptap')).toHaveText(text);
    await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
    await expect(page.locator('.busy-indicator')).toHaveCount(0);
  };
  try {
    let page = await launch([first]);
    await page.locator('.document-stack .tiptap').waitFor();
    if (baseline) {
      await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
      expect(await page.locator('.document-stack .tiptap').textContent()).not.toContain('더블클릭으로 연 DOCX 본문');
      expect(await application.evaluate(({ app }) => app.listenerCount('open-file'))).toBe(0);
      result.checks.push('reproduced: startup document ignored and no macOS open-file listener');
    } else {
      await shown(page, '더블클릭으로 연 DOCX 본문'); result.checks.push('cold launch DOCX with Korean/space/uppercase path');
      await emit(second); await shown(page, '실행 중 더블클릭으로 연 두 번째 본문'); result.checks.push('running application OS request');
      await page.locator('.document-stack .tiptap').fill('전환 전에 보호해야 할 편집 내용');
      await emit(first); await shown(page, '더블클릭으로 연 DOCX 본문');
      const recovery = await page.evaluate(() => window.wordDesktop.readRecovery());
      expect(recovery.some(record => record.current.includes('전환 전에 보호해야 할 편집 내용'))).toBe(true);
      result.checks.push('current edits preserved in recovery before switching');
      await emit(broken, path.join(documents, 'missing.docx'), second);
      await shown(page, '실행 중 더블클릭으로 연 두 번째 본문'); result.checks.push('corrupt and missing requests do not block the next file');
      await application.evaluate(({ dialog }) => { dialog.showOpenDialog = async (_window, options) => {
        globalThis.__fileOpenDirectory = options.defaultPath; return { canceled: true, filePaths: [] };
      }; });
      await page.evaluate(() => window.wordDesktop.openFile());
      expect(await application.evaluate(() => globalThis.__fileOpenDirectory)).toBe(documents);
      result.checks.push('native chooser follows the externally opened folder');
      const hwp = path.resolve(__dirname, '../tests/fixtures/hwp/reader.hwp');
      await emit(hwp, first);
      await expect(page.getByRole('dialog', { name: 'HWP 읽기', exact: true })).toBeVisible();
      await expect(page.locator('.hwp-sheet svg')).not.toHaveCount(0);
      expect(await page.locator('.document-stack .tiptap').textContent()).toBe('실행 중 더블클릭으로 연 두 번째 본문');
      await page.getByRole('button', { name: 'HWP 읽기 닫기', exact: true }).click();
      await shown(page, '더블클릭으로 연 DOCX 본문'); result.checks.push('HWP uses original viewer and defers the next request until close');
      await application.evaluate(({ app }, args) => app.emit('second-instance', {}, ['Word Studio', args.name], args.documents), { name: path.basename(second), documents });
      await shown(page, '실행 중 더블클릭으로 연 두 번째 본문'); result.checks.push('second-instance relative launch argument');
      await page.screenshot({ path: path.join(output, 'native-opened.png') });
      const closed = page.waitForEvent('close'); await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close()); await closed;
      const created = application.waitForEvent('window'); await emit(first); page = await created; page.setDefaultTimeout(15000);
      await shown(page, '더블클릭으로 연 DOCX 본문');
      expect(await application.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(1);
      result.checks.push('closed editor recreated with one document window');
      await close(); application = null;
      page = await launch(); await shown(page, '더블클릭으로 연 DOCX 본문'); result.checks.push('normal launch still restores the most recent document');
    }
    expect(result.rendererErrors).toEqual([]); result.passed = true;
    await fs.writeFile(path.join(output, 'summary.json'), JSON.stringify(result, null, 2) + '\n'); console.log(JSON.stringify(result));
  } catch (error) {
    if (application) { try { const page = await application.firstWindow(); await page.screenshot({ path: path.join(output, 'failure.png') }); await fs.writeFile(path.join(output, 'failure.txt'), await page.locator('body').innerText()); } catch {} }
    throw error;
  } finally {
    if (application) await close().catch(() => application.process().kill('SIGKILL'));
    await fs.rm(temporary, { recursive: true, force: true });
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
