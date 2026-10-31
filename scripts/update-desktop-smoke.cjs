"use strict";
// Exercise the real main/preload/renderer bridge with a disposable profile and
// harmless local bytes. No real release is fetched or installer executed.
const { _electron: electron, expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const fsSync = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
async function run() {
  await fs.mkdir(path.join(root, '.cache'), { recursive: true });
  const directory = await fs.mkdtemp(path.join(root, '.cache', 'update-desktop-'));
  const profile = path.join(directory, 'profile');
  const bootstrap = path.join(directory, 'bootstrap.cjs');
  await fs.writeFile(bootstrap, `
    const { app, BrowserWindow, dialog, shell, Notification } = require('electron');
    const { createHash } = require('node:crypto');
    const root = ${JSON.stringify(root)};
    Object.defineProperty(app, 'isPackaged', { get: () => true });
    app.getVersion = () => require(root + '/package.json').version;
    BrowserWindow.prototype.show = () => {};
    BrowserWindow.prototype.focus = () => {};
    Notification.isSupported = () => false;
    const payload = Buffer.alloc(2 * 1024 * 1024, 90);
    const version = '99.0.0';
    const filename = process.platform === 'win32' ? 'Word-Studio-' + version + '-windows-x64-setup.exe'
      : 'Word-Studio-' + version + '-mac-' + process.arch + '.zip';
    const url = 'https://github.com/beaur8-prog/word-studio/releases/download/v' + version + '/' + filename;
    const release = { tag_name: 'v' + version, assets: [{ name: filename, state: 'uploaded', size: payload.length,
      digest: 'sha256:' + createHash('sha256').update(payload).digest('hex'), browser_download_url: url }] };
    globalThis.updateTest = { requests: [], dialogs: [], opened: [], fail: false };
    dialog.showMessageBox = async (...args) => { globalThis.updateTest.dialogs.push(args.at(-1)); return { response: 1 }; };
    shell.openExternal = async url => { globalThis.updateTest.opened.push(url); };
    shell.showItemInFolder = file => { globalThis.updateTest.opened.push(file); };
    globalThis.fetch = async (request, options) => {
      globalThis.updateTest.requests.push(request);
      if (request.startsWith('https://api.github.com/repos/beaur8-prog/word-studio/releases')) {
        if (globalThis.updateTest.fail) return new Response('', { status: 503 });
        return Response.json(process.platform === 'darwin' ? [release] : release);
      }
      if (request !== url) throw new Error('Unexpected network request: ' + request);
      return new Response(new ReadableStream({ start(stream) {
        stream.enqueue(payload.subarray(0, payload.length / 2));
        const abort = () => stream.error(options.signal.reason);
        options.signal.addEventListener('abort', abort, { once: true });
        globalThis.updateTest.finish = () => {
          options.signal.removeEventListener('abort', abort);
          stream.enqueue(payload.subarray(payload.length / 2)); stream.close();
        };
      } }));
    };
    require(root + '/electron/main.cjs');
  `);
  const suffix = process.platform === 'darwin' ? 'Electron.app/Contents/MacOS/Electron'
    : process.platform === 'win32' ? 'electron.exe' : 'electron';
  const executablePath = [root, path.dirname(root), path.join(path.dirname(root), 'slide-studio')]
    .map(dir => path.join(dir, 'node_modules/electron/dist', suffix)).find(file => fsSync.existsSync(file));
  if (!executablePath) throw new Error('Electron is required for the update desktop smoke test.');
  const env = { ...process.env, STUDIO_DISABLE_UPDATES: '1' };
  delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const application = await electron.launch({ executablePath, args: [bootstrap, `--user-data-dir=${profile}`], env });
  const close = await require('./native-smoke-lifetime.cjs')(application, 'Updater');
  try {
    const window = await application.firstWindow();
    await window.setViewportSize({ width: 1440, height: 1000 });
    const errors = [];
    window.on('pageerror', error => errors.push(error.message));
    const editor = window.locator('.tiptap[role="textbox"]');
    await expect(editor).toBeVisible();
    await editor.fill('업데이트 중에도 보존할 문서 · Preserve this document');
    const before = await editor.innerText();
    await window.getByRole('button', { name: '파일', exact: true }).click();
    await window.locator('.backstage-nav').getByRole('button', { name: '수동 업데이트·설정', exact: true }).click();
    const dialog = window.getByRole('dialog', { name: 'Word Studio 업데이트', exact: true });
    await expect(dialog).toBeVisible();
    await expect(dialog.getByText(require('../package.json').version, { exact: true })).toBeVisible();
    const phases = [];
    await window.evaluate(() => {
      window.updatePhases = [];
      window.wordDesktop.onUpdateState(state => window.updatePhases.push(state.phase));
    });
    if (process.platform !== 'linux') {
      await dialog.getByRole('checkbox').uncheck();
      await expect.poll(() => window.evaluate(() => window.wordDesktop.getUpdateState().then(state => state.automatic))).toBe(false);
      await dialog.getByRole('button', { name: '업데이트 확인', exact: true }).click();
      await expect(dialog.getByRole('button', { name: /^다운로드 \(/ })).toBeVisible();
      assert.equal(await application.evaluate(() => globalThis.updateTest.requests.length), 1);
      await dialog.getByRole('button', { name: /^다운로드 \(/ }).click();
      await expect(dialog.getByRole('progressbar', { name: '업데이트 다운로드' })).toHaveAttribute('value', '50');
      await window.screenshot({ path: path.join(directory, 'download-50.png') });
      await dialog.getByRole('button', { name: '취소', exact: true }).click();
      await expect(dialog.getByRole('button', { name: /^다운로드 \(/ })).toBeEnabled();
      await dialog.getByRole('button', { name: /^다운로드 \(/ }).click();
      await expect(dialog.getByRole('progressbar', { name: '업데이트 다운로드' })).toHaveAttribute('value', '50');
      await dialog.getByRole('button', { name: '업데이트 창 닫기', exact: true }).click();
      await expect(window.locator('.update-status')).toContainText('50%');
      await window.locator('.update-status').click();
      await application.evaluate(() => globalThis.updateTest.finish());
      await expect(dialog.getByRole('button', { name: process.platform === 'darwin' ? '다운로드 파일·교체 안내' : '종료하고 설치', exact: true })).toBeVisible();
      assert.equal(await application.evaluate(() => globalThis.updateTest.dialogs.length), 0, 'Downloading alone must never prompt installation.');
      await window.screenshot({ path: path.join(directory, 'ready.png') });
      phases.push(...await window.evaluate(() => window.updatePhases));
      for (const phase of ['checking', 'available', 'downloading', 'verifying', 'ready']) assert.ok(phases.includes(phase));
      // Exercise both the macOS application menu and cross-platform Help menu.
      await dialog.getByRole('button', { name: '업데이트 창 닫기', exact: true }).click();
      await application.evaluate(({ Menu }) => Menu.getApplicationMenu().items.find(item => item.label === '도움말').submenu.items[0].click());
      await expect(dialog).toBeVisible();
      await expect(dialog.getByRole('button', { name: process.platform === 'darwin' ? '다운로드 파일·교체 안내' : '종료하고 설치', exact: true })).toBeVisible();
      await application.evaluate(() => { globalThis.updateTest.fail = true; });
      await dialog.getByRole('button', { name: '업데이트 확인', exact: true }).click();
      await expect(dialog.getByRole('alert')).toContainText('503');
      await application.evaluate(() => { globalThis.updateTest.fail = false; });
      await dialog.getByRole('button', { name: '업데이트 확인', exact: true }).click();
      await expect(dialog.getByRole('alert')).toHaveCount(0);
      if (process.platform === 'darwin') {
        await dialog.getByRole('button', { name: '다운로드 파일·교체 안내', exact: true }).click();
        await expect.poll(() => application.evaluate(() => globalThis.updateTest.dialogs.length)).toBe(1);
        assert.ok((await application.evaluate(() => globalThis.updateTest.dialogs[0].detail)).includes('직접 교체'));
      }
    } else {
      await expect(dialog.getByRole('button', { name: '업데이트 확인', exact: true })).toBeDisabled();
      assert.equal(await application.evaluate(() => globalThis.updateTest.requests.length), 0);
    }
    await dialog.getByRole('button', { name: '다운로드 페이지 열기', exact: true }).click();
    assert.deepEqual(await application.evaluate(() => globalThis.updateTest.opened), ['https://github.com/beaur8-prog/word-studio/releases/latest']);
    await dialog.getByRole('button', { name: '업데이트 창 닫기', exact: true }).click();
    assert.equal(await editor.innerText(), before);
    assert.deepEqual(errors, []);
    await fs.writeFile(path.join(directory, 'result.json'), JSON.stringify({ passed: true, platform: process.platform, phases,
      documentUnchanged: true, rendererErrors: errors, isolatedProfile: true, realNetwork: false, installerExecuted: false }, null, 2));
    console.log('Update desktop smoke passed:', path.join(directory, 'result.json'));
  } finally {
    await close();
    await fs.rm(profile, { recursive: true, force: true });
    await fs.rm(bootstrap, { force: true });
  }
}
run().catch(error => { console.error(error); process.exitCode = 1; });
