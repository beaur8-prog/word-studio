'use strict';
const { _electron: electron, expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const path = require('node:path');
const assert = require('node:assert/strict');
const { createHash } = require('node:crypto');
const root = path.resolve(__dirname, '..');

async function main() {
  const directory = await fs.mkdtemp(path.join(root, '.cache/hwp-editing-native-'));
  const packaged = process.env.WORD_STUDIO_EXECUTABLE;
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const app = await electron.launch({ executablePath: packaged || require('electron'), args: [...(packaged ? [] : [root]), '--user-data-dir=' + path.join(directory, 'profile')], env });
  const close = await require('./native-smoke-lifetime.cjs')(app, 'HWP editing');
  try {
    const page = await app.firstWindow(), errors = [], reports = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.setViewportSize({ width: 1440, height: 1000 });
    await page.locator('.tiptap').fill('NATIVE HWP EDITING — disposable test document');
    await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
    const open = async file => {
      await app.evaluate(({ dialog }, value) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [value] }); }, file);
      await page.keyboard.press('ControlOrMeta+o');
      const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
      await expect(reader.getByTestId('hwp-page')).toBeVisible({ timeout: 60000 });
      await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
      const panel = reader.getByRole('region', { name: '한글 편집 변환 안내' });
      await expect(panel).toContainText('쪽 수 비교', { timeout: 60000 });
      await panel.getByRole('button', { name: '편집본 만들기', exact: true }).click();
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(file).replace(/\.(hwp|hwpx)$/i, ''));
    };
    await open(path.join(root, 'tests/fixtures/hwp/reader.hwp'));
    await page.locator('.tiptap td').filter({ hasText: '테이블 속 한글' }).locator('p').first().fill('NATIVE 편집한 표 2026');
    for (const format of ['hwp', 'hwpx']) {
      const file = path.join(directory, 'edited.' + format);
      await app.evaluate(({ dialog }, value) => {
        dialog.showSaveDialog = async (_window, options) => {
          if (!options.filters.some(filter => filter.extensions.includes(value.format))) throw Error('Missing native Hancom save filter');
          return { canceled: true };
        };
      }, { format });
      await app.evaluate(({ Menu }, value) => {
        const fileMenu = Menu.getApplicationMenu().items.find(item => item.label === 'File');
        const item = fileMenu.submenu.items.find(item => item.label === `Export Hangul Document (${value.toUpperCase()})…`);
        if (!item) throw Error('Missing native export menu'); item.click();
      }, format);
      const dialog = page.getByRole('dialog', { name: '한글 문서로 내보내기', exact: true });
      await expect(dialog).toBeVisible({ timeout: 60000 });
      await dialog.getByRole('button', { name: format.toUpperCase() + ' 파일 저장', exact: true }).click();
      await expect(dialog.getByRole('button', { name: format.toUpperCase() + ' 파일 저장', exact: true })).toBeEnabled();
      await expect(dialog).toBeVisible(); // Native cancellation retains the staged export.
      await app.evaluate(({ dialog }, value) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: value }); }, file);
      await dialog.getByRole('button', { name: format.toUpperCase() + ' 파일 저장', exact: true }).click();
      await expect(dialog).not.toBeVisible();
      const bytes = await fs.readFile(file);
      assert.equal(bytes.subarray(0, format === 'hwp' ? 8 : 2).toString('hex'), format === 'hwp' ? 'd0cf11e0a1b11ae1' : '504b');
      await open(file);
      await expect(page.locator('.tiptap')).toContainText('NATIVE 편집한 표 2026');
      await expect(page.locator('.tiptap table')).toHaveCount(1);
      reports.push({ format, bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex'), file });
    }
    await page.screenshot({ path: path.join(directory, 'edited-word-screen.png') });
    assert.deepEqual(errors, []);
    const report = { passed: true, platform: process.platform, packaged: Boolean(packaged), nativeCancellation: true, nativeMenuAndFilters: true, reports, rendererErrors: errors, directory };
    await fs.writeFile(path.join(directory, 'result.json'), JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report, null, 2));
  } finally { await close(); }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
