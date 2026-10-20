'use strict';
const { _electron: electron, expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');

async function main() {
  const directory = await fs.mkdtemp(path.join(root, '.cache/hwp-native-'));
  const packaged = process.env.WORD_STUDIO_EXECUTABLE;
  const executablePath = packaged || require('electron');
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const app = await electron.launch({ executablePath, args: [...(packaged ? [] : [root]), '--user-data-dir=' + path.join(directory, 'profile')], env });
  const closeTestApp = await require('./native-smoke-lifetime.cjs')(app, 'HWP reader');
  const clipboard = await app.evaluate(({ clipboard }) => ({ text: clipboard.readText(), html: clipboard.readHTML(), rtf: clipboard.readRTF(), image: clipboard.readImage().toDataURL() }));
  try {
    const page = await app.firstWindow(), errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.setViewportSize({ width: 1440, height: 1000 });
    await page.locator('.tiptap').fill('NATIVE HWP KEEP — 기존 Word 문서');
    await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
    const inputs = ['tests/fixtures/hwp/reader.hwp', 'tests/fixtures/hwp/reader.hwpx', 'tests/fixtures/hwp/landscape-pagination.hwp', 'tests/fixtures/hwp/landscape-pagination.hwpx'];
    if (process.env.HWP_PUBLIC_SAMPLES === '1') inputs.push('analysis/hwp-validation/samples/01-employment-contract.hwp', 'analysis/hwp-validation/samples/03-health-duties.hwp');
    if (process.env.HWP_LONG_SAMPLE === '1') inputs.push('analysis/hwp-validation/samples/02-work-rules.hwp');
    const reports = [];
    for (const [i, source] of inputs.entries()) {
      console.log('HWP native check:', source);
      const fixture = path.join(root, source), pdf = path.join(directory, 'document-' + i + '.pdf');
      await app.evaluate(({ dialog, BrowserWindow }, files) => {
        dialog.showOpenDialog = async (_window, options) => {
          if (!options.filters.some(filter => filter.extensions.includes('hwp') && filter.extensions.includes('hwpx'))) throw Error('HWP file filter missing');
          return { canceled: false, filePaths: [files.fixture] };
        };
        dialog.showSaveDialog = async () => {
          const duplicated = await BrowserWindow.getAllWindows()[0].webContents.executeJavaScript(`(() => { const ids = [...document.querySelectorAll('[id]')].map(el => el.id); return ids.length !== new Set(ids).size; })()`);
          if (duplicated) throw Error('Printed pages must not share SVG resource IDs');
          return { canceled: false, filePath: files.pdf };
        };
      }, { fixture, pdf });
      await page.keyboard.press('ControlOrMeta+o');
      const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
      await expect(reader.getByTestId('hwp-page')).toBeVisible({ timeout: 60000 });
      await expect(reader.locator('.hwp-loading')).not.toBeVisible();
      const count = Number((await reader.getByTestId('hwp-page-count').textContent()).replace(/\D/g, ''));
      if (source.includes('landscape-pagination')) assert.equal(count, 4);
      if (source.includes('02-work-rules')) assert.equal(count, 163);
      await page.evaluate(() => document.fonts.ready);
      await reader.getByTestId('hwp-page').screenshot({ path: path.join(directory, 'document-' + i + '.png') });
      await reader.getByRole('button', { name: '텍스트 복사', exact: true }).click();
      await expect(reader.locator('.hwp-status')).toContainText('텍스트를 복사했습니다');
      const copied = await app.evaluate(({ clipboard }) => clipboard.readText());
      assert.ok(copied.length > 10);
      if (i < 2) assert.ok(copied.includes('테이블 속 한글'));
      if (source.includes('landscape-pagination')) {
        for (let row = 1; row <= 24; row++) for (let col = 1; col <= 2; col++) {
          const label = `ROW-${String(row).padStart(2, '0')}-COLUMN-${col}`;
          assert.equal(copied.split(label).length - 1, 1, 'Each numbered table cell must be copied once');
        }
      }
      await reader.getByRole('button', { name: 'PDF 저장', exact: true }).click();
      await expect(reader.locator('.hwp-status')).toContainText('PDF를 저장했습니다', { timeout: 60000 });
      const bytes = await fs.readFile(pdf), pages = (bytes.toString('latin1').match(/\/Type\s*\/Page\b/g) || []).length;
      assert.equal(pages, count, 'Native PDF must preserve the HWP reader pages');
      await expect(page.locator('body')).not.toHaveClass(/studio-printing/);
      assert.equal(await page.locator('body > .hwp-print-layout').count(), 0);
      await reader.getByRole('button', { name: 'HWP 읽기 닫기' }).click();
      await expect(page.locator('.tiptap')).toHaveText('NATIVE HWP KEEP — 기존 Word 문서');
      reports.push({ source, pages: count, pdfPages: pages, pdfBytes: bytes.length, clipboardCharacters: copied.length, pdf, screenshot: path.join(directory, 'document-' + i + '.png') });
    }
    assert.deepEqual(errors, []);
    const result = { passed: true, packaged: Boolean(packaged), platform: process.platform, reports, rendererErrors: errors, directory };
    await fs.writeFile(path.join(directory, 'result.json'), JSON.stringify(result, null, 2) + '\n');
    console.log(JSON.stringify(result, null, 2));
  } catch (error) { console.error('Native HWP assertion failed:', error); throw error; }
  finally {
    await app.evaluate(({ clipboard, nativeImage }, value) => clipboard.write({ ...value, image: nativeImage.createFromDataURL(value.image) }), clipboard);
    await closeTestApp();
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
