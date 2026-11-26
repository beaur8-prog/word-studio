'use strict';

const { _electron, expect } = require('@playwright/test');
const { mkdtemp, readFile } = require('node:fs/promises');
const { createHash } = require('node:crypto');
const { basename, join, resolve } = require('node:path');
const smokeLifetime = require('../native-smoke-lifetime.cjs');

async function selectFile(app, file) {
  await app.evaluate(({ dialog, Menu }, file) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] });
    const command = Menu.getApplicationMenu().items.find(item => item.label === 'File').submenu.items
      .find(item => item.label === 'Open…');
    setTimeout(() => command.click(), 0);
  }, file);
}

async function open(executable, source, label, outgoing) {
  const profile = await mkdtemp(join(process.env.RUNNER_TEMP, 'schedule-'));
  const app = await _electron.launch({
    executablePath: resolve(executable),
    args: ['--user-data-dir=' + profile, '--force-device-scale-factor=1'],
    env: { ...process.env, STUDIO_DISABLE_UPDATES: '1' },
  });
  const close = await smokeLifetime(app, label, 210000);
  try {
    const page = await app.firstWindow();
    await page.locator('.document-stack .tiptap').first().waitFor();
    const title = page.locator('.topbar').getByRole('textbox', { name: '문서 제목', exact: true });
    await selectFile(app, outgoing);
    const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
    await reader.getByTestId('hwp-page').waitFor({ state: 'visible', timeout: 75000 });
    await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
    await reader.getByRole('region', { name: '한글 편집 변환 안내' })
      .getByRole('button', { name: '편집본 만들기', exact: true }).click({ timeout: 90000 });
    await reader.waitFor({ state: 'hidden', timeout: 90000 });
    await expect(title).toHaveValue(basename(outgoing, '.hwp'), { timeout: 60000 });
    await page.locator('.busy-indicator').waitFor({ state: 'hidden', timeout: 90000 });
    const originalPictures = await page.locator('.document-stack .tiptap img').count();
    const started = Date.now();
    await selectFile(app, source);
    const reopened = await expect(title).toHaveValue(basename(source, '.docx'), { timeout: 45000 })
      .then(() => true, () => false);
    if (reopened) await page.locator('.busy-indicator').waitFor({ state: 'hidden', timeout: 30000 });
    const millis = Date.now() - started;
    const state = await page.evaluate(() => ({
      media: document.querySelectorAll('.document-stack .tiptap img').length,
      busy: !!document.querySelector('.busy-indicator'),
      visibility: document.visibilityState,
    }));
    return { label, reopened, millis, originalPictures, ...state };
  } finally {
    await close();
  }
}

async function main() {
  const [control, candidate, input, outgoing] = process.argv.slice(2);
  if (!control || !candidate || !input || !outgoing) throw Error('Expected two isolated executable paths and one local test document');
  const source = resolve(input);
  const digest = async () => createHash('sha256').update(await readFile(source)).digest('hex');
  const before = await digest();
  const previous = await open(control, source, 'control', outgoing);
  const current = await open(candidate, source, 'candidate', outgoing);
  const sameInput = before === await digest();
  console.log('WINDOWS_WARM_DOCX_PAIR ' + JSON.stringify({ sameInput, previous, current }));
  if (!sameInput || !current.reopened || current.busy || previous.originalPictures < 1 || current.media !== previous.originalPictures)
    throw Error('Unexpected render result on the exact saved input');
}

main().catch(error => {
  console.error('WINDOWS_WARM_DOCX_FAILED ' + error.name);
  process.exitCode = 1;
});
