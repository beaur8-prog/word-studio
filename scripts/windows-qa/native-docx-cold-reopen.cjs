'use strict';

const { _electron, expect } = require('@playwright/test');
const { mkdtemp, readFile } = require('node:fs/promises');
const { createHash } = require('node:crypto');
const { basename, join, resolve } = require('node:path');
const smokeLifetime = require('../native-smoke-lifetime.cjs');

async function open(executable, source, label) {
  const profile = await mkdtemp(join(process.env.RUNNER_TEMP, 'schedule-'));
  const app = await _electron.launch({
    executablePath: resolve(executable),
    args: ['--user-data-dir=' + profile, '--force-device-scale-factor=1'],
    env: { ...process.env, STUDIO_DISABLE_UPDATES: '1' },
  });
  const close = await smokeLifetime(app, label, 90000);
  try {
    const page = await app.firstWindow();
    await page.locator('.document-stack .tiptap').first().waitFor();
    const title = page.locator('.topbar').getByRole('textbox', { name: '문서 제목', exact: true });
    const started = Date.now();
    await app.evaluate(({ dialog, Menu }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] });
      const command = Menu.getApplicationMenu().items.find(item => item.label === 'File').submenu.items
        .find(item => item.label === 'Open…');
      setTimeout(() => command.click(), 0);
    }, source);
    const reopened = await expect(title).toHaveValue(basename(source, '.docx'), { timeout: 30000 })
      .then(() => true, () => false);
    if (reopened) await page.locator('.busy-indicator').waitFor({ state: 'hidden', timeout: 30000 });
    const millis = Date.now() - started;
    const state = await page.evaluate(() => ({
      media: document.querySelectorAll('.document-stack .tiptap img').length,
      busy: !!document.querySelector('.busy-indicator'),
      visibility: document.visibilityState,
    }));
    return { label, reopened, millis, ...state };
  } finally {
    await close();
  }
}

async function main() {
  const [control, candidate, input] = process.argv.slice(2);
  if (!control || !candidate || !input) throw Error('Expected two isolated executable paths and one local test document');
  const source = resolve(input);
  const digest = async () => createHash('sha256').update(await readFile(source)).digest('hex');
  const before = await digest();
  const previous = await open(control, source, 'control');
  const current = await open(candidate, source, 'candidate');
  const sameInput = before === await digest();
  console.log('WINDOWS_COLD_DOCX_PAIR ' + JSON.stringify({ sameInput, previous, current }));
  if (!sameInput || !current.reopened || current.busy || current.media < 1)
    throw Error('Unexpected render result on the exact saved input');
}

main().catch(error => {
  console.error('WINDOWS_COLD_DOCX_FAILED ' + error.name);
  process.exitCode = 1;
});
