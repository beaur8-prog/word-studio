'use strict';

const { spawn } = require('node:child_process');
const { mkdir, mkdtemp, readFile, writeFile } = require('node:fs/promises');
const { dirname, join, resolve } = require('node:path');
const { pathToFileURL } = require('node:url');
const { _electron, expect } = require('@playwright/test');
const { Document, Packer, Paragraph } = require('docx');
const smokeLifetime = require('../native-smoke-lifetime.cjs');

async function makeInputs(directory) {
  const items = [
    { name: '첫 문서 · 확인 ①', ext: 'docx', text: 'DOCX 셸 열기 기준' },
    { name: '다음 문서 · 그림 📝', ext: 'hwp', text: 'HWP 셸 열기 기준' },
    { name: '세 번째 문서 e\u0301', ext: 'HWPX', text: 'HWPX 셸 열기 기준' },
  ];
  const engineRoot = dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(join(engineRoot, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(join(engineRoot, 'rhwp_bg.wasm')) });
  for (const item of items) {
    item.filename = item.name + '.' + item.ext;
    item.path = join(directory, item.filename);
    let bytes;
    if (item.ext === 'docx') {
      const doc = new Document({ sections: [{ children: [new Paragraph(item.text)] }] });
      bytes = await Packer.toBuffer(doc);
    } else {
      const doc = HwpDocument.createEmpty();
      try {
        doc.createBlankDocument();
        doc.insertText(0, 0, 0, item.text);
        bytes = item.ext === 'hwp' ? doc.exportHwp() : doc.exportHwpx();
      } finally { doc.free(); }
    }
    await writeFile(item.path, bytes);
  }
  return items;
}

async function launchAssociation(executable, profile, argument, cwd, env) {
  const child = spawn(executable, ['--user-data-dir=' + profile, argument], { cwd, env, windowsHide: true, stdio: 'ignore' });
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => { child.kill(); reject(new Error('Secondary process remained open')); }, 30000);
    child.once('error', error => { clearTimeout(timer); reject(error); });
    child.once('exit', code => {
      clearTimeout(timer);
      if (code === 0) resolve();
      else reject(new Error('Secondary process did not hand off its document'));
    });
  });
}

async function main() {
  const executable = resolve(process.env.WORD_STUDIO_EXECUTABLE);
  const workspace = await mkdtemp(join(process.env.RUNNER_TEMP, 'studio-shell-'));
  const profile = join(workspace, 'test-profile');
  const folder = join(workspace, 'Windows 문서 시험 · 한글');
  await mkdir(folder);
  const [docx, ...hangul] = await makeInputs(folder);
  const env = { ...process.env, STUDIO_DISABLE_UPDATES: '1' };
  const app = await _electron.launch({ executablePath: executable, args: ['--user-data-dir=' + profile, docx.path], env });
  const close = await smokeLifetime(app, 'Unicode shell open');
  try {
    const page = await app.firstWindow();
    const title = page.locator('.topbar').getByRole('textbox', { name: '문서 제목', exact: true });
    const editor = page.locator('.document-stack .tiptap').first();
    const busy = page.locator('.busy-indicator');
    const verifyEditor = async item => {
      await expect(title).toHaveValue(item.name, { timeout: 90000 });
      await busy.waitFor({ state: 'hidden', timeout: 90000 });
      await expect(editor).toContainText(item.text);
      const windows = await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length);
      if (windows !== 1) throw Error('A second editor window was opened');
      console.log('WINDOWS_UNICODE_SHELL_OPEN ' + JSON.stringify({ extension: item.ext.toLowerCase(), windows, title: true, text: true }));
    };
    await verifyEditor(docx);
    for (const [index, item] of hangul.entries()) {
      await launchAssociation(executable, profile, index ? item.filename : item.path, folder, env);
      const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
      await reader.getByTestId('hwp-page').waitFor({ state: 'visible', timeout: 75000 });
      await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
      await reader.getByRole('button', { name: '편집본 만들기', exact: true }).click({ timeout: 90000 });
      await reader.waitFor({ state: 'hidden', timeout: 90000 });
      await verifyEditor(item);
    }
  } finally { await close(); }
}

main().catch(error => {
  console.error('WINDOWS_UNICODE_SHELL_FAILED ' + error.name);
  process.exitCode = 1;
});
