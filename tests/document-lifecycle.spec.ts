import { test, expect, type Page } from '@playwright/test';

async function openText(page: Page, name: string, text: string) {
  await page.locator('input[type=file]').first().setInputFiles({ name, mimeType: 'text/plain', buffer: Buffer.from(text) });
}
const editor = (page: Page) => page.locator('.document-stack .tiptap');
const saved = (page: Page) => page.getByText('이 PC에 저장됨', { exact: true });
const failureDialog = (page: Page) => page.getByRole('dialog', { name: '현재 문서를 저장하지 못했습니다' });

// Exercise renderer behavior with controllable disk failures and latency.
async function desktop(page: Page) {
  await page.addInitScript(() => {
    const state = { fail: false, held: false, writes: [] as string[], files: [] as string[], result: 'success', release: null as null | (() => void) };
    (window as any).testDisk = state;
    (window as any).wordDesktop = {
      platform: 'test', readRecovery: async () => [], setTheme: () => {}, onMenu: () => () => {},
      saveRecovery: async (serialized: string) => {
        if (state.fail) throw new Error('disk full');
        if (state.held) await new Promise<void>(resolve => { state.release = resolve; });
        state.writes.push(serialized); return true;
      },
      saveFile: async (payload: { data: Uint8Array }) => {
        if (state.result === 'cancel') return false;
        if (state.result === 'error') throw new Error('disk full');
        state.files.push(new TextDecoder().decode(payload.data)); return true;
      },
    };
  });
  await page.goto('/');
  await openText(page, 'A.txt', 'ORIGINAL CONTENT');
  await expect(editor(page)).toHaveText('ORIGINAL CONTENT');
  await expect(saved(page)).toBeVisible();
}

async function failSwitch(page: Page) {
  await page.evaluate(() => { (window as any).testDisk.fail = true; });
  await editor(page).fill('UNSAVED IMPORTANT CONTENT');
  await openText(page, 'B.txt', 'SECOND DOCUMENT');
  await expect(failureDialog(page)).toBeVisible();
  await expect(editor(page)).toHaveText('UNSAVED IMPORTANT CONTENT');
  await expect(page.getByRole('textbox', { name: '문서 제목' })).toHaveValue('A');
}

test('opening another document resets undo and redo, while saving retains undo', async ({ page }) => {
  await page.goto('/');
  await openText(page, 'A.txt', 'FIRST DOCUMENT');
  await expect(editor(page)).toHaveText('FIRST DOCUMENT');
  await expect(saved(page)).toBeVisible();
  await openText(page, 'B.txt', 'SECOND DOCUMENT');
  await expect(editor(page)).toHaveText('SECOND DOCUMENT');
  const undo = page.getByRole('button', { name: '실행 취소 (Ctrl+Z)', exact: true });
  const redo = page.getByRole('button', { name: '다시 실행 (Ctrl+Y)', exact: true });
  await expect(undo).toBeDisabled(); await expect(redo).toBeDisabled();
  await editor(page).fill('EDITED SECOND DOCUMENT');
  const download = page.waitForEvent('download');
  await page.keyboard.press('ControlOrMeta+s'); await download;
  await undo.click();
  await expect(editor(page)).toHaveText('SECOND DOCUMENT');
  await redo.click();
  await expect(editor(page)).toHaveText('EDITED SECOND DOCUMENT');
  await page.keyboard.press('ControlOrMeta+n');
  await page.getByRole('button', { name: /빈 문서 자유롭게/ }).click();
  await expect(editor(page)).toHaveText('');
  await expect(undo).toBeDisabled(); await expect(redo).toBeDisabled();
});

test('failed recovery keeps the current document and Cancel restores editing', async ({ page }) => {
  await desktop(page); await failSwitch(page);
  await expect(failureDialog(page).getByRole('button', { name: '취소', exact: true })).toBeFocused();
  await page.screenshot({ path: 'test-results/save-failure-dialog.png' });
  await page.keyboard.press('Escape');
  await expect(failureDialog(page)).toHaveCount(0);
  await expect(editor(page)).toHaveText('UNSAVED IMPORTANT CONTENT');
  await page.evaluate(() => { (window as any).testDisk.fail = false; });
  await openText(page, 'B.txt', 'SECOND DOCUMENT');
  await expect(editor(page)).toHaveText('SECOND DOCUMENT');
  const writes = await page.evaluate(() => (window as any).testDisk.writes.map((s: string) => JSON.parse(s)));
  expect(writes.some((d: any) => d.title === 'A' && JSON.stringify(d.content).includes('UNSAVED IMPORTANT CONTENT'))).toBe(true);
});

test('cancelled or failed file saves retain the dialog; successful save preserves a full copy before switching', async ({ page }) => {
  await desktop(page); await failSwitch(page);
  for (const result of ['cancel', 'error']) {
    await page.evaluate(result => { (window as any).testDisk.result = result; }, result);
    await failureDialog(page).getByRole('button', { name: '파일로 저장 후 계속' }).click();
    await expect(failureDialog(page).getByRole('button', { name: '파일로 저장 후 계속' })).toBeEnabled();
    await expect(editor(page)).toHaveText('UNSAVED IMPORTANT CONTENT');
  }
  await page.evaluate(() => { (window as any).testDisk.result = 'success'; });
  await failureDialog(page).getByRole('button', { name: '파일로 저장 후 계속' }).click();
  await expect(failureDialog(page)).toHaveCount(0);
  await expect(editor(page)).toHaveText('SECOND DOCUMENT');
  const file = await page.evaluate(() => JSON.parse((window as any).testDisk.files[0]));
  expect(file.title).toBe('A');
  expect(JSON.stringify(file.content)).toContain('UNSAVED IMPORTANT CONTENT');
});

test('discard is explicit and document switching leaves no old undo history', async ({ page }) => {
  await desktop(page); await failSwitch(page);
  await failureDialog(page).getByRole('button', { name: '저장하지 않고 계속' }).click();
  await expect(editor(page)).toHaveText('SECOND DOCUMENT');
  await expect(page.getByRole('button', { name: '실행 취소 (Ctrl+Z)', exact: true })).toBeDisabled();
});

test('edits made while a recovery write is pending are flushed before switching', async ({ page }) => {
  await desktop(page);
  await page.evaluate(() => { (window as any).testDisk.held = true; });
  await editor(page).fill('EDIT BEFORE WRITE');
  await openText(page, 'B.txt', 'SECOND DOCUMENT');
  await expect.poll(() => page.evaluate(() => Boolean((window as any).testDisk.release))).toBe(true);
  await editor(page).fill('EDIT DURING WRITE');
  await page.evaluate(() => { const disk = (window as any).testDisk; disk.held = false; disk.release(); });
  await expect(editor(page)).toHaveText('SECOND DOCUMENT');
  const writes = await page.evaluate(() => (window as any).testDisk.writes.map((s: string) => JSON.parse(s)).filter((d: any) => d.title === 'A'));
  expect(JSON.stringify(writes.at(-1).content)).toContain('EDIT DURING WRITE');
});
