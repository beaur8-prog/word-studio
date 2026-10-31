import { test, expect } from '@playwright/test';

// Delay only the native boundary; exercise the actual editor commands and UI.
async function desktop(page: import('@playwright/test').Page) {
  await page.addInitScript(() => {
    const state = { opens: 0, saves: 0, writes: [] as string[], listener: null as any,
      finish: null as any, fail: null as any };
    (window as any).panels = state;
    const panel = (kind: string) => {
      state.listener?.(kind);
      return new Promise<any>((resolve, reject) => {
        state.finish = (value: any) => { state.listener?.(null); resolve(value); };
        state.fail = () => { state.listener?.(null); reject(Error('native panel failed')); };
      });
    };
    (window as any).wordDesktop = { platform: 'darwin', setTheme() {}, onMenu: () => () => {},
      readRecovery: async () => [], saveRecovery: async (raw: string) => { state.writes.push(raw); return true; },
      onFileOperation: (listener: any) => { state.listener = listener; listener(null); return () => { state.listener = null; }; },
      openFile: () => { state.opens++; return panel('open'); },
      saveFile: () => { state.saves++; return panel('save'); } };
  });
  await page.goto('/'); await expect(page.locator('.tiptap')).toBeVisible();
}
test('open waits visibly, ignores duplicate commands, clears on cancel and permits another selection', async ({ page }) => {
  await desktop(page); await page.locator('.tiptap').fill('KEEP BEFORE NATIVE PANEL');
  await page.keyboard.press('ControlOrMeta+o');
  await expect(page.getByRole('status').filter({ hasText: '파일 선택창을 기다리는 중' })).toBeVisible();
  await page.keyboard.press('ControlOrMeta+o'); await page.keyboard.press('ControlOrMeta+s');
  expect(await page.evaluate(() => ({ opens: (window as any).panels.opens, saves: (window as any).panels.saves }))).toEqual({ opens: 1, saves: 0 });
  await page.evaluate(() => (window as any).panels.finish(null));
  await expect(page.getByRole('status').filter({ hasText: '파일 선택창을 기다리는 중' })).toHaveCount(0);
  await expect(page.locator('.tiptap')).toHaveText('KEEP BEFORE NATIVE PANEL');
  await page.keyboard.press('ControlOrMeta+o');
  await expect.poll(() => page.evaluate(() => (window as any).panels.opens)).toBe(2);
  await page.evaluate(() => (window as any).panels.finish({ name: '선택.txt', data: new TextEncoder().encode('SELECTED AFTER CANCEL') }));
  await expect(page.locator('.tiptap')).toHaveText('SELECTED AFTER CANCEL');
  expect(await page.evaluate(() => (window as any).panels.writes.some((raw: string) => raw.includes('KEEP BEFORE NATIVE PANEL')))).toBe(true);
});
test('Save As cancel and error clear waiting state and allow a later save without changing the document', async ({ page }) => {
  await desktop(page); await page.locator('.tiptap').fill('KEEP THROUGH SAVE PANELS');
  await page.keyboard.press('ControlOrMeta+Shift+s');
  await page.getByRole('region', { name: '다른 형식으로 저장' }).getByRole('button', { name: /^Word 문서/ }).click();
  await expect(page.getByRole('status').filter({ hasText: '저장 위치 선택창을 기다리는 중' })).toBeVisible();
  await page.keyboard.press('ControlOrMeta+o'); await page.keyboard.press('ControlOrMeta+s');
  expect(await page.evaluate(() => ({ opens: (window as any).panels.opens, saves: (window as any).panels.saves }))).toEqual({ opens: 0, saves: 1 });
  await page.evaluate(() => (window as any).panels.finish(false));
  await expect(page.locator('.busy-indicator')).toHaveCount(0);
  await page.keyboard.press('ControlOrMeta+s');
  await expect.poll(() => page.evaluate(() => (window as any).panels.saves)).toBe(2);
  await page.evaluate(() => (window as any).panels.fail());
  await expect(page.getByRole('status').filter({ hasText: 'native panel failed' })).toBeVisible();
  await expect(page.locator('.busy-indicator')).toHaveCount(0);
  await page.keyboard.press('ControlOrMeta+s');
  await expect.poll(() => page.evaluate(() => (window as any).panels.saves)).toBe(3);
  await page.evaluate(() => (window as any).panels.finish(true));
  await expect(page.locator('.busy-indicator')).toHaveCount(0);
  await expect(page.locator('.tiptap')).toHaveText('KEEP THROUGH SAVE PANELS');
});
