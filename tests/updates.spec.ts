import { test, expect, type Page } from '@playwright/test';
import type { StudioUpdateState } from '../src/desktop';

async function updater(page: Page, options: Partial<StudioUpdateState> = {}) {
  await page.addInitScript(options => {
    let state = { phase: 'idle', percent: 0, version: null, currentVersion: '1.4.9', automatic: false,
      transferred: 0, total: 0, error: null, checkedAt: null, result: null, platform: 'win32', supported: true, packaged: true, ...options };
    let listener: ((next: any) => void) | undefined;
    let menu: ((action: string) => void) | undefined;
    const calls: string[] = [];
    (window as any).updateTest = {
      calls, push: (patch: any) => { state = { ...state, ...patch }; listener?.({ ...state }); },
      menu: (action: string) => menu?.(action),
    };
    (window as any).wordDesktop = {
      platform: state.platform, setTheme() {}, readRecovery: async () => [], saveRecovery: async () => true,
      onMenu: (callback: any) => { menu = callback; return () => { menu = undefined; }; },
      getUpdateState: async () => ({ ...state }),
      onUpdateState: (callback: any) => { listener = callback; return () => { listener = undefined; }; },
      updateAction: async (action: string, value?: boolean) => {
        calls.push(action);
        if (action === 'check') state = { ...state, phase: 'available', version: '1.5.0', result: 'available', total: 104857600, checkedAt: 1793422800000 };
        if (action === 'download') state = { ...state, phase: 'downloading', percent: 35, transferred: 36700160 };
        if (action === 'cancel') state = { ...state, phase: 'available', percent: 0 };
        if (action === 'automatic') state = { ...state, automatic: value };
        listener?.({ ...state }); return { ...state };
      },
    };
  }, options);
  await page.goto('/');
  await expect(page.locator('.tiptap[role="textbox"]')).toBeVisible();
}

test('file menu offers explicit check, download, progress, and manual installation', async ({ page }) => {
  await updater(page);
  const before = await page.locator('.tiptap[role="textbox"]').innerHTML();
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.locator('.backstage-nav').getByRole('button', { name: '수동 업데이트·설정', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Word Studio 업데이트', exact: true });
  await expect(dialog).toBeVisible();
  expect(await page.evaluate(() => (window as any).updateTest.calls)).toEqual([]);
  await expect(dialog.getByText('1.4.9', { exact: true })).toBeVisible();
  await dialog.getByRole('button', { name: '업데이트 확인', exact: true }).click();
  await expect(dialog.getByRole('button', { name: /^다운로드 \(/ })).toBeVisible();
  expect(await page.evaluate(() => (window as any).updateTest.calls)).toEqual(['check']);
  await dialog.getByRole('button', { name: /^다운로드 \(/ }).click();
  await expect(dialog.getByRole('progressbar', { name: '업데이트 다운로드' })).toHaveAttribute('value', '35');
  await expect(dialog.getByText('35.0 MB / 100.0 MB · 35%', { exact: true })).toBeVisible();
  await dialog.getByRole('button', { name: '업데이트 창 닫기', exact: true }).click();
  await expect(page.locator('.update-status')).toContainText('35%');
  await expect(page.locator('.tiptap[role="textbox"]')).toHaveJSProperty('innerHTML', before);
  await page.locator('.update-status').click();
  await expect(dialog).toBeVisible();
  await page.evaluate(() => (window as any).updateTest.push({ phase: 'verifying' }));
  await expect(dialog.getByText('다운로드 파일이 정상인지 확인하고 있습니다.', { exact: true })).toBeVisible();
  await page.evaluate(() => (window as any).updateTest.push({ phase: 'ready', percent: 100 }));
  await expect(dialog.getByRole('button', { name: '종료하고 설치', exact: true })).toBeVisible();
  await dialog.getByRole('button', { name: '다운로드 페이지 열기', exact: true }).click();
  expect(await page.evaluate(() => (window as any).updateTest.calls)).toEqual(['check', 'download', 'release-page']);
});
test('view and native menus open the same updater and native edit shortcuts cannot edit through its modal', async ({ page }) => {
  await updater(page);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '업데이트 확인', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Word Studio 업데이트', exact: true });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('checkbox')).not.toBeChecked();
  await dialog.getByRole('checkbox').check();
  await expect(dialog.getByRole('checkbox')).toBeChecked();
  const before = await page.locator('.tiptap[role="textbox"]').innerHTML();
  await page.evaluate(() => (window as any).updateTest.menu('new'));
  await page.keyboard.press('ControlOrMeta+n');
  await expect(page.locator('.tiptap[role="textbox"]')).toHaveJSProperty('innerHTML', before);
  await page.keyboard.press('Escape');
  await expect(dialog).not.toBeVisible();
  await page.evaluate(() => (window as any).updateTest.menu('updates'));
  await expect(dialog).toBeVisible();
});
test('automatic progress remains visible, cancellation and network failure can be retried', async ({ page }) => {
  await updater(page, { phase: 'downloading', version: '1.5.0', percent: 60, total: 104857600, transferred: 62914560, automatic: true });
  await page.locator('.update-status').click();
  const dialog = page.getByRole('dialog', { name: 'Word Studio 업데이트', exact: true });
  await dialog.getByRole('button', { name: '취소', exact: true }).click();
  await expect(dialog.getByRole('button', { name: /^다운로드 \(/ })).toBeVisible();
  await page.evaluate(() => (window as any).updateTest.push({ phase: 'error', error: '네트워크 연결 실패' }));
  await expect(dialog.getByRole('alert')).toHaveText('네트워크 연결 실패');
  await dialog.getByRole('button', { name: '업데이트 확인', exact: true }).click();
  expect(await page.evaluate(() => (window as any).updateTest.calls)).toEqual(['cancel', 'check']);
});
test('Mac shows replacement instructions rather than claiming to install automatically', async ({ page }) => {
  await updater(page, { platform: 'darwin', phase: 'ready', percent: 100, version: '1.5.0' });
  await page.locator('.update-status').click();
  const dialog = page.getByRole('dialog', { name: 'Word Studio 업데이트', exact: true });
  await expect(dialog.getByRole('button', { name: '다운로드 파일·교체 안내', exact: true })).toBeVisible();
  await expect(dialog.getByText('다운로드가 끝나도 Mac 앱이 자동으로 교체되지는 않습니다.', { exact: true })).toBeVisible();
});
test('Linux and development builds show manual installation without a false latest result', async ({ page }) => {
  await updater(page, { platform: 'linux', supported: false });
  await page.evaluate(() => (window as any).updateTest.menu('updates'));
  const dialog = page.getByRole('dialog', { name: 'Word Studio 업데이트', exact: true });
  await expect(dialog.getByRole('button', { name: '업데이트 확인', exact: true })).toBeDisabled();
  await expect(dialog.getByRole('checkbox')).toBeDisabled();
  await expect(dialog.getByText(/Linux: 다운로드 페이지/)).toBeVisible();
  await expect(dialog.getByText('최신 버전을 사용하고 있습니다.', { exact: true })).toHaveCount(0);
  await page.evaluate(() => (window as any).updateTest.push({ platform: 'darwin', supported: true, packaged: false }));
  await expect(dialog.getByText(/설치된 배포 앱에서 업데이트 확인/)).toBeVisible();
});
