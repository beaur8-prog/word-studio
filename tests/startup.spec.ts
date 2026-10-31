import { test, expect } from '@playwright/test';
import { defaultPage, type StudioDocument } from '../src/types';
const prior: StudioDocument = { id: 'previous-work', format: 'word-studio', version: 1, title: '이전 작업',
  content: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'KEEP PREVIOUS WORK', marks: [{ type: 'bold' }] }] }] },
  page: { ...defaultPage, orientation: 'landscape', header: 'PREVIOUS HEADER' }, comments: [], updatedAt: '2026-10-30T00:00:00.000Z' };
const editor = (page: any) => page.locator('.tiptap[role="textbox"]');

test('fresh launches and relaunches start blank without adding unused documents to recent work', async ({ page }) => {
  await page.goto('/');
  for (let i = 0; i < 3; i++) {
    await expect(editor(page)).toHaveText('');
    await expect(page.getByRole('textbox', { name: '문서 제목' })).toHaveValue('문서1');
    await expect(page.getByText('새 문서', { exact: true }).first()).toBeVisible();
    // Give a mistakenly scheduled recovery debounce time to write a record.
    await page.waitForTimeout(800);
    await page.getByRole('button', { name: '파일', exact: true }).click();
    await expect(page.locator('.recent-document')).toHaveCount(0);
    await page.keyboard.press('Escape');
    await page.reload();
  }
});

test('a saved prior document remains available with formatting while a fresh page uses default settings', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'prior.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(prior)) });
  await expect(editor(page)).toHaveText('KEEP PREVIOUS WORK');
  await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
  await page.reload();
  await expect(editor(page)).toHaveText('');
  await expect(page.getByRole('textbox', { name: '문서 제목' })).toHaveValue('문서1');
  await expect(page.locator('.document-stack')).toHaveAttribute('data-page-height', '1122.52');
  await expect(page.locator('.paper-header').first()).toHaveText('');
  await expect(page.getByRole('button', { name: '실행 취소 (Ctrl+Z)', exact: true })).toBeDisabled();
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await expect(page.locator('.recent-document')).toHaveCount(1);
  await page.locator('.recent-document').filter({ hasText: '이전 작업' }).click();
  await expect(editor(page).locator('strong')).toHaveText('KEEP PREVIOUS WORK');
  await expect(page.locator('.paper-header').first()).toHaveText('PREVIOUS HEADER');
  await expect(page.locator('.document-stack')).toHaveAttribute('data-page-height', '793.7');
});

test('legacy current-document storage is migrated without opening it automatically', async ({ page }) => {
  await page.addInitScript(doc => localStorage.setItem('word-studio-current', JSON.stringify(doc)), prior);
  await page.goto('/');
  await expect(editor(page)).toHaveText('');
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.locator('.recent-document').filter({ hasText: '이전 작업' }).click();
  await expect(editor(page)).toHaveText('KEEP PREVIOUS WORK');
});

test('native recovery falls back to the previous valid version without replacing the blank startup', async ({ page }) => {
  await page.addInitScript(doc => {
    (window as any).wordDesktop = { platform: 'test', setTheme() {}, onMenu: () => () => {},
      readRecovery: async () => [{ current: '{', previous: JSON.stringify(doc), savedAt: doc.updatedAt }],
      saveRecovery: async () => true };
  }, prior);
  await page.goto('/');
  await expect(editor(page)).toHaveText('');
  await expect(page.getByText('최근 복구본을 읽지 못했습니다. 파일 메뉴의 최근 작업한 문서에서 이전 정상 버전을 열 수 있습니다.', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.locator('.recent-document').filter({ hasText: '이전 작업' }).click();
  await expect(editor(page).locator('strong')).toHaveText('KEEP PREVIOUS WORK');
});

test('editing and undoing back to blank still saves the latest intentional state', async ({ page }) => {
  await page.goto('/');
  await editor(page).fill('TEXT TO UNDO');
  await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: '실행 취소 (Ctrl+Z)', exact: true }).click();
  await expect(editor(page)).toHaveText('');
  await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await expect(page.locator('.recent-document')).toHaveCount(1);
  await page.locator('.recent-document').click();
  await expect(editor(page)).toHaveText('');
});

test('an explicitly saved blank document remains in recent work', async ({ page }) => {
  await page.goto('/');
  await expect(editor(page)).toHaveText('');
  const download = page.waitForEvent('download');
  await page.keyboard.press('ControlOrMeta+s'); await download;
  await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await expect(page.locator('.recent-document')).toHaveCount(1);
});
