import { test, expect } from '@playwright/test';

test('print preview reports its measured output count and restores the editor counter on return', async ({ page }) => {
  await page.goto('/');
  const html = '<p>FIRST</p><div data-page-break="true"></div><p>LAST</p>';
  await page.locator('input[type=file]').first().setInputFiles({ name: 'two-pages.html', mimeType: 'text/html', buffer: Buffer.from(html) });
  const status = page.locator('footer.statusbar > button').first();
  await expect(status).toContainText('/2 페이지');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-page')).toHaveCount(2);
  await expect(status).toHaveText('2 페이지 (인쇄)');
  await expect(status).toBeDisabled();
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  await expect(status).toBeEnabled();
  await expect(status).toContainText('/2 페이지');
});
