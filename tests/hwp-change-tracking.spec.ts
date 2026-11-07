import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
const CFB = createRequire(import.meta.url)('cfb');

test('tracked HWP remains viewable but cannot silently become an unmarked Word document; next ordinary file still edits', async ({ page }) => {
  const original = await fs.readFile('tests/fixtures/hwp/landscape-pagination.hwp');
  const cfb = CFB.read(original, { type: 'buffer' });
  const header = cfb.FileIndex[cfb.FullPaths.findIndex((p: string) => p.endsWith('/FileHeader'))];
  header.content = Buffer.from(header.content); header.content[37] |= 0x40;
  const tracked = CFB.write(cfb, { type: 'buffer' });
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'tracked.hwp', mimeType: 'application/octet-stream', buffer: tracked });
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByRole('alert')).toContainText('변경 내용 추적');
  await expect(reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true })).toBeDisabled();
  await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', '1쪽');
  const result = await page.evaluate(async bytes => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const session = new HwpReaderSession();
    try {
      const info = await session.open(new File([new Uint8Array(bytes)], 'tracked.hwp'));
      let blocked = '';
      try { await session.editable(); } catch (error) { blocked = String(error); }
      return { reason: info.conversionBlock, blocked, text: await session.pageText(0) };
    } finally { session.close(); }
  }, [...tracked]);
  expect(result.reason).toContain('변경 내용 추적');
  expect(result.blocked).toContain(result.reason!);
  expect(result.text.trim().length).toBeGreaterThan(0);
  await reader.getByRole('button', { name: 'HWP 읽기 닫기', exact: true }).click();
  await page.locator('input[type=file]').first().setInputFiles({ name: 'ordinary.hwp', mimeType: 'application/octet-stream', buffer: original });
  await expect(reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true })).toBeEnabled();
  await expect(reader.getByRole('alert')).toHaveCount(0);
});
