import { test, expect } from '@playwright/test';

for (const control of ['slider', 'title'] as const) test(`cell selection takes keyboard focus back from the ${control}`, async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'table-focus.html', mimeType: 'text/html',
    buffer: Buffer.from('<p>LEAD</p><table><tr><td><p>FIRST CELL</p></td><td><p>TARGET CELL</p></td></tr></table><p>TAIL</p>') });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
  const target = editor.locator('td p').filter({ hasText: 'TARGET CELL' });
  const other = control === 'slider' ? page.getByRole('slider', { name: '확대 축소' }) : page.getByRole('textbox', { name: '문서 제목', exact: true });
  const before = await editor.evaluate(el => (el as any).editor.getJSON());
  // A toolbar interaction between clicks must not leave a selected cell with
  // keyboard focus on that toolbar. ProseMirror counts nearby clicks itself.
  const box = (await target.boundingBox())!;
  for (let click = 0; click < 3; click++) {
    await other.focus();
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  }
  await expect(editor.locator('td').nth(1)).toHaveClass(/selectedCell/);
  await expect(editor).toBeFocused();
  expect(await editor.evaluate(el => (el as any).editor.view.hasFocus())).toBe(true);
  expect(await editor.evaluate(el => (el as any).editor.getJSON())).toEqual(before);
  await page.keyboard.insertText('REPLACED');
  await expect(editor.locator('td').nth(1)).toHaveText('REPLACED');
  await expect(editor.locator('p').first()).toHaveText('LEAD');
  await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor.evaluate(el => (el as any).editor.getJSON())).toEqual(before);
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await expect(editor.locator('td').nth(1)).toHaveText('REPLACED');
  if (control === 'slider') await expect(other).toHaveValue('100');
});
