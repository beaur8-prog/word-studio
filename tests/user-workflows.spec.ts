import { test, expect, type Page } from '@playwright/test';

const edit = (page: Page) => page.locator('.document-stack .tiptap');
const undo = (page: Page) => page.getByRole('button', { name: '실행 취소 (Ctrl+Z)', exact: true });
const redo = (page: Page) => page.getByRole('button', { name: '다시 실행 (Ctrl+Y)', exact: true });
async function openText(page: Page, text = '원래 본문') {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'workflow.txt', mimeType: 'text/plain', buffer: Buffer.from(text) });
  await expect(edit(page)).toHaveText(text);
}

test('page setup can be undone and redone without removing the preceding text edit', async ({ page }) => {
  await openText(page);
  await edit(page).fill('원래 본문에 추가한 내용');
  await page.getByRole('button', { name: '레이아웃', exact: true }).click();
  await page.getByRole('button', { name: '용지 방향', exact: true }).click();
  await page.getByRole('button', { name: '가로', exact: true }).click();
  await expect(page.locator('.document-stack')).toHaveAttribute('data-page-height', '793.7');
  await undo(page).click();
  await expect(page.locator('.document-stack')).toHaveAttribute('data-page-height', '1122.52');
  await expect(edit(page)).toHaveText('원래 본문에 추가한 내용');
  await redo(page).click();
  await expect(page.locator('.document-stack')).toHaveAttribute('data-page-height', '793.7');
  await undo(page).click(); await undo(page).click();
  await expect(edit(page)).toHaveText('원래 본문');
});

test('comment creation is undone and redone together with its text anchor', async ({ page }) => {
  await openText(page);
  await edit(page).click(); await page.keyboard.press('ControlOrMeta+a');
  await page.getByRole('button', { name: '검토', exact: true }).click();
  await page.getByRole('button', { name: '새 메모', exact: true }).click();
  await page.getByRole('textbox', { name: '메모 내용' }).fill('검토 메모');
  await page.getByRole('button', { name: '메모 추가', exact: true }).click();
  await expect(page.locator('.comment-card')).toHaveCount(1);
  await expect(edit(page).locator('[data-comment-ids]')).toHaveCount(1);
  await undo(page).click();
  await expect(page.locator('.comment-card')).toHaveCount(0);
  await expect(edit(page).locator('[data-comment-ids]')).toHaveCount(0);
  await redo(page).click();
  await expect(page.locator('.comment-card')).toContainText('검토 메모');
  await expect(edit(page).locator('[data-comment-ids]')).toHaveCount(1);
});

test('comment resolve and delete keep records and anchors together through undo, redo and DOCX reopen', async ({ page }, info) => {
  await openText(page);
  await edit(page).click(); await page.keyboard.press('ControlOrMeta+a');
  await page.getByRole('button', { name: '검토', exact: true }).click();
  await page.getByRole('button', { name: '새 메모', exact: true }).click();
  await page.getByRole('textbox', { name: '메모 내용' }).fill('확인할 내용');
  await page.getByRole('button', { name: '메모 추가', exact: true }).click();
  const card = page.locator('.comment-card');
  await card.getByRole('button', { name: '해결', exact: true }).click();
  await expect(card).toHaveClass(/resolved/);
  await undo(page).click(); await expect(card).not.toHaveClass(/resolved/);
  await redo(page).click(); await expect(card).toHaveClass(/resolved/);
  await card.getByRole('button', { name: '메모 삭제', exact: true }).click();
  await expect(card).toHaveCount(0); await expect(edit(page).locator('[data-comment-ids]')).toHaveCount(0);
  await undo(page).click();
  await expect(card).toHaveClass(/resolved/); await expect(edit(page).locator('[data-comment-ids]')).toHaveCount(1);
  await redo(page).click(); await expect(card).toHaveCount(0);
  await undo(page).click();
  await page.getByRole('button', { name: '파일', exact: true }).click();
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
  const file = info.outputPath('review.docx'); await (await downloaded).saveAs(file);
  await page.locator('input[type=file]').first().setInputFiles(file);
  await expect(card).toContainText('확인할 내용');
  await expect(edit(page).locator('[data-comment-ids]')).toHaveCount(1);
  await expect(undo(page)).toBeDisabled();
});

test('document-wide comments and header/page settings are undoable even without a body edit', async ({ page }, info) => {
  await openText(page);
  await page.getByRole('button', { name: '검토', exact: true }).click();
  await page.getByRole('button', { name: '새 메모', exact: true }).click();
  await page.getByRole('textbox', { name: '메모 내용' }).fill('문서 전체 메모');
  await page.getByRole('button', { name: '메모 추가', exact: true }).click();
  await expect(edit(page).locator('[data-comment-ids]')).toHaveCount(0);
  await undo(page).click(); await expect(page.locator('.comment-card')).toHaveCount(0);
  await redo(page).click(); await expect(page.locator('.comment-card')).toContainText('문서 전체 메모');
  await page.getByRole('button', { name: '삽입', exact: true }).click();
  await page.getByRole('button', { name: '머리글·바닥글', exact: true }).click();
  await page.getByRole('textbox', { name: '머리글', exact: true }).fill('검수 보고서');
  await page.getByRole('button', { name: '완료', exact: true }).click();
  await expect(page.locator('.paper-header').first()).toHaveText('검수 보고서');
  await edit(page).click(); await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.paper-header').first()).toHaveText('');
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await expect(page.locator('.paper-header').first()).toHaveText('검수 보고서');
  const downloaded = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const file = info.outputPath('metadata.wordstudio'); await (await downloaded).saveAs(file);
  await page.locator('input[type=file]').first().setInputFiles(file);
  await expect(page.locator('.paper-header').first()).toHaveText('검수 보고서');
  await expect(page.locator('.comment-card')).toContainText('문서 전체 메모');
  await expect(undo(page)).toBeDisabled(); await expect(redo(page)).toBeDisabled();
  await edit(page).fill('재열기 후 수정');
  await undo(page).click(); await expect(edit(page)).toHaveText('원래 본문');
  await expect(page.locator('.paper-header').first()).toHaveText('검수 보고서');
});

test('read mode cannot change body or history using editing keyboard shortcuts', async ({ page }) => {
  await openText(page);
  await edit(page).fill('수정한 본문');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '읽기 모드', exact: true }).click();
  await edit(page).click();
  for (const key of ['ControlOrMeta+a', 'ControlOrMeta+b', 'ControlOrMeta+z', 'ControlOrMeta+y', 'Tab', 'Backspace']) await page.keyboard.press(key);
  await expect(edit(page)).toHaveText('수정한 본문');
  await expect(edit(page).locator('strong,[data-docx-tab]')).toHaveCount(0);
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  await undo(page).click(); await expect(edit(page)).toHaveText('원래 본문');
});

test('invalid files preserve the edited document and its undo/redo branches', async ({ page }) => {
  await openText(page);
  await edit(page).fill('보존할 수정 내용');
  for (const [name, text] of [['bad.wordstudio', '{bad json'], ['bad.docx', 'not a zip']]) {
    await page.locator('input[type=file]').first().setInputFiles({ name, mimeType: 'application/octet-stream', buffer: Buffer.from(text) });
    await expect(page.getByText(/파일을 열지 못했습니다:/)).toBeVisible();
    await expect(edit(page)).toHaveText('보존할 수정 내용');
  }
  await undo(page).click(); await expect(edit(page)).toHaveText('원래 본문');
  await redo(page).click(); await expect(edit(page)).toHaveText('보존할 수정 내용');
});

test('keyboard edits, list structure, replace all and history branching survive file saves', async ({ page }, info) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await openText(page, '업무 메모');
  await edit(page).click(); await page.keyboard.press('ControlOrMeta+a');
  await page.keyboard.press('ControlOrMeta+b');
  await expect(edit(page).locator('strong')).toHaveText('업무 메모');
  // Allow native selectionchange to reach the editor before the next key.
  await page.keyboard.press('ArrowRight', { delay: 60 });
  await page.keyboard.press('Enter', { delay: 60 }); await page.keyboard.press('ControlOrMeta+b');
  await expect(edit(page).locator('p').first()).toHaveText('업무 메모');
  await page.keyboard.type('- '); await page.keyboard.insertText('첫 번째 항목');
  await page.keyboard.press('Enter'); await page.keyboard.insertText('두 번째 항목');
  await expect(edit(page).locator('li')).toHaveCount(2);
  await page.keyboard.press('Enter'); await page.keyboard.press('Enter');
  await page.keyboard.insertText('검토 대상과 검토 대상');
  await page.getByRole('button', { name: '바꾸기', exact: true }).click();
  await page.getByRole('textbox', { name: '문서에서 찾기' }).fill('검토 대상');
  await page.getByRole('textbox', { name: '바꿀 내용' }).fill('완료 항목');
  await page.getByRole('button', { name: '모두 바꾸기', exact: true }).click();
  await expect(edit(page)).toContainText('완료 항목과 완료 항목');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const file = info.outputPath('workflow.wordstudio'); await (await download).saveAs(file);
  await undo(page).click(); await expect(edit(page)).toContainText('검토 대상과 검토 대상');
  await redo(page).click(); await expect(edit(page)).toContainText('완료 항목과 완료 항목');
  await undo(page).click();
  await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowDown' : 'Control+End');
  await page.keyboard.insertText(' 새로운 수정');
  await expect(redo(page)).toBeDisabled();
  await page.locator('input[type=file]').first().setInputFiles(file);
  await expect(edit(page)).toContainText('완료 항목과 완료 항목');
  await expect(edit(page)).not.toContainText('새로운 수정');
  await expect(edit(page).locator('li')).toHaveCount(2);
  await expect(edit(page).locator('strong')).toHaveText('업무 메모');
  await expect(undo(page)).toBeDisabled(); await expect(redo(page)).toBeDisabled();
  expect(errors).toEqual([]);
});

test('authored paragraph spacing and long-document page count survive two DOCX saves', async ({ page }, info) => {
  const texts = Array.from({ length: 240 }, (_, i) => `문단 ${i + 1}. ` + '한글과 English가 섞인 장문 보고서의 줄 간격과 쪽 수를 확인합니다. '.repeat(3));
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'long.html', mimeType: 'text/html', buffer: Buffer.from(texts.map(t => `<p>${t}</p>`).join('')) });
  await expect(edit(page).locator(':scope > p')).toHaveCount(texts.length);
  await page.evaluate(() => document.fonts.ready);
  await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).not.toBe('1');
  const pages = await page.locator('.document-stack').getAttribute('data-pages');
  const metrics = () => edit(page).locator(':scope > p').first().evaluate(el => {
    const css = getComputedStyle(el); return { font: css.fontFamily, line: parseFloat(css.lineHeight), after: parseFloat(css.marginBottom), height: el.getBoundingClientRect().height };
  });
  const original = await metrics();
  for (let pass = 1; pass <= 2; pass++) {
    await page.getByRole('button', { name: '파일', exact: true }).click();
    const downloaded = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서로 내보내기 DOCX', exact: true }).click();
    const file = info.outputPath(`long-${pass}.docx`); await (await downloaded).saveAs(file);
    await page.locator('input[type=file]').first().setInputFiles(file);
    await expect(page.getByRole('textbox', { name: '문서 제목' })).toHaveValue(`long-${pass}`);
    await expect.poll(async () => (await metrics()).line).toBeCloseTo(original.line, 1);
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe(pages);
    expect((await metrics()).after).toBeCloseTo(original.after, 1);
    expect((await metrics()).font).toBe(original.font);
    expect((await metrics()).height).toBeCloseTo(original.height, 1);
    expect(await edit(page).locator(':scope > p').allTextContents()).toEqual(texts);
  }
});
