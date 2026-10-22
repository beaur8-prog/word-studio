'use strict';
const { expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const path = require('node:path');
const assert = require('node:assert/strict');

// Actual renderer controls + native IPC/filesystem. Only the OS file chooser is
// replaced, using the enclosing smoke run's disposable profile and directory.
module.exports = async function workflowDesktopSmoke(app, window, directory) {
  const editor = window.locator('.document-stack .tiptap');
  const undo = window.getByRole('button', { name: '실행 취소 (Ctrl+Z)', exact: true });
  const redo = window.getByRole('button', { name: '다시 실행 (Ctrl+Y)', exact: true });
  const open = async file => {
    await app.evaluate(({ dialog }, name) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [name] }); }, file);
    await window.keyboard.press('ControlOrMeta+o');
    await expect(window.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(file, path.extname(file)), { timeout: 60000 });
  };
  const chooseSave = file => app.evaluate(({ dialog }, name) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: name });
  }, file);
  const short = path.join(directory, 'user-workflow.txt');
  await fs.writeFile(short, '원래 본문'); await open(short);
  await editor.fill('수정한 본문');
  await window.getByRole('button', { name: '레이아웃', exact: true }).click();
  await window.getByRole('button', { name: '용지 방향', exact: true }).click();
  await window.getByRole('button', { name: '가로', exact: true }).click();
  await undo.click();
  await expect(window.locator('.document-stack')).toHaveAttribute('data-page-height', '1122.52');
  await expect(editor).toHaveText('수정한 본문');
  await redo.click();
  await expect(window.locator('.document-stack')).toHaveAttribute('data-page-height', '793.7');
  await undo.click();
  await editor.click(); await window.keyboard.press('ControlOrMeta+a');
  await window.getByRole('button', { name: '검토', exact: true }).click();
  await window.getByRole('button', { name: '새 메모', exact: true }).click();
  await window.getByRole('textbox', { name: '메모 내용' }).fill('네이티브 메모 검증');
  await window.getByRole('button', { name: '메모 추가', exact: true }).click();
  await undo.click();
  await expect(window.locator('.comment-card')).toHaveCount(0);
  await expect(editor.locator('[data-comment-ids]')).toHaveCount(0);
  await redo.click();
  await expect(window.locator('.comment-card')).toContainText('네이티브 메모 검증');
  await expect(editor.locator('[data-comment-ids]')).toHaveCount(1);
  const copy = path.join(directory, 'user-workflow.wordstudio');
  await chooseSave(copy); await window.keyboard.press('ControlOrMeta+s');
  await expect(window.getByText('편집본 (.wordstudio)을 저장했습니다', { exact: true })).toBeVisible();
  const saved = JSON.parse(await fs.readFile(copy, 'utf8'));
  assert.equal(saved.comments[0].text, '네이티브 메모 검증');
  assert.equal(saved.page.orientation, 'portrait');
  await open(copy);
  await expect(undo).toBeDisabled(); await expect(redo).toBeDisabled();
  await expect(editor).toHaveText('수정한 본문');
  await expect(window.locator('.comment-card')).toContainText('네이티브 메모 검증');
  await window.getByRole('button', { name: '메모 창 닫기', exact: true }).click();

  const paragraphs = Array.from({ length: 1000 }, (_, i) => `LONG-${String(i + 1).padStart(4, '0')} ` +
    '한글과 English가 섞인 장문 보고서입니다. 모든 문단과 수정 내용을 저장하고 다시 열어 확인합니다. '.repeat(3));
  const large = path.join(directory, 'long-user-workflow.html');
  await fs.writeFile(large, paragraphs.map(text => `<p>${text}</p>`).join(''));
  const started = performance.now(); await open(large);
  const metrics = () => editor.locator(':scope > p').nth(1).evaluate(el => {
    const css = getComputedStyle(el), box = el.getBoundingClientRect(), range = document.createRange(); range.selectNodeContents(el);
    return { font: css.fontFamily, size: css.fontSize, line: css.lineHeight, after: css.marginBottom, width: box.width, height: box.height,
      runs: [...el.querySelectorAll('span')].map(s => ({ font: getComputedStyle(s).fontFamily, size: getComputedStyle(s).fontSize })),
      lines: [...range.getClientRects()].map(r => ({ top: r.top, width: r.width, height: r.height })) };
  });
  await expect(editor.locator(':scope > p')).toHaveCount(paragraphs.length);
  await expect.poll(async () => Number(await window.locator('.document-stack').getAttribute('data-pages'))).toBeGreaterThan(50);
  const openMs = Math.round(performance.now() - started);
  const initialPages = Number(await window.locator('.document-stack').getAttribute('data-pages'));
  const originalMetrics = await metrics();
  const first = editor.locator(':scope > p').first();
  const edited = 'EDITED ' + paragraphs[0];
  const editingStarted = performance.now();
  await first.fill(edited); await window.keyboard.press('ControlOrMeta+z');
  await expect(first).toHaveText(paragraphs[0]);
  await window.keyboard.press('ControlOrMeta+Shift+z'); await expect(first).toHaveText(edited);
  const editUndoRedoMs = Math.round(performance.now() - editingStarted);
  const expected = [edited, ...paragraphs.slice(1)];
  assert.deepEqual(await editor.locator(':scope > p').allTextContents(), expected);
  const docx = path.join(directory, 'long-user-workflow.docx');
  await chooseSave(docx);
  await window.getByRole('button', { name: '내보내기', exact: true }).click();
  // Check the export-panel toggle and restoration of editor focus as well.
  await window.getByRole('button', { name: '내보내기', exact: true }).click();
  await expect(editor).toBeFocused();
  await window.getByRole('button', { name: '내보내기', exact: true }).click();
  const saveStarted = performance.now();
  await window.getByRole('region', { name: '다른 형식으로 저장', exact: true }).getByRole('button', { name: /^Word 문서 \(\.docx\)/ }).click();
  await expect(window.getByText('DOCX 파일을 저장했습니다', { exact: true })).toBeVisible({ timeout: 60000 });
  const saveMs = Math.round(performance.now() - saveStarted);
  // Use a distinct filename so a title assertion cannot pass before import.
  const reopened = path.join(directory, 'long-reopened.docx'); await fs.copyFile(docx, reopened);
  const reopenStarted = performance.now(); await open(reopened);
  await expect(editor.locator(':scope > p')).toHaveCount(paragraphs.length, { timeout: 60000 });
  assert.deepEqual(await editor.locator(':scope > p').allTextContents(), expected);
  const reopenMs = Math.round(performance.now() - reopenStarted);
  await fs.writeFile(path.join(directory, 'long-document-metrics.json'), JSON.stringify({ originalMetrics, reopenedMetrics: await metrics(),
    initialPages, reopenedPages: Number(await window.locator('.document-stack').getAttribute('data-pages')) }, null, 2));
  await expect.poll(async () => Number(await window.locator('.document-stack').getAttribute('data-pages'))).toBe(initialPages);
  await expect(undo).toBeDisabled(); await expect(redo).toBeDisabled();
  await editor.locator(':scope > p').last().scrollIntoViewIfNeeded();
  await expect(editor.locator(':scope > p').last()).toBeInViewport();
  await window.screenshot({ path: path.join(directory, 'long-document-last-page.png') });
  return { metadataUndoRedo: true, commentAnchors: true, nativeSaveReopen: true, exportToggle: true,
    longDocument: { paragraphs: paragraphs.length, characters: expected.join('').length, initialPages,
      reopenedPages: Number(await window.locator('.document-stack').getAttribute('data-pages')),
      docxBytes: (await fs.stat(docx)).size, allParagraphsPreserved: true, openMs, editUndoRedoMs, saveMs, reopenMs } };
};
