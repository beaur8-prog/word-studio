import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import path from 'node:path';

async function blank(page: Page) {
  await page.goto('/'); await page.getByRole('button', { name: '새 문서', exact: true }).click();
  const editor = page.getByLabel('Android 문서 편집 영역');
  await expect(editor).toHaveAttribute('contenteditable', 'true');
  await editor.click(); await page.keyboard.type('RIBBON CONTENT'); return editor;
}
const modes = (page: Page) => page.locator('.android-topbar .android-mode-switch');
async function resize(page: Page, size: { width: number; height: number }) {
  await page.setViewportSize(size);
  await expect.poll(() => page.evaluate(() => parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--android-height')))).toBe(size.height);
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
}

test('persistent top icons and Escape return to editing without losing text or undo history', async ({ page }) => {
  const editor = await blank(page);
  await page.keyboard.type(' KEEP EDIT');
  await expect(modes(page).getByRole('button', { name: '편집 모드', exact: true })).toHaveAttribute('aria-pressed', 'true');
  for (let round = 0; round < 3; round++) {
    await modes(page).getByRole('button', { name: '전체 화면 뷰어', exact: true }).click();
    const viewer = page.locator('.android-presentation');
    await expect(viewer.locator('.android-presentation-page')).toContainText('RIBBON CONTENT KEEP EDIT');
    await expect(viewer.getByRole('button', { name: '전체 화면 뷰어', exact: true })).toHaveAttribute('aria-pressed', 'true');
    await expect(viewer.locator('.android-presentation-controls')).toBeHidden({ timeout: 5000 });
    await expect(viewer.getByRole('button', { name: '편집 모드', exact: true })).toBeVisible();
    if (round === 1) await viewer.getByRole('button', { name: '편집 모드', exact: true }).click();
    else if (round === 2) await viewer.dispatchEvent('studio-presentation-edit');
    else await page.keyboard.press('Escape');
    await expect(viewer).toHaveCount(0); await expect(editor).toBeVisible();
    await expect(editor).toHaveAttribute('contenteditable', 'true'); await expect(editor).toContainText('KEEP EDIT');
    await expect(editor).toBeFocused();
  }
  await page.keyboard.type(' AFTER ESC'); await expect(editor).toContainText('AFTER ESC');
  await page.getByRole('button', { name: '실행 취소', exact: true }).click();
  await expect(editor).not.toContainText('AFTER ESC');
  await page.getByRole('button', { name: '다시 실행', exact: true }).click(); await expect(editor).toContainText('AFTER ESC');
  await editor.dispatchEvent('compositionstart', { data: '한' });
  await expect(modes(page).getByRole('button', { name: '전체 화면 뷰어', exact: true })).toBeDisabled();
  await editor.dispatchEvent('compositionend', { data: '한' });
  await expect(modes(page).getByRole('button', { name: '전체 화면 뷰어', exact: true })).toBeEnabled();
});

test('ribbon font controls apply to selected text and survive viewer, Escape and DOCX saving', async ({ page }) => {
  const editor = await blank(page); await page.keyboard.press('ControlOrMeta+a');
  await page.getByRole('combobox', { name: '글꼴', exact: true }).selectOption('Noto Serif KR');
  await page.getByRole('spinbutton', { name: '글꼴 크기', exact: true }).fill('13.5');
  await page.getByRole('button', { name: '굵게', exact: true }).click();
  await expect(editor.locator('strong')).toContainText('RIBBON CONTENT');
  await expect(editor.locator('span[style*="font-size"]')).toHaveCSS('font-size', '18px');
  await expect(editor.locator('span[style*="font-family"]')).toHaveCSS('font-family', '"Noto Serif KR"');
  await page.getByRole('button', { name: '가운데 정렬', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveCSS('text-align', 'center');
  await modes(page).getByRole('button', { name: '전체 화면 뷰어', exact: true }).click();
  await expect(page.locator('.android-presentation-page')).toContainText('RIBBON CONTENT');
  await page.keyboard.press('Escape');
  await expect(editor.locator('strong')).toContainText('RIBBON CONTENT');
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: '문서 사본 저장', exact: true }).click();
  const saved = await download, zip = await JSZip.loadAsync(await fs.readFile((await saved.path())!));
  const xml = await zip.file('word/document.xml')!.async('string');
  expect(xml).toContain('RIBBON CONTENT'); expect(xml).toContain('Noto Serif KR'); expect(xml).toContain('w:val="27"');
});

test('desktop-like tabs insert page breaks and tables; folding tools preserves a usable paper on phone and tablet', async ({ page }, testInfo) => {
  const editor = await blank(page); await page.keyboard.press('ControlOrMeta+End');
  await page.getByRole('tab', { name: '삽입', exact: true }).click();
  await page.getByRole('button', { name: '쪽 나누기', exact: true }).click();
  await expect(editor.locator('[data-page-break]')).toHaveCount(1);
  await page.getByRole('button', { name: '표 삽입', exact: true }).click();
  await page.getByRole('button', { name: '표 삽입 (3×3)', exact: true }).click();
  await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
  await expect(editor.locator('table')).toHaveCount(1); await expect(editor.locator('tr')).toHaveCount(3);
  await page.getByRole('tab', { name: '홈', exact: true }).click();
  for (const size of [{ width: 360, height: 740 }, { width: 412, height: 840 }, { width: 1280, height: 900 }]) {
    await resize(page, size);
    const mode = modes(page); const rect = (await mode.boundingBox())!;
    expect(rect.x).toBeGreaterThanOrEqual(0); expect(rect.x + rect.width).toBeLessThanOrEqual(size.width);
    const small = await page.locator('.android-topbar button, .android-editor-ribbon button, .android-editor-ribbon input, .android-editor-ribbon select').evaluateAll(elements => elements.filter(el => {
      const r = el.getBoundingClientRect(); return r.width < 47.5 || r.height < 47.5;
    }).map(el => el.getAttribute('aria-label') || el.textContent));
    expect(small).toEqual([]);
    expect(await page.locator('.android-shell').evaluate(el => el.scrollWidth <= el.clientWidth + 1)).toBe(true);
    expect((await page.locator('.android-editor-scroll').boundingBox())!.height).toBeGreaterThan(300);
    await page.locator('.android-editor-scroll').evaluate(el => { el.scrollTop = 0; });
    await page.screenshot({ path: testInfo.outputPath(`ribbon-${size.width}.png`) });
  }
  await resize(page, { width: 412, height: 740 });
  const before = (await page.locator('.android-editor-scroll').boundingBox())!.height;
  await page.getByRole('button', { name: '도구 모음 접기', exact: true }).click();
  await expect(page.locator('.android-ribbon-content')).toBeHidden();
  expect((await page.locator('.android-editor-scroll').boundingBox())!.height).toBeGreaterThan(before + 80);
  await expect(editor).toContainText('RIBBON CONTENT');
  await page.getByRole('button', { name: '도구 모음 펼치기', exact: true }).click();
  await resize(page, { width: 740, height: 360 });
  expect((await page.locator('.android-editor-scroll').boundingBox())!.height).toBeGreaterThan(70);
});

test('viewer page is retained across Escape and another icon switch; HWP Escape keeps conversion explicit', async ({ page }) => {
  await page.goto('/');
  const buffer = await Packer.toBuffer(new Document({ sections: [{ children: [1, 2, 3].map(i => new Paragraph({ text: `MODE PAGE ${i}`, pageBreakBefore: i > 1 })) }] }));
  await page.locator('input[type=file]').setInputFiles({ name: 'modes.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer });
  await modes(page).getByRole('button', { name: '전체 화면 뷰어', exact: true }).click();
  await page.getByRole('button', { name: '다음 페이지', exact: true }).click();
  await expect(page.locator('.android-presentation-page')).toContainText('MODE PAGE 2');
  await page.keyboard.press('Escape'); await expect(page.getByLabel('Android 문서 편집 영역')).toBeVisible();
  await modes(page).getByRole('button', { name: '전체 화면 뷰어', exact: true }).click();
  await expect(page.locator('.android-presentation')).toHaveAttribute('data-page', '2');
  await page.getByRole('button', { name: '뷰어 닫기', exact: true }).click();
  await page.locator('input[type=file]').setInputFiles(path.resolve('tests/fixtures/hwp/reader.hwpx'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByRole('button', { name: '한글 전체 화면 뷰어', exact: true }).click();
  await expect(reader.locator('.android-presentation-page svg')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(reader.locator('.android-presentation')).toHaveCount(0);
  await expect(reader.getByRole('region', { name: '한글 편집 변환 안내', exact: true })).toBeVisible();
  await expect(page.locator('[contenteditable=true]')).toHaveCount(0);
});
