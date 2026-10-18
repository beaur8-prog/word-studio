import { test, expect, type Page } from '@playwright/test';

async function table(page: Page) {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'cells.html', mimeType: 'text/html', buffer: Buffer.from(
    '<table>' + [1, 2, 3].map(row => `<tr>${['A', 'B', 'C'].map(col => `<td><p>${col}${row}</p></td>`).join('')}</tr>`).join('') + '</table><p>표 바깥</p>',
  ) });
  await expect(page.locator('.tiptap td')).toHaveCount(9);
}
async function fourCells(page: Page) {
  const cells = page.locator('.tiptap td');
  await cells.nth(0).click(); await cells.nth(4).click({ modifiers: ['Shift'] });
  await expect(page.locator('.tiptap .selectedCell')).toHaveCount(4);
  // Right click a selected cell that is neither the anchor nor the head.
  await cells.nth(1).click({ button: 'right' });
  await expect(page.getByRole('dialog', { name: '표 셀 편집' })).toBeVisible();
  await expect(page.getByText('4개 셀 선택됨', { exact: true })).toBeVisible();
}

test('multi-cell context formatting preserves the range, undo, file roundtrip and print styles', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  await table(page); await fourCells(page);
  const menu = page.getByRole('dialog', { name: '표 셀 편집' });
  await menu.getByLabel('셀 글꼴', { exact: true }).selectOption('바탕');
  await menu.getByLabel('셀 글꼴 크기').fill('18'); await menu.getByLabel('셀 글꼴 크기').press('Enter');
  await menu.getByRole('button', { name: '굵게', exact: true }).click();
  await menu.getByRole('button', { name: '글자 색', exact: true }).click();
  await menu.getByRole('button', { name: '글자 색 #c00000', exact: true }).click();
  await menu.getByRole('button', { name: '셀 배경', exact: true }).click();
  await menu.getByRole('button', { name: '셀 배경 #fff2cc', exact: true }).click();
  await menu.getByRole('button', { name: '가운데 정렬', exact: true }).click();
  await menu.getByLabel('셀 세로 정렬').selectOption('middle');
  await expect(page.locator('.tiptap .selectedCell')).toHaveCount(4);
  const cells = page.locator('.tiptap td');
  for (const i of [0, 1, 3, 4]) {
    await expect(cells.nth(i)).toHaveCSS('background-color', 'rgb(255, 242, 204)');
    await expect(cells.nth(i)).toHaveCSS('vertical-align', 'middle');
    await expect(cells.nth(i).locator('p')).toHaveCSS('text-align', 'center');
    await expect(cells.nth(i).locator('strong')).toHaveText(/[AB][12]/);
    await expect(cells.nth(i).locator('span').first()).toHaveCSS('color', 'rgb(192, 0, 0)');
  }
  await expect(cells.nth(2)).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
  await expect(cells.nth(2).locator('strong')).toHaveCount(0);
  await menu.getByRole('tab', { name: '테두리', exact: true }).click();
  await menu.getByLabel('테두리 선 종류').selectOption('double');
  await menu.getByLabel('테두리 굵기').selectOption('3');
  await menu.getByRole('button', { name: '테두리 색 #0070c0', exact: true }).click();
  await menu.getByRole('button', { name: '모든 테두리', exact: true }).click();
  await page.screenshot({ path: 'test-results/table-context-menu.png' });
  await expect(cells.nth(0)).toHaveCSS('border-top-style', 'double');
  await expect(cells.nth(0)).toHaveCSS('border-top-width', '4px');
  await menu.getByRole('button', { name: '테두리 없음', exact: true }).click();
  await expect(cells.nth(0)).toHaveCSS('border-top-style', 'none');
  await expect(cells.nth(2)).toHaveCSS('border-left-style', 'none');
  await page.keyboard.press('Escape'); await page.keyboard.press('ControlOrMeta+z');
  await expect(cells.nth(0)).toHaveCSS('border-top-style', 'double');
  await expect(cells.nth(0)).toHaveCSS('background-color', 'rgb(255, 242, 204)');
  await cells.nth(1).click({ button: 'right' });
  await menu.getByRole('tab', { name: '글자·셀 서식', exact: true }).click();
  await menu.getByRole('button', { name: '선택한 셀 서식 지우기', exact: true }).click();
  await expect(cells.nth(1)).toHaveCSS('border-right-width', '1px');
  await expect(cells.nth(2)).toHaveCSS('border-left-width', '1px');
  await page.keyboard.press('Escape'); await page.keyboard.press('ControlOrMeta+z');
  await expect(cells.nth(1)).toHaveCSS('border-right-style', 'double');
  await expect(cells.nth(0)).toHaveCSS('background-color', 'rgb(255, 242, 204)');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = await (await download).path();
  await page.locator('input[type=file]').first().setInputFiles(saved!);
  await expect(cells.nth(0)).toHaveCSS('border-top-style', 'double');
  await expect(cells.nth(0)).toHaveCSS('background-color', 'rgb(255, 242, 204)');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  await expect(page.locator('.print-layout.preview td').first()).toHaveCSS('border-top-style', 'double');
  await expect(page.locator('.print-layout.preview td').first()).toHaveCSS('background-color', 'rgb(255, 242, 204)');
  expect(errors).toEqual([]);
});

test('empty header cell defaults and no-fill survive DOCX, and paste fills the selected rectangle', async ({ page }) => {
  await table(page);
  const defaults = await page.evaluate(async () => {
    // @ts-expect-error Vite source module.
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    // @ts-expect-error Vite source module.
    const { editorExtensions } = await import('/src/extensions.ts');
    // @ts-expect-error Vite dependency module.
    const { generateJSON } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const source = { format: 'word-studio', version: 1, title: '빈 셀', updatedAt: new Date().toISOString(),
      page: { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', footer: '', pageNumbers: true }, comments: [],
      content: generateJSON('<table><tr><th style="background-color:transparent;font-family:바탕;font-size:18pt;color:#0070c0"><p></p></th></tr></table>', editorExtensions) };
    const result = await importFile(new File([await exportDocx(source)], 'empty.docx'));
    return generateJSON(result.html, editorExtensions).content[0].content[0].content[0];
  });
  expect(defaults.type).toBe('tableHeader');
  expect(defaults.attrs).toMatchObject({ backgroundColor: 'transparent', fontFamily: '바탕', fontSize: '18pt', textColor: 'rgb(0, 112, 192)' });
  await fourCells(page);
  await page.evaluate(() => {
    // Supply a local clipboard fixture without changing the user's OS clipboard.
    Object.defineProperty(navigator.clipboard, 'read', { configurable: true, value: async () => [{ types: ['text/html'],
      getType: async () => new Blob(['<table><tr><td style="background-color:#fff2cc"><p>붙임1</p></td><td><p>붙임2</p></td></tr><tr><td><p>붙임3</p></td><td><p>붙임4</p></td></tr></table>'], { type: 'text/html' }),
    }] });
  });
  await page.getByRole('dialog', { name: '표 셀 편집' }).getByRole('button', { name: '붙여넣기', exact: true }).click();
  await expect(page.locator('.tiptap td')).toHaveCount(9);
  for (const [i, text] of [[0, '붙임1'], [1, '붙임2'], [3, '붙임3'], [4, '붙임4'], [2, 'C1']] as const) await expect(page.locator('.tiptap td').nth(i)).toHaveText(text);
  await expect(page.locator('.tiptap td').first()).toHaveCSS('background-color', 'rgb(255, 242, 204)');
});

test('context table actions, keyboard opening, outside targeting and read mode', async ({ page }) => {
  await table(page); await fourCells(page);
  let menu = page.getByRole('dialog', { name: '표 셀 편집' });
  await menu.getByRole('tab', { name: '표 편집', exact: true }).click();
  await menu.getByRole('button', { name: '셀 병합', exact: true }).click();
  await expect(page.locator('.tiptap td')).toHaveCount(6);
  await expect(page.locator('.tiptap td').first()).toHaveAttribute('colspan', '2');
  await expect(page.locator('.tiptap td').first()).toHaveAttribute('rowspan', '2');
  await page.keyboard.press('Shift+F10');
  await expect(menu).toBeVisible();
  await menu.getByRole('button', { name: '병합 셀 분할', exact: true }).click();
  await expect(page.locator('.tiptap td')).toHaveCount(9);
  await page.locator('.tiptap td').nth(8).click({ button: 'right' });
  await expect(menu).toContainText('1개 셀 선택됨');
  await menu.getByRole('button', { name: '행 선택', exact: true }).click();
  await expect(page.locator('.tiptap .selectedCell')).toHaveCount(3);
  await page.keyboard.press('Shift+F10');
  await menu.getByRole('button', { name: '아래에 행 삽입', exact: true }).click();
  await expect(page.locator('.tiptap tr')).toHaveCount(4);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.tiptap tr')).toHaveCount(3);
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '읽기 모드', exact: true }).click();
  await page.locator('.tiptap td').first().click({ button: 'right' });
  await expect(menu).toHaveCount(0);
});

test('DOCX preserves direct cell borders, fill, padding and run fonts through two conversions', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    // @ts-expect-error Vite source module.
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    // @ts-expect-error Vite source module.
    const { editorExtensions } = await import('/src/extensions.ts');
    // @ts-expect-error Vite dependency module.
    const { generateJSON } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const html = '<table><tr><td style="background-color:#fff2cc;border-top:3pt double #0070c0;border-right:1.5pt dashed #c00000;border-bottom:0px none #000000;border-left:1pt dotted #000000;vertical-align:middle;padding:12px;font-family:바탕;font-size:18pt;color:#c00000"><p style="text-align:center"><strong>반복</strong></p></td><td><p>반복</p></td></tr></table>';
    let content = generateJSON(html, editorExtensions);
    const source = { format: 'word-studio', version: 1, title: '셀 서식', updatedAt: new Date().toISOString(),
      page: { size: 'A4', orientation: 'portrait', margin: 25.4, header: '', footer: '', pageNumbers: true }, comments: [] };
    let loaded;
    for (let i = 0; i < 2; i++) {
      loaded = await importFile(new File([await exportDocx({ ...source, content })], 'cells.docx'));
      content = generateJSON(loaded.html, editorExtensions);
    }
    const template = document.createElement('template'); template.innerHTML = loaded.html;
    const cell = template.content.querySelector<HTMLElement>('td,th');
    if (!cell) throw new Error(loaded.html);
    const run = cell.querySelector('span')!;
    return { style: { fill: cell.style.backgroundColor, top: cell.style.borderTopStyle, width: cell.style.borderTopWidth,
      right: cell.style.borderRightStyle, bottom: cell.style.borderBottomStyle, left: cell.style.borderLeftStyle,
      align: cell.style.verticalAlign, padding: cell.style.paddingTop, font: run.style.fontFamily, size: run.style.fontSize, color: run.style.color }, content };
  });
  expect(result.style).toEqual({ fill: 'rgb(255, 242, 204)', top: 'double', width: '3pt', right: 'dashed', bottom: 'none', left: 'dotted',
    align: 'middle', padding: '12px', font: '바탕', size: '18pt', color: 'rgb(192, 0, 0)' });
  const cells = result.content.content[0].content[0].content;
  expect(cells[0].type).toBe('tableCell');
  expect(cells[1].type).toBe('tableCell');
  expect(cells[1].content[0].content[0].marks ?? []).toHaveLength(0);
});

test('splitting ordinary cells preserves neighbours and empty-cell font defaults', async ({ page }) => {
  await table(page); await fourCells(page);
  const menu = page.getByRole('dialog', { name: '표 셀 편집' });
  await menu.getByRole('tab', { name: '표 편집', exact: true }).click();
  await menu.getByLabel('나눌 행 수').fill('2');
  await menu.getByLabel('나눌 열 수').fill('2');
  await menu.getByRole('button', { name: '각 셀 나누기', exact: true }).click();
  await expect(page.locator('.tiptap td')).toHaveCount(21);
  await expect(page.locator('.tiptap tr')).toHaveCount(5);
  for (const text of ['A1', 'A2', 'A3', 'B1', 'B2', 'B3', 'C1', 'C2', 'C3']) await expect(page.locator('.tiptap td').getByText(text, { exact: true })).toHaveCount(1);
  await expect(page.locator('.tiptap td').filter({ hasText: 'C1' })).toHaveAttribute('rowspan', '2');
  await expect(page.locator('.tiptap td').filter({ hasText: 'A3' })).toHaveAttribute('colspan', '2');
  await page.keyboard.press('ControlOrMeta+z'); await expect(page.locator('.tiptap td')).toHaveCount(9);
  await page.keyboard.press('ControlOrMeta+y'); await expect(page.locator('.tiptap td')).toHaveCount(21);
  const empty = page.locator('.tiptap td').filter({ hasText: /^$/ }).first();
  await empty.click({ button: 'right' });
  await menu.getByRole('tab', { name: '글자·셀 서식', exact: true }).click();
  await menu.getByLabel('셀 글꼴 크기').fill('16'); await menu.getByLabel('셀 글꼴 크기').press('Enter');
  await menu.getByRole('button', { name: '글자 색', exact: true }).click();
  await menu.getByRole('button', { name: '글자 색 #0070c0', exact: true }).click();
  await page.keyboard.press('Escape'); await empty.click(); await page.keyboard.insertText('새 글자');
  const typed = page.locator('.tiptap td').filter({ hasText: '새 글자' });
  await expect(typed).toHaveCSS('color', 'rgb(0, 112, 192)');
  await expect(typed).toHaveCSS('font-size', '21.3333px');
  await expect(page.locator('.tiptap td').filter({ hasText: 'C3' })).not.toHaveCSS('color', 'rgb(0, 112, 192)');
  await page.setViewportSize({ width: 1100, height: 700 });
  await typed.click({ button: 'right' });
  await menu.getByRole('tab', { name: '표 편집', exact: true }).click();
  const box = await menu.boundingBox();
  expect(box!.x).toBeGreaterThanOrEqual(0); expect(box!.y).toBeGreaterThanOrEqual(0);
  expect(box!.x + box!.width).toBeLessThanOrEqual(1100); expect(box!.y + box!.height).toBeLessThanOrEqual(700);
});
