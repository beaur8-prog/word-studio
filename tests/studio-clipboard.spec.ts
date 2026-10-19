import {test,expect} from '@playwright/test';
import fs from 'node:fs';
const rich=JSON.parse(fs.readFileSync('tests/fixtures/studio-table.json','utf8'));
test('Home paste retains an editable Sheet Studio table, merges, text and borders',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'clipboard-empty.html',mimeType:'text/html',buffer:Buffer.from('<p>붙여넣기 위치</p>')});await expect(page.locator('.tiptap')).toHaveText('붙여넣기 위치');await page.locator('.tiptap p').click();await page.keyboard.press('End');
 await page.evaluate(rich=>{Object.defineProperty(navigator.clipboard,'read',{configurable:true,value:async()=>[{types:['text/html','text/plain'],getType:async(type:string)=>new Blob([type==='text/html'?rich.html:rich.text],{type})}]});},rich);
 await page.getByRole('button',{name:'붙여넣기',exact:true}).first().click();
 await expect(page.locator('.tiptap table')).toHaveCount(1);const cells=page.locator('.tiptap td');await expect(cells).toHaveCount(3);await expect(cells.first()).toHaveAttribute('colspan','2');await expect(cells.first()).toHaveCSS('background-color','rgb(255, 242, 204)');await expect(cells.first()).toHaveCSS('border-bottom-color','rgb(255, 0, 0)');await expect(cells.first()).toContainText('보고서');await expect(cells.nth(1)).toContainText('줄1');await expect(cells.nth(1)).toContainText('줄2');await expect(cells.nth(2)).toHaveText('=literal');
 await page.keyboard.press('ControlOrMeta+z');await expect(page.locator('.tiptap table')).toHaveCount(0);await page.keyboard.press('ControlOrMeta+Shift+z');await expect(page.locator('.tiptap table')).toHaveCount(1);
 // A normal keyboard paste goes through the same sanitizer and table schema.
 await page.locator('.tiptap').click();await page.keyboard.press('ControlOrMeta+a');await page.evaluate(rich=>{const data=new DataTransfer();data.setData('text/html',rich.html);data.setData('text/plain',rich.text);document.querySelector('.tiptap')!.dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));},rich);await expect(page.locator('.tiptap table')).toHaveCount(1);
 const documentHtml=await page.locator('.tiptap').innerHTML();expect(documentHtml).toContain('colspan="2"');await expect(page.locator('.tiptap td').first()).toHaveCSS('border-bottom-width','3px');expect(await page.locator('.tiptap table').first().evaluate(el=>el.getBoundingClientRect().width)).toBeCloseTo(250,0);
 await page.screenshot({path:'test-results/studio-table-paste.png'});
});

test('a delayed Home paste cannot change a document after switching to read mode', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'paste-target.html', mimeType: 'text/html', buffer: Buffer.from('<p>원래 문서</p>'),
  });
  await expect(page.locator('.tiptap')).toHaveText('원래 문서');
  await page.locator('.tiptap p').click();
  await page.evaluate(rich => {
    Object.defineProperty(navigator.clipboard, 'read', { configurable: true, value: () => new Promise(resolve => {
      (window as any).finishClipboardRead = () => resolve([{
        types: ['text/html', 'text/plain'],
        getType: async (type: string) => new Blob([type === 'text/html' ? rich.html : rich.text], { type }),
      }]);
    }) });
  }, rich);
  await page.getByRole('button', { name: '붙여넣기', exact: true }).first().click();
  await page.getByRole('button', { name: '읽기 모드 전환', exact: true }).click();
  await expect(page.locator('.tiptap')).toHaveAttribute('contenteditable', 'false');
  await page.evaluate(async () => {
    (window as any).finishClipboardRead();
    await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
  });
  await expect(page.locator('.tiptap table')).toHaveCount(0);
  await expect(page.locator('.tiptap')).toHaveText('원래 문서');
  await expect(page.getByRole('status')).toContainText('다시 붙여넣으세요');
});

test('malformed clipboard table spans are rejected before the selection changes', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'paste-target.html', mimeType: 'text/html', buffer: Buffer.from('<p>보존할 문서</p>'),
  });
  await expect(page.locator('.tiptap')).toHaveText('보존할 문서');
  await page.locator('.tiptap').click();
  await page.keyboard.press('ControlOrMeta+a');
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/html', '<table><tr><td rowspan="2">잘못된 병합</td></tr></table>');
    data.setData('text/plain', '잘못된 병합');
    document.querySelector('.tiptap')!.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true, cancelable: true, clipboardData: data,
    }));
  });
  await expect(page.locator('.tiptap')).toHaveText('보존할 문서');
  await expect(page.locator('.tiptap table')).toHaveCount(0);
  await expect(page.getByRole('status')).toContainText('병합 범위');
  expect(errors).toEqual([]);
});

for (const change of ['focus', 'selection', 'document'] as const) {
  test(`a delayed Home paste is cancelled after a ${change} change`, async ({ page }) => {
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles({
      name: 'paste-target.html', mimeType: 'text/html', buffer: Buffer.from('<p>원래 문서</p>'),
    });
    await expect(page.locator('.tiptap')).toHaveText('원래 문서');
    await page.locator('.tiptap p').click();
    await page.keyboard.press('End');
    await page.evaluate(html => {
      Object.defineProperty(navigator.clipboard, 'read', { configurable: true, value: () => new Promise(resolve => {
        (window as any).finishClipboardRead = () => resolve([{
          types: ['text/html'], getType: async () => new Blob([html], { type: 'text/html' }),
        }]);
      }) });
    }, rich.html);
    await page.getByRole('button', { name: '붙여넣기', exact: true }).first().click();
    if (change === 'focus') await page.getByRole('textbox', { name: '문서 제목', exact: true }).focus();
    if (change === 'selection') {
      // Returning to the original cursor position must not revive an old read.
      await page.keyboard.press('ArrowLeft');
      await page.evaluate(() => new Promise<void>(resolve => requestAnimationFrame(() => resolve())));
      await page.keyboard.press('ArrowRight');
    }
    if (change === 'document') {
      await page.locator('input[type=file]').first().setInputFiles({
        name: 'next-document.html', mimeType: 'text/html', buffer: Buffer.from('<p>다음 문서</p>'),
      });
      await expect(page.locator('.tiptap')).toHaveText('다음 문서');
    }
    await page.evaluate(() => (window as any).finishClipboardRead());
    await expect(page.getByRole('status')).toContainText('다시 붙여넣으세요');
    await expect(page.locator('.tiptap table')).toHaveCount(0);
    await expect(page.locator('.tiptap')).toHaveText(change === 'document' ? '다음 문서' : '원래 문서');
    if (change === 'focus') await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toBeFocused();
  });
}

test('clipboard table limits count UTF-8 bytes and reject overlapping or excessive grids', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { clipboardHtml } = await import('/src/clipboard.ts');
    const rejected = [
      '<table><tr><td colspan="1001">too wide</td></tr></table>',
      '<table><tr><td>a</td><td rowspan="2">b</td></tr><tr><td colspan="2">overlap</td></tr></table>',
      '<table>' + '<tr><td colspan="1000">wide</td></tr>'.repeat(101) + '</table>',
      '<table><tr><td>' + '가'.repeat(Math.ceil(10 * 1024 * 1024 / 3)) + '</td></tr></table>',
    ].map(html => { try { clipboardHtml(html); return ''; } catch (error) { return (error as Error).message; } });
    const valid = clipboardHtml('<table><tr><td><p>일반 Word 표</p><table><tr><td>중첩 표</td></tr></table></td></tr></table>');
    return { rejected, valid };
  });
  expect(result.rejected[0]).toContain('병합 범위');
  expect(result.rejected[1]).toContain('겹칩니다');
  expect(result.rejected[2]).toContain('100,000셀');
  expect(result.rejected[3]).toContain('10 MB');
  expect(result.valid).toContain('일반 Word 표');
  expect(result.valid).toContain('중첩 표');
});

test('Home paste falls back to plain text when rich clipboard reading is unavailable', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'paste-target.html', mimeType: 'text/html', buffer: Buffer.from('<p>원래 문서</p>'),
  });
  await expect(page.locator('.tiptap')).toHaveText('원래 문서');
  await page.locator('.tiptap p').click();
  await page.keyboard.press('ControlOrMeta+a');
  await page.evaluate(() => {
    Object.defineProperty(navigator.clipboard, 'read', { configurable: true, value: undefined });
    Object.defineProperty(navigator.clipboard, 'readText', { configurable: true, value: async () => '<표 아님>\n다음 줄' });
  });
  await page.getByRole('button', { name: '붙여넣기', exact: true }).first().click();
  await expect(page.locator('.tiptap')).toContainText('<표 아님>');
  await expect(page.locator('.tiptap')).toContainText('다음 줄');
  await expect(page.locator('.tiptap table')).toHaveCount(0);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(page.locator('.tiptap')).toHaveText('원래 문서');
});

test('Word table copy sends quoted TSV with the rich HTML and preserves ordinary text copy', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'copy-table.html', mimeType: 'text/html',
    buffer: Buffer.from(rich.html.replace('=literal', '문자&quot;Q&quot;\tX')),
  });
  await page.locator('.tiptap td').first().click({ button: 'right' });
  const menu = page.getByRole('dialog', { name: '표 셀 편집' });
  await menu.getByRole('tab', { name: '표 편집', exact: true }).click();
  await menu.getByRole('button', { name: '표 전체 선택', exact: true }).click();
  const copied = await page.evaluate(() => {
    const data = new DataTransfer();
    document.querySelector('.tiptap')!.dispatchEvent(new ClipboardEvent('copy', {
      bubbles: true, cancelable: true, clipboardData: data,
    }));
    return { text: data.getData('text/plain'), html: data.getData('text/html') };
  });
  expect(copied.text).toBe('보고서\t\n"줄1\n줄2"\t"문자""Q""\tX"');
  expect(copied.html).toContain('<table');
  expect(copied.html).toContain('colspan="2"');
  expect(copied.html).toContain('줄1');

  await page.locator('input[type=file]').first().setInputFiles({
    name: 'copy-text.html', mimeType: 'text/html', buffer: Buffer.from('<p>첫 문단</p><p>둘째 문단</p>'),
  });
  await expect(page.locator('.tiptap')).toHaveText('첫 문단둘째 문단');
  await page.locator('.tiptap p').first().click();
  await page.keyboard.press('ControlOrMeta+a');
  const plain = await page.evaluate(() => {
    const data = new DataTransfer();
    document.querySelector('.tiptap')!.dispatchEvent(new ClipboardEvent('copy', {
      bubbles: true, cancelable: true, clipboardData: data,
    }));
    return data.getData('text/plain');
  });
  expect(plain).toBe('첫 문단\n\n둘째 문단');
});
