import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { Document, Paragraph, Packer } from 'docx';

type RequestedFile = { name: string; data: number[] };
const textFile = (name: string, text: string): RequestedFile => ({ name, data: [...Buffer.from(text)] });
const editor = (page: Page) => page.locator('.document-stack .tiptap');
async function desktop(page: Page, startup: RequestedFile[] = [], delayStartup = false) {
  await page.addInitScript(({ startup, delayStartup }) => {
    const state = {
      queue: startup, writes: [] as string[], fail: false, holdRead: false,
      releaseRead: null as null | (() => void), releaseStartup: null as null | (() => void),
      notify: null as null | (() => void),
    };
    (window as any).osFiles = state;
    window.wordDesktop = {
      platform: 'test', setTheme() {}, onMenu: () => () => {},
      readRecovery: async () => {
        if (delayStartup) await new Promise<void>(resolve => { state.releaseStartup = resolve; });
        return [];
      },
      saveRecovery: async (serialized: string) => {
        if (state.fail) throw new Error('disk full');
        state.writes.push(serialized); return true;
      },
      onOpenRequested: (notify: () => void) => { state.notify = notify; return () => { state.notify = null; }; },
      takeRequestedFile: async () => {
        const file = state.queue.shift();
        if (file && state.holdRead) await new Promise<void>(resolve => { state.releaseRead = resolve; });
        return file ? { name: file.name, data: new Uint8Array(file.data) } : null;
      },
    } as any;
  }, { startup, delayStartup });
  await page.goto('/');
  if (!delayStartup) await expect(editor(page)).toBeVisible();
}
async function request(page: Page, ...files: RequestedFile[]) {
  await page.evaluate(files => { const state = (window as any).osFiles; state.queue.push(...files); state.notify?.(); }, files);
}

test('startup DOCX is opened after recovery initialization even when its notification predates the renderer', async ({ page }) => {
  const docx = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('더블클릭한 Word 문서 본문')] }] }));
  await desktop(page, [{ name: '한글 공백.DOCX', data: [...docx] }], true);
  await expect(editor(page)).toHaveCount(0);
  await page.evaluate(() => (window as any).osFiles.releaseStartup());
  await expect(editor(page)).toHaveText('더블클릭한 Word 문서 본문');
  await expect(page.getByRole('textbox', { name: '문서 제목' })).toHaveValue('한글 공백');
  await expect(page.getByText('이 PC에 저장됨', { exact: true })).toBeVisible();
});

test('several OS requests open in order and each previous document is preserved in recovery', async ({ page }) => {
  await desktop(page, [textFile('첫 문서.txt', 'FIRST'), textFile('두 번째.txt', 'SECOND'), textFile('마지막.txt', 'LAST')]);
  await expect(editor(page)).toHaveText('LAST');
  const writes = await page.evaluate(() => (window as any).osFiles.writes.map((raw: string) => JSON.parse(raw)));
  expect(writes.some((doc: any) => doc.title === '첫 문서' && JSON.stringify(doc.content).includes('FIRST'))).toBe(true);
  expect(writes.some((doc: any) => doc.title === '두 번째' && JSON.stringify(doc.content).includes('SECOND'))).toBe(true);
  await expect(page.getByRole('button', { name: '실행 취소 (Ctrl+Z)', exact: true })).toBeDisabled();
});

test('requests received while a dialog is open wait without consuming the requested file', async ({ page }) => {
  await desktop(page);
  await page.keyboard.press('ControlOrMeta+n');
  await expect(page.getByRole('dialog')).toBeVisible();
  await request(page, textFile('대기 문서.txt', 'OPEN AFTER CLOSE'));
  expect(await page.evaluate(() => (window as any).osFiles.queue.length)).toBe(1);
  await page.keyboard.press('Escape');
  await expect(editor(page)).toHaveText('OPEN AFTER CLOSE');
});

test('a modal opened during disk I/O delays replacement and late edits are saved before switching', async ({ page }) => {
  await desktop(page);
  await editor(page).fill('BEFORE READ');
  await page.evaluate(() => { (window as any).osFiles.holdRead = true; });
  await request(page, textFile('요청 문서.txt', 'AFTER READ'));
  await expect.poll(() => page.evaluate(() => Boolean((window as any).osFiles.releaseRead))).toBe(true);
  await editor(page).fill('EDITED WHILE READING');
  await page.keyboard.press('ControlOrMeta+n');
  await expect(page.getByRole('dialog')).toBeVisible();
  await page.evaluate(() => { const state = (window as any).osFiles; state.holdRead = false; state.releaseRead(); });
  await expect(editor(page)).toHaveText('EDITED WHILE READING');
  await page.keyboard.press('Escape');
  await expect(editor(page)).toHaveText('AFTER READ');
  const writes = await page.evaluate(() => (window as any).osFiles.writes);
  expect(writes.some((raw: string) => raw.includes('EDITED WHILE READING'))).toBe(true);
});

test('failed recovery during an external open protects edits and cancellation keeps the current document', async ({ page }) => {
  await desktop(page);
  await editor(page).fill('IMPORTANT UNSAVED EDITS');
  await page.evaluate(() => { (window as any).osFiles.fail = true; });
  await request(page, textFile('외부 문서.txt', 'REQUESTED CONTENT'));
  const failure = page.getByRole('dialog', { name: '현재 문서를 저장하지 못했습니다' });
  await expect(failure).toBeVisible(); await expect(editor(page)).toHaveText('IMPORTANT UNSAVED EDITS');
  await failure.getByRole('button', { name: '취소', exact: true }).click();
  await expect(failure).toHaveCount(0); await expect(editor(page)).toHaveText('IMPORTANT UNSAVED EDITS');
  await page.evaluate(() => { (window as any).osFiles.fail = false; });
  await request(page, textFile('외부 문서.txt', 'REQUESTED CONTENT'));
  await expect(editor(page)).toHaveText('REQUESTED CONTENT');
  const writes = await page.evaluate(() => (window as any).osFiles.writes);
  expect(writes.some((raw: string) => raw.includes('IMPORTANT UNSAVED EDITS'))).toBe(true);
});

test('a corrupt document reports failure without blocking later OS requests', async ({ page }) => {
  await desktop(page);
  await request(page, textFile('broken.docx', 'not a zip'), textFile('정상 문서.txt', 'VALID NEXT FILE'));
  await expect(editor(page)).toHaveText('VALID NEXT FILE');
  await expect(page.locator('.busy-indicator')).toHaveCount(0);
});

test('external HWPX opens in its original viewer and keeps subsequent requests until the viewer closes', async ({ page }) => {
  await desktop(page);
  await editor(page).fill('EXISTING WORD DOCUMENT');
  const hwpx = [...await fs.readFile('tests/fixtures/hwp/reader.hwpx')];
  await request(page, { name: '한글 원본.hwpx', data: hwpx }, textFile('다음 문서.txt', 'AFTER HWPX'));
  await expect(page.getByRole('dialog', { name: 'HWP 읽기', exact: true })).toBeVisible();
  await expect(page.locator('.hwp-sheet svg')).not.toHaveCount(0);
  await expect(editor(page)).toHaveText('EXISTING WORD DOCUMENT');
  expect(await page.evaluate(() => (window as any).osFiles.queue.length)).toBe(1);
  await page.getByRole('button', { name: 'HWP 읽기 닫기', exact: true }).click();
  await expect(editor(page)).toHaveText('AFTER HWPX');
});


test('the update modal preserves the current document and queues OS file requests until closed', async ({ page }) => {
  await desktop(page);
  await editor(page).fill('KEEP WHILE UPDATING');
  await page.getByRole('button', { name: '파일', exact: true }).click();
  await page.locator('.backstage-nav').getByRole('button', { name: '수동 업데이트·설정', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Word Studio 업데이트', exact: true });
  await expect(dialog).toBeVisible();
  await request(page, textFile('대기한 파일.txt', 'OPEN AFTER UPDATER CLOSE'));
  expect(await page.evaluate(() => (window as any).osFiles.queue.length)).toBe(1);
  await expect(editor(page)).toHaveText('KEEP WHILE UPDATING');
  await dialog.getByRole('button', { name: '업데이트 창 닫기', exact: true }).click();
  await expect(editor(page)).toHaveText('OPEN AFTER UPDATER CLOSE');
  const writes = await page.evaluate(() => (window as any).osFiles.writes);
  expect(writes.some((raw: string) => raw.includes('KEEP WHILE UPDATING'))).toBe(true);
});
