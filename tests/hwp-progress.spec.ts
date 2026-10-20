import { test, expect } from '@playwright/test';
import path from 'node:path';

test('Worker progress keeps the serial request pending, preserves its deadline and ignores late updates after close', async ({ page }) => {
  await page.goto('/'); await page.locator('.tiptap').waitFor();
  await page.clock.install();
  await page.evaluate(async () => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const workers: any[] = [];
    class ControlledWorker {
      onmessage: any; onerror: any; messages: any[] = [];
      constructor() { workers.push(this); }
      postMessage(message: any) { this.messages.push(message); }
      terminate() {}
    }
    const original = window.Worker; window.Worker = ControlledWorker as any;
    try {
      const reader = new HwpReaderSession(), events: string[] = [];
      let settled = false;
      const first = reader.editable(value => events.push(value.label)).then(value => { settled = true; return value; });
      const second = reader.pageText(0).catch(error => error.message);
      const worker = workers[0], id = worker.messages[0].id;
      worker.onmessage({ data: { id, kind: 'progress', progress: { label: '본문', completed: 2, total: 3 } } });
      await Promise.resolve();
      if (settled || worker.messages.length !== 1) throw Error('Progress completed a request or dispatched queued work');
      worker.onmessage({ data: { id, result: { htmlBytes: new TextEncoder().encode('<p>real result</p>').buffer, warnings: [], page: {} } } });
      const value = await first;
      if (value.html !== '<p>real result</p>' || worker.messages.length !== 2) throw Error('The real result did not advance the queue');
      reader.close(); await second;
      const before = events.length;
      worker.onmessage({ data: { id, kind: 'progress', progress: { label: 'late update' } } });
      if (events.length !== before) throw Error('Closed request received progress');
      const timeoutReader = new HwpReaderSession();
      const waiting = timeoutReader.editable(() => {}).catch(error => error.message);
      const queued = timeoutReader.pageText(1).catch(error => error.message);
      (window as any).progressDeadline = { worker: workers[1], waiting, queued };
    } finally { window.Worker = original; }
  });
  await page.clock.fastForward(100000);
  await page.evaluate(() => {
    const { worker } = (window as any).progressDeadline;
    worker.onmessage({ data: { id: worker.messages[0].id, kind: 'progress', progress: { label: 'still working' } } });
  });
  await page.clock.fastForward(80001);
  const result = await page.evaluate(async () => {
    const { worker, waiting, queued } = (window as any).progressDeadline;
    return { waiting: await waiting, queued: await queued, requests: worker.messages.length };
  });
  expect(result.waiting).toContain('처리 시간이 초과');
  expect(result.queued).toBe(result.waiting);
  expect(result.requests).toBe(1);
});

test('real HWP conversion displays measured stage progress, an hourglass and elapsed time, then opens the edited document', async ({ page }, testInfo) => {
  // Slow delivery, not conversion or document data, so every real Worker stage
  // remains visible long enough to verify its UI on fast machines.
  await page.addInitScript(() => {
    const NativeWorker = window.Worker;
    window.Worker = class extends NativeWorker {
      actions = new Map<number, string>();
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        const forwarded = new WeakSet<Event>(); let next = 0;
        this.addEventListener('message', event => {
          if (forwarded.has(event) || this.actions.get(event.data.id) !== 'editable') return;
          event.stopImmediatePropagation();
          const delivered = new MessageEvent('message', { data: event.data }); forwarded.add(delivered);
          next = Math.max(next, Date.now()) + 350;
          setTimeout(() => this.dispatchEvent(delivered), next - Date.now());
        });
      }
      postMessage(message: any, transfer: Transferable[] = []) { this.actions.set(message.id, message.action); super.postMessage(message, transfer); }
    };
  });
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.goto('/'); await page.locator('.tiptap').waitFor();
  await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/reader.hwp'));
  const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
  await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
  const progress = reader.getByRole('region', { name: 'HWP 처리 진행', exact: true });
  await expect(progress.getByTestId('hwp-hourglass')).toBeVisible();
  await expect(progress.getByLabel('경과 시간')).toContainText(/경과 [1-9]\d*초/);
  const bar = progress.getByRole('progressbar');
  await expect(bar).toBeVisible();
  const counts = await bar.evaluate((element: HTMLProgressElement) => ({ total: element.max, completed: element.value }));
  expect(counts.total).toBeGreaterThan(0); expect(counts.completed).toBeLessThanOrEqual(counts.total);
  await expect(progress).toContainText('(현재 단계)');
  await page.screenshot({ path: testInfo.outputPath('measured-conversion-progress.png') });
  const review = reader.getByRole('region', { name: '한글 편집 변환 안내' });
  await expect(review).toBeVisible({ timeout: 15000 });
  await expect(progress).not.toBeVisible();
  await review.getByRole('button', { name: '편집본 만들기', exact: true }).click();
  await expect(reader).not.toBeVisible();
  await expect(page.locator('.tiptap')).toContainText('테이블 속 한글');
  await expect(page.getByRole('region', { name: 'HWP 처리 진행', exact: true })).toHaveCount(0);
  expect(errors).toEqual([]);
});

test('closing during indeterminate work clears the hourglass and clock and keeps the current Word document', async ({ page }, testInfo) => {
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  await page.route('**/rhwp_bg.wasm*', async route => { await gate; await route.continue().catch(() => {}); });
  try {
    await page.goto('/'); await page.locator('.tiptap').fill('KEEP WHEN PROGRESS IS CANCELLED');
    await page.locator('input[type=file]').first().setInputFiles(path.resolve('tests/fixtures/hwp/reader.hwpx'));
    const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
    const progress = reader.getByRole('region', { name: 'HWP 처리 진행', exact: true });
    await expect(progress.getByTestId('hwp-hourglass')).toBeVisible();
    await expect(progress.getByRole('progressbar')).toHaveCount(0);
    await expect(progress.getByLabel('경과 시간')).toContainText(/경과 [1-9]\d*초/);
    await page.screenshot({ path: testInfo.outputPath('indeterminate-hourglass.png') });
    await reader.getByRole('button', { name: 'HWP 읽기 닫기', exact: true }).click();
    release();
    await expect(page.getByRole('region', { name: 'HWP 처리 진행', exact: true })).toHaveCount(0);
    await expect(page.locator('.tiptap')).toHaveText('KEEP WHEN PROGRESS IS CANCELLED');
  } finally { release(); }
});
