'use strict';

const { _electron: electron, expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const { createHash } = require('node:crypto');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');

async function saveHangulFile(app, page, item, directory, Core, phase) {
  const requested = (process.env.WINDOWS_NATIVE_HWP_FORMATS || 'hwp,hwpx').split(',');
  const formats = ['hwp', 'hwpx'].filter(format => requested.includes(format));
  if (!formats.length) throw new Error('No native Hangul output requested');
  if (process.env.WINDOWS_NATIVE_HWP_TRACE === '1') {
    await page.exposeFunction('qaHangulProgress', values => console.log('WINDOWS_NATIVE_HANGUL_PHASE ' + JSON.stringify(values)));
    await page.evaluate(() => {
      const groups = [
        ['prepare', /저장할 편집본|저장할 본문/],
        ['fonts', /문서의 글꼴/],
        ['images', /사용할 그림/],
        ['layout', /쪽 배치|레이아웃/],
        ['page-conversion', /쪽 나누기를 준비|문서의 쪽/],
        ['worker', /본문·표·그림|글자 서식|쪽 나누기를 반영/],
        ['format', /파일을 만드는/],
        ['reopen', /다시 읽어 확인/],
      ];
      let previous = '';
      const observer = new MutationObserver(() => {
        const section = document.querySelector('.hwp-progress-panel strong');
        const toast = document.querySelector('.toast span')?.textContent ?? '';
        const review = !!document.querySelector('.hwp-conversion-dialog[open]');
        const failure = /저장 실패/.test(toast) ? /시간이 초과/.test(toast) ? 'timeout' : 'other' : '';
        const step = groups.find(([, pattern]) => pattern.test(section?.textContent || ''))?.[0] ?? (section ? 'progress' : 'idle');
        const key = [step, review, failure].join('|');
        if (key === previous) return;
        previous = key;
        void window.qaHangulProgress({ step, review, failure, time: Math.round(performance.now()) });
      });
      observer.observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true,
        attributeFilter: ['open', 'class'] });
    });
  }
  for (const format of formats) {
    const output = path.join(directory, item.id + '-actual.' + format);
    phase('native-' + format + '-command');
    const started = Date.now();
    await app.evaluate(({ dialog, Menu }, { output, format }) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: output });
      const file = Menu.getApplicationMenu().items.find(item => item.label === 'File');
      const command = file?.submenu?.items.find(item => item.label === 'Export Hangul Document (' + format.toUpperCase() + ')…');
      if (!command) throw new Error('Native Hangul command unavailable');
      setTimeout(() => command.click(), 0);
    }, { output, format });
    const review = page.getByRole('dialog', { name: '한글 문서로 내보내기' });
    phase('native-' + format + '-review');
    await review.waitFor({ state: 'visible', timeout: 240000 });
    const comparison = await review.getByText(/^쪽 수 비교:/).innerText();
    const pages = comparison.match(/Word 인쇄 (\d+)쪽.*?원본 보기 (\d+)쪽/);
    phase('native-' + format + '-save');
    await review.getByRole('button', {name: format.toUpperCase() + ' 파일 저장', exact: true}).click({ timeout: 30000 });
    await review.waitFor({ state: 'hidden', timeout: 90000 });
    const bytes = await fs.readFile(output), native = new Core(bytes);
    try {
      const readback = native.pageCount();
      if (!pages || readback < 1 || readback !== Number(pages[2])) throw new Error('Native Hangul readback disagreed');
      console.log('WINDOWS_NATIVE_HANGUL ' + JSON.stringify({ id:item.id, format, bytes:bytes.length,
        printPages:Number(pages[1]), savedPages:readback, seconds:(Date.now()-started)/1000 }));
      if (item.id === '04-language-plan' && process.env.WINDOWS_NATIVE_LINE_WITNESS === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          const witness = require('./hancom-line-witness.cjs');
          const result = witness(original, native);
          console.log('WINDOWS_HANCOM_LINE_WITNESS ' + JSON.stringify({format, ...result}));
          if (result.lines !== 1 || result.nextPage !== 8) throw Error('Native saved line disagreed with the pinned witness');
        } finally { original.free(); }
      }
      const comparison = process.env.WINDOWS_NATIVE_HWP_COMPARISON;
      if (comparison) {
        const hash = createHash('sha256');
        for (let i = 0; i < readback; i++) for (const evidence of [native.getPageInfo(i),
          native.getPageTextLayout(i), native.getPageControlLayout(i)]) hash.update(evidence).update('\n');
        await fs.writeFile(comparison, hash.digest('hex'));
      }
    } finally { native.free(); }
  }
}

async function main() {
  const inventory = JSON.parse(await fs.readFile(process.env.WINDOWS_PUBLIC_QA_INVENTORY, 'utf8'));
  const directory = await fs.mkdtemp(path.join(process.env.RUNNER_TEMP || os.tmpdir(), 'studio-native-preview-'));
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  delete env.WORD_STUDIO_DEV;
  let Core;
  if (process.env.WINDOWS_NATIVE_HWP_SAVES === '1') {
    const directory = path.dirname(require.resolve('@rhwp/core'));
    const { default: init, HwpDocument } = await import(pathToFileURL(path.join(directory, 'rhwp.js')).href);
    await init({ module_or_path: await fs.readFile(path.join(directory, 'rhwp_bg.wasm')) });
    Core = HwpDocument;
  }
  const requestedScale = process.env.WINDOWS_FORCE_DEVICE_SCALE;
  const dpiArgument = requestedScale && ['1','1.25','1.5','2'].includes(requestedScale)
    ? ['--force-device-scale-factor=' + requestedScale] : [];
  if (requestedScale && !dpiArgument.length) throw Error('Invalid requested scale');
  let failures = 0;
  let measured = 0;

  for (const item of inventory.filter(row => row.verified)) {
    const pdf = path.join(directory, item.id + '.pdf');
    let stage = 'launch';
    let app;
    let close;
    try {
      app = await electron.launch({
        executablePath: path.resolve(process.env.WORD_STUDIO_EXECUTABLE),
        args: ['--user-data-dir=' + path.join(directory, item.id), ...dpiArgument],
        env,
      });
      close = await require('../native-smoke-lifetime.cjs')(app, 'Public original ' + item.id, 1050000);
      const page = await app.firstWindow();
      await page.setViewportSize({ width: 1440, height: 1000 });
      if (requestedScale) {
        const available = await page.evaluate(() => devicePixelRatio);
        console.log('WINDOWS_DENSITY ' + JSON.stringify({ requested:Number(requestedScale), actual:available }));
        if (available !== Number(requestedScale)) throw Error('The renderer did not use the requested device density');
      }
      await page.evaluate(() => {
        const tasks = [];
        window.qaLongTasks = tasks;
        new PerformanceObserver(list => {
          for (const task of list.getEntries()) tasks.push({ at: task.startTime, duration: task.duration });
        }).observe({ type: 'longtask' });
      });
      await page.locator('.tiptap').first().waitFor({ state: 'visible' });
      await app.evaluate(({ dialog, BrowserWindow }, files) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [files.source] });
        const markers = {};
        globalThis.qaPdfStages = markers;
        dialog.showSaveDialog = async () => {
          markers.dialog = Date.now();
          return { canceled: false, filePath: files.pdf };
        };
        const contents = BrowserWindow.getAllWindows()[0].webContents;
        const makePdf = contents.printToPDF.bind(contents);
        contents.printToPDF = async options => {
          markers.chromiumStarted = Date.now();
          const result = await makePdf(options);
          markers.chromiumCompleted = Date.now();
          return result;
        };
      }, { source: item.path, pdf });
      stage = 'open';
      await page.keyboard.press('ControlOrMeta+o');
      let readerPages = null, readyStarted = Date.now(), rendererStarted = 0;
      if (item.ext !== 'docx') {
        const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
        await reader.getByTestId('hwp-page').waitFor({ state: 'visible', timeout: 75000 });
        readerPages = Number((await reader.getByTestId('hwp-page-count').textContent()).replace(/\D/g, ''));
        if (!Number.isSafeInteger(readerPages) || readerPages < 1) throw new Error('Reader page count unavailable');
        console.log('WINDOWS_NATIVE_READER ' + item.id + ' ' + readerPages);
        stage = 'convert-analysis';
        const conversionStarted = Date.now();
        await reader.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
        const panel = reader.getByRole('region', { name: '한글 편집 변환 안내' });
        await panel.waitFor({ state: 'visible', timeout: 240000 });
        const comparison = await panel.locator('strong').first().textContent();
        const previewEstimate = comparison?.match(/Word 인쇄 (\d+)쪽/);
        console.log('WINDOWS_NATIVE_ESTIMATE ' + item.id + ' ' + (previewEstimate?.[1] ?? 'unavailable'));
        stage = 'convert-editor-command';
        readyStarted = Date.now();
        rendererStarted = await page.evaluate(() => performance.now());
        await panel.getByRole('button', { name: '편집본 만들기', exact: true }).click({ timeout: 90000 });
        stage = 'convert-editor-view';
        await reader.waitFor({ state: 'hidden', timeout: 360000 });
        console.log('WINDOWS_NATIVE_CONVERT ' + item.id + ' ' + (Date.now() - conversionStarted));
      }
      stage = 'editor-title';
      await expect(page.locator('.topbar').getByRole('textbox', { name: '문서 제목', exact: true }))
        .toHaveValue(item.id, { timeout: 60000 });
      stage = 'editor-ready';
      await page.locator('.busy-indicator').waitFor({ state: 'hidden', timeout: 90000 });
      await page.evaluate(() => document.fonts.ready);
      const stack = page.locator('.document-stack');
      await expect(stack).toHaveAttribute('data-pages', /^[1-9]\d*$/, { timeout: 60000 });
      const editorPages = Number(await stack.getAttribute('data-pages'));
      const editorMilliseconds = Date.now() - readyStarted;
      const editorLongTasks = await page.evaluate(start => window.qaLongTasks.filter(task => task.at >= start)
        .map(task => Math.round(task.duration)).sort((a,b) => b-a).slice(0,5), rendererStarted);
      console.log('WINDOWS_HWP_EDITOR_READY ' + JSON.stringify({ id: item.id, editorPages, editorMilliseconds, editorLongTasks }));
      if (Core) {
        const finishHangulProfile = process.env.WINDOWS_NATIVE_HWP_PROFILE === '1'
          ? await require('./print-profile.cjs')(page, 'hangul') : undefined;
        try {
          await saveHangulFile(app, page, item, directory, Core, name => { stage = name; });
          measured++;
        } finally {
          // Keep the timing error if profiling could not read a stalled renderer.
          await finishHangulProfile?.().catch(() => console.log('WINDOWS_NATIVE_PROFILE unavailable'));
        }
        continue;
      }
      const finishProfile = process.env.WINDOWS_NATIVE_PRINT_PROFILE === '1'
        ? await require('./print-profile.cjs')(page) : undefined;
      stage = 'native-export-command';
      const rendererStart = await page.evaluate(() => performance.now());
      const started = Date.now();
      await app.evaluate(({ Menu }) => {
        const file = Menu.getApplicationMenu().items.find(item => item.label === 'File');
        const exportPdf = file?.submenu?.items.find(item => item.label === 'Export PDF…');
        if (!exportPdf) throw new Error('Native PDF menu command unavailable');
        setTimeout(() => {
          globalThis.qaPdfStages.menu = Date.now();
          exportPdf.click();
        }, 0);
      });
      console.log('WINDOWS_NATIVE_PDF_MENU scheduled');
      stage = 'pdf-file';
      let pages = 0;
      await expect.poll(async () => {
        const bytes = await fs.readFile(pdf).catch(() => null);
        if (!bytes || !bytes.subarray(-16).toString('latin1').includes('%%EOF')) return 0;
        pages = (bytes.toString('latin1').match(/\/Type\s*\/Page\b/g) || []).length;
        return pages;
      }, { timeout: 420000 }).toBeGreaterThan(0);
      const previewPages = await page.locator('.print-layout.preview .print-page').count();
      const markers = await app.evaluate(() => globalThis.qaPdfStages);
      console.log('WINDOWS_NATIVE_PDF_RESULT ' + JSON.stringify({ id: item.id, readerPages, editorPages,
        previewPages, pdfPages: pages, totalMilliseconds: Date.now() - started,
        timings: Object.fromEntries(Object.entries(markers).map(([key, value]) => [key, value - started])) }));
      const longest = await page.evaluate(start => Math.max(0, ...window.qaLongTasks
        .filter(task => task.at >= start).map(task => Math.round(task.duration))), rendererStart);
      console.log('WINDOWS_NATIVE_HOLD_LONGEST ' + longest);
      if (process.env.WINDOWS_NATIVE_PRINT_GEOMETRY) {
        stage = 'pdf-geometry';
        const geometry = await require('./print-geometry.cjs')(page);
        await fs.writeFile(process.env.WINDOWS_NATIVE_PRINT_GEOMETRY, JSON.stringify(geometry));
      }
      if (process.env.WINDOWS_NATIVE_PDF_COPY) {
        if (inventory.length !== 1) throw new Error('A fixed single PDF case is required');
        await fs.copyFile(pdf, process.env.WINDOWS_NATIVE_PDF_COPY);
      }
      if (process.env.WINDOWS_WORD_TABLE_WITNESS === '1') {
        const { sheets, fonts } = await require('./word-table-witness.cjs')(page);
        console.log('WINDOWS_WORD_TABLE_FACES ' + JSON.stringify(fonts));
        for (const sheet of sheets) console.log('WINDOWS_WORD_TABLE_SHEET ' + JSON.stringify(sheet));
      }
      await finishProfile?.();
      measured++;
    } catch (error) {
      // Playwright errors can include source content. Emit only fixed stage labels.
      const reason = error?.name === 'TimeoutError'
        ? (/intercepts pointer events/.test(error.message) ? 'pointer-interception' : 'timeout')
        : 'error';
      console.log('WINDOWS_NATIVE_FAILURE ' + item.id + ' ' + stage + ' ' + reason);
      if (app) {
        const markers = await app.evaluate(() => globalThis.qaPdfStages || {}).catch(() => ({}));
        console.log('WINDOWS_NATIVE_PDF_STAGES ' + JSON.stringify(Object.keys(markers)));
      }
      failures++;
    } finally {
      if (close) {
        const closeStarted = Date.now();
        try {
          await close();
          console.log('WINDOWS_NATIVE_CLOSE ' + JSON.stringify({ id:item.id, clean:true, milliseconds:Date.now()-closeStarted }));
        } catch {
          // A forced close of this disposable app is a separate failure from
          // a PDF already written after the native print-to-file dialog.
          console.log('WINDOWS_NATIVE_CLOSE ' + JSON.stringify({ id:item.id, clean:false, milliseconds:Date.now()-closeStarted }));
          if (process.env.WINDOWS_NATIVE_PDF_CLOSE_CONTROL !== '1') failures++;
        }
      }
      else if (app?.process().exitCode === null) app.process().kill();
    }
  }
  console.log('WINDOWS_NATIVE_TOTAL expected ' + inventory.length + ' measured ' + measured + ' failed ' + failures);
  if (failures || measured !== inventory.length) process.exitCode = 1;
}

main().catch(() => { console.error('Native original check failed outside a document stage'); process.exitCode = 1; });
