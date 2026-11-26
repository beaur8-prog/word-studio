'use strict';

const { _electron: electron, expect } = require('@playwright/test');
const fs = require('node:fs/promises');
const { createHash } = require('node:crypto');
const path = require('node:path');
const os = require('node:os');
const { pathToFileURL } = require('node:url');

async function saveTwoDocxAndReopen(app, page, item, directory, selectStage) {
  const title = page.locator('.topbar').getByRole('textbox', { name: '문서 제목', exact: true });
  await fs.mkdir(directory, { recursive: true });
  const readStoredAdvances = () => page.locator('.document-stack p[data-hwp-body-line^="14 "]')
    .evaluateAll(paragraphs => paragraphs.map(p => [p.dataset.hwpBodyLine, p.textContent]));
  const storedAdvances = process.env.WINDOWS_CHECK_HWP_MIXED_ADVANCE === '1' ? await readStoredAdvances() : null;
  if (storedAdvances && storedAdvances.length < 3) throw Error('Missing original mixed-height measurements');
  for (let generation = 1; generation <= 2; generation++) {
    selectStage('native-docx-save-' + generation);
    const name = 'textbox-save-' + item.id + '-native-' + generation;
    const output = path.join(directory, name + '.docx');
    const started = Date.now();
    const trace = process.env.WINDOWS_NATIVE_DOCX_SAVE_TRACE === '1';
    const rendererStarted = trace ? await page.evaluate(() => performance.now()) : 0;
    if (trace) await app.evaluate(({ BrowserWindow }) => {
      const view = BrowserWindow.getAllWindows()[0]?.webContents;
      globalThis.qaDocxSave = { dialogCalls:0, rendererExit:'none' };
      view?.once('render-process-gone', (_event, detail) => {
        globalThis.qaDocxSave.rendererExit = detail.reason;
      });
    });
    await app.evaluate(({ dialog, Menu }, destination) => {
      dialog.showSaveDialog = async (_parent, options) => {
        if (globalThis.qaDocxSave) {
          globalThis.qaDocxSave.dialogCalls++;
          globalThis.qaDocxSave.dialogTime = Date.now();
        }
        if (!options.filters?.some(filter => filter.extensions.includes('docx')))
          throw Error('Missing DOCX filter');
        return { canceled: false, filePath: destination };
      };
      const file = Menu.getApplicationMenu().items.find(item => item.label === 'File');
      const command = file?.submenu?.items.find(item => item.label.startsWith('Export Word Document '));
      if (!command) throw Error('Native DOCX export missing');
      setTimeout(() => command.click(), 0);
    }, output);
    const timeout = trace ? 300000 : 180000;
    const readStatus = async () => {
      const state = await app.evaluate(({ app }) => ({
        ...globalThis.qaDocxSave,
        renderer:app.getAppMetrics().filter(p=>p.type==='Tab').map(p=>({
          cpu:Number(p.cpu.percentCPUUsage.toFixed(1)), workingSetKB:p.memory.workingSetSize,
        })),
      }));
      const stats = await fs.stat(output).catch(() => null);
      console.log('WINDOWS_LARGE_DOCX_SAVE_TRACE ' + JSON.stringify({
        generation, duration:Date.now()-started, bytes:stats?.size ?? 0,
        ...state, dialogTime:state.dialogTime ? state.dialogTime-started : undefined,
      }));
    };
    let monitor;
    try {
      if (trace) {
        monitor = setInterval(() => { void readStatus().catch(() => {}); }, 15000);
        monitor.unref();
      }
      await expect.poll(async () => {
        const bytes = await fs.readFile(output).catch(() => null);
        return bytes?.length > 64 && bytes.subarray(0, 2).toString('latin1') === 'PK';
      }, {timeout}).toBe(true);
    } finally {
      clearInterval(monitor);
      if (trace) await readStatus().catch(() => {});
    }
    if (trace) {
      const peak = await page.evaluate(start => ({
        longest: Math.max(0,...window.qaLongTasks.filter(task => task.at >= start).map(task => Math.round(task.duration))),
        tasks: window.qaLongTasks.filter(task => task.at >= start).length,
      }), rendererStarted);
      console.log('WINDOWS_LARGE_DOCX_RENDERER ' + JSON.stringify({generation,...peak}));
    }
    // Bytes appear before the file operation and renderer save guard end.
    // Windows can ignore a new Open command during that interval.
    await page.locator('.busy-indicator').waitFor({state: 'hidden', timeout: 180000});
    selectStage('native-docx-reopen-' + generation);
    const openAt = Date.now();
    await app.evaluate(({ dialog, Menu }, source) => {
      globalThis.qaDocxOpenDialogCount = 0;
      dialog.showOpenDialog = async () => {
        globalThis.qaDocxOpenDialogCount++;
        return { canceled: false, filePaths: [source] };
      };
      const file = Menu.getApplicationMenu().items.find(item => item.label === 'File');
      const command = file?.submenu?.items.find(item => item.label === 'Open…');
      if (!command) throw Error('Native Open command unavailable');
      setTimeout(() => command.click(), 0);
    }, output);
    if (process.env.WINDOWS_NATIVE_DOCX_DIAGNOSE === '1') {
      const classify = () => page.evaluate(() => {
        const toast = document.querySelector('.toast')?.textContent ?? '';
        const category = toast.startsWith('파일을 열지 못했습니다') ?
          (/본문 속성/.test(toast) ? 'body-attribute' : /행 높이|셀 구조|표 /.test(toast) ? 'table-structure'
          : /이미지|그림/.test(toast) ? 'picture' : /형식|DOCX/.test(toast) ? 'docx-import'
          : /용량|너무|한도|MiB/.test(toast) ? 'size-limit' : /Invalid content|Unknown node/.test(toast) ? 'schema'
          : 'other-import') : toast.startsWith('자동 저장 실패') ? 'recovery' : null;
        return { category, busy: !!document.querySelector('.busy-indicator'),
          recovery: [...document.querySelectorAll('dialog,[role="dialog"]')].some(el => /복구|자동 저장/.test(el.textContent ?? '')),
          pictures: document.querySelectorAll('.tiptap img').length };
      });
      const reopened = await Promise.race([
        expect(title).toHaveValue(name, {timeout:45000}).then(()=>true).catch(()=>false),
        expect.poll(async()=>{const stage=await classify();return !!stage.category || stage.recovery;},
          {timeout:45000}).toBe(true).then(()=>false).catch(()=>false),
      ]);
      const state = await classify();
      const dialogCalls = await app.evaluate(() => globalThis.qaDocxOpenDialogCount);
      console.log('WINDOWS_NATIVE_DOCX_DIAGNOSIS ' + JSON.stringify({ id:item.id, generation, dialogCalls, reopened,
        bytes:(await fs.stat(output)).size, milliseconds:Date.now()-openAt, ...state }));
      if (!reopened) throw Error('Native DOCX reopen did not reach the editor');
    } else {
      await expect(title).toHaveValue(name, {timeout: 180000});
    }
    await page.locator('.busy-indicator').waitFor({state: 'hidden', timeout: 180000});
    await page.evaluate(() => document.fonts.ready);
    if (storedAdvances) {
      const roundtrip = await readStoredAdvances();
      const equal = JSON.stringify(roundtrip) === JSON.stringify(storedAdvances);
      console.log('WINDOWS_MIXED_ADVANCE_ROUNDTRIP ' + JSON.stringify({ generation, expected:storedAdvances.length, actual:roundtrip.length, equal }));
      if (!equal) throw Error('Stored mixed-font advances changed');
    }
    console.log('WINDOWS_NATIVE_DOCX_SAVED_AND_REOPENED ' + generation);
  }
}

async function saveHangulFile(app, page, item, directory, Core, phase) {
  const expectedImages = process.env.WINDOWS_NATIVE_PICTURE_PAINT === '1'
    ? await page.locator('.tiptap img[data-inline-image]').evaluateAll(images => images.map(image => image.src.split(',')[1]))
    : null;
  if (expectedImages && expectedImages.length !== 2) throw Error('Expected both source group previews in the editor');
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
    const timeout = Number(process.env.WINDOWS_NATIVE_HWP_REVIEW_MS || 240000);
    const exportError = page.locator('.toast span').filter({ hasText: /^저장 실패:/ });
    const reviewResult = await Promise.race([
      review.waitFor({ state: 'visible', timeout }).then(() => 'review'),
      exportError.waitFor({ state: 'visible', timeout }).then(() => 'error'),
    ]);
    if (reviewResult === 'error') {
      const detail = await exportError.textContent();
      const reason = /시간이 초과/.test(detail || '') ? 'native-worker-timeout' : 'engine-or-bridge-error';
      console.log('WINDOWS_NATIVE_HANGUL_REVIEW_ERROR ' + JSON.stringify({id: item.id, format, reason}));
      throw Error(reason);
    }
    const comparison = await review.getByText(/^쪽 수 비교:/).innerText();
    const pages = comparison.match(/Word 인쇄 (\d+)쪽.*?원본 보기 (\d+)쪽/);
    phase('native-' + format + '-save');
    await review.getByRole('button', {name: format.toUpperCase() + ' 파일 저장', exact: true}).click({ timeout: 30000 });
    await review.waitFor({ state: 'hidden', timeout: 90000 });
    const bytes = await fs.readFile(output), native = new Core(bytes);
    if (process.env.WINDOWS_NATIVE_HWP_COPY)
      await fs.writeFile(process.env.WINDOWS_NATIVE_HWP_COPY + '.' + format, bytes);
    try {
      const readback = native.pageCount();
      if (!pages || readback < 1 || readback !== Number(pages[2])) throw new Error('Native Hangul readback disagreed');
      if (process.env.WINDOWS_CHECK_NATIVE_NBSPACE === '1') {
        const source = new Core(await fs.readFile(item.path));
        try {
          const literal = source.getTextRange(0, 32, 0, 0xffffffff);
          const compact = text => text.replace(/\s/gu, '');
          const matching = document => {
            const found = [];
            const count = document.getParagraphCount(0);
            for (let index = 0; index < count; index++) {
              const current = document.getTextRange(0, index, 0, 0xffffffff);
              if (compact(current) === compact(literal)) found.push(current);
            }
            return found;
          };
          const spaces = text => [...text].filter(value => value === '\u00a0').length;
          const originals = matching(source), saved = matching(native);
          const exact = originals.length === saved.length && originals.every((value, index) => value === saved[index]);
          console.log('WINDOWS_NATIVE_NBSPACE ' + JSON.stringify({ format, exact, sourceNoBreak: spaces(literal),
            sourceMatches: originals.length, savedMatches: saved.length,
            savedNoBreak: saved.map(spaces), savedPages: native.pageCount() }));
          if (spaces(literal) !== 5 || !exact) throw Error('Original nonbreaking spaces changed');
        } finally { source.free(); }
      }
      const expected = (process.env.WINDOWS_EXPECT_NATIVE_PAGES || undefined)?.split(',').map(Number);
      console.log('WINDOWS_NATIVE_HANGUL ' + JSON.stringify({ id:item.id, format, bytes:bytes.length,
        printPages:Number(pages[1]), savedPages:readback, seconds:(Date.now()-started)/1000 }));
      if (process.env.WINDOWS_NATIVE_BLANK_SHEET_AUDIT === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          const metrics = require('./hancom-blank-sheet-boundaries.cjs')(Core, original, native, bytes);
          console.log('WINDOWS_NATIVE_BLANK_SHEET_AUDIT ' + JSON.stringify({ format, ...metrics }));
        } finally { original.free(); }
      }
      if (process.env.WINDOWS_HWP_FIRST_PAGE_SOURCE_TRIAL === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          const metrics = require('./hancom-first-page-source-lines.cjs')(Core,original,native);
          console.log('WINDOWS_HWP_FIRST_PAGE_SOURCE_TRIAL ' + JSON.stringify(metrics));
        } catch (error) {
          const reported = new Set(['Need one complete paper-width table','Cannot uniquely map table structure',
            'Missing table placement','Unexpected title carrier geometry','Missing an unchanged authored tab row',
            'Changed paragraph payload']);
          console.log('WINDOWS_HWP_FIRST_PAGE_TRIAL_STAGE ' + (reported.has(error.message) ? error.message : 'engine or measurement error'));
          throw error;
        } finally { original.free(); }
      }
      if (process.env.WINDOWS_MEASURE_PACKAGED_PAGE_BOUNDARIES === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          const agreement = require('./hancom-page-agreement.cjs')(original, native);
          if (process.env.WINDOWS_PACKAGED_TABLE_FIT === '1') {
            const result = require('./hancom-packaged-table-fit.cjs')(Core, original, native);
            console.log('WINDOWS_PACKAGED_TABLE_FIT ' + JSON.stringify(result));
          }
          const pages = Array.from({length:readback}, (_, page) => {
            const runs = JSON.parse(native.getPageTextLayout(page)).runs.filter(r => !r.headerFooter && r.paraIdx!==0xffffffff && r.text.trim());
            const controls = JSON.parse(native.getPageControlLayout(page)).controls.filter(c => !c.headerFooter);
            const owner = runs.find(r=>!r.cellPath?.length);
            return {page:page+1, body:runs.filter(r=>!r.cellPath?.length).reduce((sum,r)=>sum+r.text.replace(/\s/gu,'').length,0),
              cells:runs.filter(r=>r.cellPath?.length).reduce((sum,r)=>sum+r.text.replace(/\s/gu,'').length,0),
              first:owner?[owner.secIdx,owner.paraIdx,owner.charStart,owner.y]:null,
              objects:controls.map(c=>[c.type,c.secIdx,c.paraIdx,c.x,c.y,c.w,c.h])};
          });
          console.log('WINDOWS_PACKAGED_PAGE_AGREEMENT ' + JSON.stringify({format,product:process.env.WINDOWS_PRODUCT,
            fixed: {body:agreement.body,cell:agreement.cell},pages}));
        } finally { original.free(); }
      }
      if (process.env.WINDOWS_MEASURE_BINARY_TAIL === '1') {
        const tail = Array.from({length:readback}, (_, i) => {
          const runs = JSON.parse(native.getPageTextLayout(i)).runs;
          const owner = runs.filter(r => r.secIdx === 0 && r.paraIdx === 38 && !r.cellPath?.length && !r.headerFooter && r.text.trim());
          return [i+1, ...owner.map(r=>[r.charStart,r.y,r.h,r.text.replace(/\s/gu,'').length])];
        }).filter(line => line.length > 1);
        console.log('WINDOWS_PACKAGED_BINARY_TAIL ' + JSON.stringify({format,pages:readback,tail}));
      }
      if (expected && readback !== expected[format === 'hwp' ? 0 : 1])
        throw new Error('Native Hangul final page count disagreed');
      if (item.id === '04-language-plan' && process.env.WINDOWS_NATIVE_LINE_WITNESS === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          const witness = require('./hancom-line-witness.cjs');
          const result = witness(original, native);
          console.log('WINDOWS_HANCOM_LINE_WITNESS ' + JSON.stringify({format, ...result}));
          if (result.lines !== 1 || result.nextPage !== 8) throw Error('Native saved line disagreed with the pinned witness');
        } finally { original.free(); }
      }
      if (item.id === '06-mixed-document' && process.env.WINDOWS_NATIVE_TERMINAL_WITNESS === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          const witness = require('./hancom-terminal-witness.cjs')(original, native);
          console.log('WINDOWS_NATIVE_TERMINAL_WITNESS ' + JSON.stringify({format, ...witness}));
          if (!witness.matches) throw Error('Native terminal positions disagreed with the original');
        } finally { original.free(); }
      }
      if (item.id === '06-mixed-document' && process.env.WINDOWS_NATIVE_PAGE_TAIL_READBACK === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          const { bodyTail } = require('./hancom-source-page-offsets.cjs')(original, native);
          const adjacent = require('./hancom-adjacent-tables-witness.cjs')(original, native);
          console.log('WINDOWS_NATIVE_MENU_PAGE_TAIL ' + JSON.stringify({ format, bodyTail,
            mainTablePage: adjacent.saved.page, mainTableGlyphs: adjacent.saved.glyphs, originalTablePage: adjacent.source.page }));
          if (!bodyTail || bodyTail.saved.length !== 2 || bodyTail.saved.some(line => line.page !== 7) ||
              bodyTail.saved.reduce((count, line) => count + line.glyphs, 0) !== 60 || !adjacent.matches)
            throw Error('The packaged Windows save disagreed with the native page-tail witness');
        } finally { original.free(); }
      }
      if (item.id === '02-work-rules' && process.env.WINDOWS_NATIVE_LARGE_HANGUL === '1') {
        const original = new Core(await fs.readFile(item.path));
        try {
          if (process.env.WINDOWS_NATIVE_OVERHANG_READBACK === '1') {
            const boundary = doc => {
              const { controls } = JSON.parse(doc.getPageControlLayout(8));
              const first = controls.filter(c=>c.type==='table'&&!c.cellPath?.length&&!c.headerFooter);
              if (first.length!==1) return {tableCount:first.length};
              const [table]=first;
              const runs=JSON.parse(doc.getPageTextLayout(8)).runs.filter(r=>r.parentParaIdx===table.paraIdx &&
                r.secIdx===table.secIdx && !r.headerFooter && r.cellPath?.[0]?.controlIndex===table.controlIdx);
              const count=entries=>entries.reduce((sum,run)=>sum+run.text.replace(/\s/gu,'').length,0);
              const row=table.cells.filter(c=>c.row===7);
              return { shapes:first.length, xy:[table.x,table.y], height:table.h, glyphs:count(runs),
                heads:row.map(cell=>count(runs.filter(run=>run.cellPath[0].cellIndex===cell.cellIdx))) };
            };
            const expected=boundary(original), actual=boundary(native);
            console.log('WINDOWS_LARGE_NATIVE_TABLE_BOUNDARY ' + JSON.stringify({format,expected,actual}));
            if (process.env.WINDOWS_NATIVE_OVERHANG_ORIGIN === '1' && (!actual.xy || !expected.xy || Math.abs(actual.xy[0] - expected.xy[0]) > .15))
              throw new Error('The first retained native table moved away from its authored origin');
            const agreement=require('./hancom-page-agreement.cjs')(original,native);
            console.log('WINDOWS_LARGE_NATIVE_TABLE_AGREEMENT ' + JSON.stringify({format,...agreement}));
            if (actual.glyphs!==expected.glyphs || String(actual.heads)!==String(expected.heads))
              throw Error('The measured first table fragment did not retain the two authored cell heads');
          }
          const snapshot = doc => {
            let glyphs = 0, tables = 0, images = 0, outside = 0;
            for (let index = 0; index < doc.pageCount(); index++) {
              const paper = JSON.parse(doc.getPageInfo(index));
              for (const run of JSON.parse(doc.getPageTextLayout(index)).runs) {
                if (!run.headerFooter && Number.isInteger(run.paraIdx) && run.paraIdx !== 0xffffffff)
                  glyphs += run.text.replace(/\s/gu, '').length;
              }
              for (const object of JSON.parse(doc.getPageControlLayout(index)).controls) {
                if (object.headerFooter) continue;
                if (object.type === 'table') tables++;
                if (object.type === 'image') images++;
                if (object.type === 'table' && (object.x < -.2 || object.x + object.w > paper.width + .2)) outside++;
              }
            }
            return { pages: doc.pageCount(), glyphs, tables, images, outside };
          };
          console.log('WINDOWS_LARGE_NATIVE_READBACK ' + JSON.stringify({ format, source: snapshot(original), saved: snapshot(native) }));
          if (process.env.WINDOWS_NATIVE_LARGE_TEXT === '1') {
            const { execFileSync } = require('node:child_process');
            const originalPackage = path.join(directory, 'logical-source.hwpx');
            const savedPackage = path.join(directory, 'logical-saved.hwpx');
            await fs.writeFile(originalPackage, original.exportHwpx());
            if (format === 'hwp') await fs.writeFile(savedPackage, native.exportHwpx());
            const result = execFileSync('python', [path.join(__dirname, 'hancom-package-text-audit.py'),
              originalPackage, format === 'hwp' ? savedPackage : output], { encoding: 'utf8' });
            console.log('WINDOWS_LARGE_NATIVE_LOGICAL ' + JSON.stringify({ format, ...JSON.parse(result) }));
          }
        } finally { original.free(); }
      }
      if (expectedImages) {
        const geometry = [];
        const savedBytes = [];
        for (let index = 0; index < readback; index++) {
          const paper = JSON.parse(native.getPageInfo(index));
          const pictures = JSON.parse(native.getPageControlLayout(index)).controls.filter(item => item.type === 'image' && !item.headerFooter);
          for (const picture of pictures) {
            const { secIdx, paraIdx, controlIdx, cellPath } = picture;
            if (cellPath?.length) throw Error('Unexpected nested output illustration');
            const data = native.getControlImageData(secIdx, paraIdx, '[]', controlIdx);
            savedBytes.push(Buffer.from(data).toString('base64'));
            geometry.push({page: index + 1, x: picture.x, y: picture.y, w: picture.w, h: picture.h,
              offset: picture.y - paper.marginTop - paper.marginHeader,
              inside: picture.x >= 0 && picture.y >= 0 && picture.x + picture.w <= paper.width && picture.y + picture.h <= paper.height });
          }
        }
        const exact = JSON.stringify(savedBytes) === JSON.stringify(expectedImages);
        console.log('WINDOWS_LARGE_NATIVE_PICTURE_PAINT ' + JSON.stringify({format, exact, geometry}));
        if (!exact || geometry.length !== 2 || geometry.some(item => !item.inside)) throw Error('Missing or modified picture in the native readback');
        const expectedOffset = Number(process.env.WINDOWS_NATIVE_PICTURE_OFFSET);
        if (Number.isFinite(expectedOffset) && geometry.some(item => Math.abs(item.offset - expectedOffset) > .12))
          throw Error('The inline illustration did not recover its authored outer margin');
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
      if (process.env.WINDOWS_TRACE_COLUMN_END === '1') page.on('console', message => {
        const line = message.text();
        if (line.startsWith('WINDOWS_HWP_COLUMN_STAGES ')) console.log(line);
      });
      if (process.env.WINDOWS_NATIVE_MATH_STAGE_TRACE === '1') page.on('console', message => {
        const line = message.text();
        if (line.startsWith('QA_MATH_NATIVE_STAGE ')) console.log(line);
      });
      if (requestedScale) {
        stage = 'display-scale';
        // Playwright's viewport emulation also overrides deviceScaleFactor.
        // Resize the actual Electron window to preserve its selected screen density.
        const screenScale = await app.evaluate(({ BrowserWindow, screen }) => {
          const window = BrowserWindow.getAllWindows()[0];
          window.setContentSize(1440, 1000);
          return screen.getDisplayMatching(window.getBounds()).scaleFactor;
        });
        await expect.poll(() => page.evaluate(() => devicePixelRatio)).toBe(Number(requestedScale));
        const available = await page.evaluate(() => devicePixelRatio);
        console.log('WINDOWS_DENSITY ' + JSON.stringify({ requested:Number(requestedScale), actual:available, screen:screenScale }));
        if (screenScale !== Number(requestedScale)) throw Error('The native window did not use the requested device density');
      } else {
        await page.setViewportSize({ width: 1440, height: 1000 });
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
      if (process.env.WINDOWS_CHECK_STABLE_EDITOR_PAGES === '1') {
        const traced = await page.evaluate(async () => {
          const states = [];
          for (let i=0; i<10; i++) {
            const stack = document.querySelector('.document-stack');
            states.push([Number(stack?.dataset.pages), Boolean(document.querySelector('.busy-indicator'))]);
            if (i < 9) await new Promise(resolve => setTimeout(resolve, 500));
          }
          return states;
        });
        console.log('WINDOWS_NATIVE_EDITOR_SETTLED ' + JSON.stringify({scale:requestedScale, states:traced}));
      }
      if (process.env.WINDOWS_NATIVE_DOCX_DIR && process.env.WINDOWS_NATIVE_DOCX_AFTER_PRINT !== '1') {
        await saveTwoDocxAndReopen(app, page, item, process.env.WINDOWS_NATIVE_DOCX_DIR, name => { stage = name; });
        if (!Core) { measured++; continue; }
      }
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
      if (process.env.WINDOWS_TEST_HWP_SERIF_ALIAS === '1') {
        await page.addStyleTag({ content: `
          :is(.tiptap,.document-furniture) [data-hwp-latin-font^='["HY신명조",'] {
            font-family: "HY신명조","Studio Korean Serif","Noto Serif KR",serif !important;
          }
        ` });
        await page.evaluate(async () => {
          await document.fonts.load('15px "Noto Serif KR"', '한글');
          await document.fonts.ready;
        });
      }
      if (process.env.WINDOWS_COUNT_INK_METRICS === '1') {
        await page.evaluate(() => {
          const original = CanvasRenderingContext2D.prototype.measureText;
          const stats = window.qaCanvasInk = { calls:0, milliseconds:0 };
          CanvasRenderingContext2D.prototype.measureText = function(text) {
            if (Number(this.font.match(/(?:^| )([0-9.]+)px /)?.[1]) < 120)
              return original.call(this,text);
            const before = performance.now();
            const metrics = original.call(this,text);
            stats.calls++;
            stats.milliseconds += performance.now() - before;
            return metrics;
          };
        });
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
      if (process.env.WINDOWS_COUNT_INK_METRICS === '1')
        console.log('WINDOWS_CANVAS_CONTOUR_CALLS ' + JSON.stringify(await page.evaluate(() => window.qaCanvasInk)));
      if (process.env.WINDOWS_DIAGNOSE_WORD_LINE_BOXES === '1') {
        const lines = await require('./word-density-line-boxes.cjs')(page, item.id);
        console.log('WINDOWS_WORD_DENSITY_LINE_BOXES ' + JSON.stringify({ id:item.id, scale:requestedScale, lines }));
      }
      if (process.env.WINDOWS_CAPTURE_PRINT_FACES === '1') {
        const cdp = await page.context().newCDPSession(page);
        try {
          await cdp.send('DOM.enable');
          await cdp.send('CSS.enable');
          const { root } = await cdp.send('DOM.getDocument');
          const { nodeIds } = await cdp.send('DOM.querySelectorAll', { nodeId:root.nodeId,
            selector:'.print-layout.preview .print-page:first-of-type .print-content p' });
          const faces = [];
          for (const nodeId of nodeIds.slice(0,12)) {
            const { fonts } = await cdp.send('CSS.getPlatformFontsForNode', { nodeId });
            faces.push(fonts.map(font=>[font.familyName,font.postScriptName,font.isCustomFont,font.glyphCount]));
          }
          console.log('WINDOWS_PRINT_ACTUAL_FACES ' + JSON.stringify(faces));
        } finally { await cdp.detach(); }
      }
      if (process.env.WINDOWS_ENGLISH_TERMINAL_DETAIL === '1')
        await require('./hwp-column-terminal-witness.cjs')(page, item.path);
      if (process.env.WINDOWS_MATH_EQUATIONS === '1')
        await require('./math-equation-density.cjs').record(page, process.env.WINDOWS_MATH_EQUATIONS_RESULT);
      if (process.env.WINDOWS_FIXED_PRINT_CONTENT_AUDIT)
        await require('./fixed-print-content-audit.cjs')(page);
      if (process.env.WINDOWS_NATIVE_PRINT_GEOMETRY) {
        stage = 'pdf-geometry';
        const geometry = await require('./print-geometry.cjs')(page);
        await fs.writeFile(process.env.WINDOWS_NATIVE_PRINT_GEOMETRY, JSON.stringify(geometry));
      }
      if (process.env.WINDOWS_WORD_WRAP_WIDTH === '1') {
        stage = 'pdf-wrap-width';
        const width = await require('./word-wrap-width-witness.cjs')(page);
        console.log('WINDOWS_WORD_WRAP_WIDTH ' + JSON.stringify({ scale: requestedScale, ...width }));
      }
      if (process.env.WINDOWS_WORD_LINE_BREAKS) {
        stage = 'pdf-line-breaks';
        const lines = await require('./word-print-line-breaks.cjs')(page);
        await fs.writeFile(process.env.WINDOWS_WORD_LINE_BREAKS, JSON.stringify(lines));
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
      if (process.env.WINDOWS_NATIVE_DOCX_DIR && process.env.WINDOWS_NATIVE_DOCX_AFTER_PRINT === '1') {
        stage = 'close-print-preview-before-docx';
        const preview = page.locator('.print-preview-controls');
        if (await preview.isVisible()) await preview.getByRole('button', { name:'편집으로 돌아가기', exact:true }).click();
        await saveTwoDocxAndReopen(app, page, item, process.env.WINDOWS_NATIVE_DOCX_DIR, name => { stage = name; });
      }
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
