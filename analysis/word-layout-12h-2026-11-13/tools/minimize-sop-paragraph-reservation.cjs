const fs = require('node:fs/promises');
const path = require('node:path');
const z = require('node:zlib');
const crypto = require('node:crypto');
const { createRequire } = require('node:module');
const R = path.resolve(__dirname, '../../..'), B = path.dirname(__dirname), req = createRequire(R + '/package.json');
const { _electron: electron, expect: baseExpect } = req('@playwright/test');
const expect = baseExpect.configure({ timeout: 120000 });
const hash = bytes => crypto.createHash('sha256').update(bytes).digest('hex');
const text = node => (node.text || '') + (node.content || []).map(text).join('');
(async () => {
  const build = JSON.parse(await fs.readFile(B + '/native-checkpoint-build.json'));
  if (build.sourceCommit !== '6bc49fae11068d43b258f5568ceccd3c3cb8da30') throw Error('Frozen source mismatch');
  const originalBytes = z.gunzipSync(await fs.readFile(B + '/native-sop-6bc49fa-cover-retry/pass-0-final.wordstudio.gz'));
  const original = JSON.parse(originalBytes), root = original.content;
  const body = root.content.find(n => n.type === 'sectionContent' && (n.content || []).some(c => text(c).includes('재해 복구 : 재해 발생 시 데이터 복구')));
  if (!body) throw Error('Source body frame missing');
  const target = body.content.find(n => n.type === 'paragraph' && text(n).includes('- 재해 복구 : 재해 발생 시 데이터 복구'));
  if (!target || target.content.some(n => n.type !== 'text')) throw Error('Expected source plain text paragraph');
  const targetText = text(target), at = body.content.indexOf(target);
  let begin = at, end = at + 1;
  while (begin > 0 && body.content[begin - 1].type !== 'pageBreak') begin--;
  while (end < body.content.length && body.content[end].type !== 'pageBreak') end++;
  const definitions = [{ id: 'full-control', full: true }, { id: 'source-segment-only', prefix: 0 },
    ...[16, 64, 256, 512, 768].map(prefix => ({ id: 'source-prefix-' + prefix, prefix })),
    { id: 'body-without-cover', prefix: begin },
    { id: 'source-segment-with-cover', prefix: 0, keepCover: true },
    { id: 'source-segment-with-cover-first-block', prefix: 0, keepCover: true, firstCoverBlock: true },
    { id: 'body-with-first-cover-block', prefix: begin, bodyAll: true, keepCover: true, firstCoverBlock: true }];
  const stage = B + '/' + (process.env.MINIMIZE_STAGE || 'minimal-sop-paragraph-reservation'); await fs.mkdir(stage, { recursive: true });
  const profile = await fs.mkdtemp('/private/tmp/ws-layout-final-minimize-'), entry = profile + '/entry.cjs';
  await fs.writeFile(entry, `const{app,BrowserWindow,systemPreferences}=require('electron');const old=systemPreferences.setUserDefault.bind(systemPreferences);systemPreferences.setUserDefault=(k,...a)=>k==='NSOSPLastRootDirectory'?undefined:old(k,...a);app.whenReady().then(()=>app.setActivationPolicy('prohibited'));BrowserWindow.prototype.show=function(){};BrowserWindow.prototype.focus=function(){};app.on('browser-window-created',(_e,w)=>{w.setSkipTaskbar(true);w.webContents.setBackgroundThrottling(false)});app.whenReady().then(()=>app.dock?.hide());require(${JSON.stringify(build.app + '/electron/main.cjs')});`);
  const env = { ...process.env, STUDIO_DISABLE_UPDATES: '1' }; delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const app = await electron.launch({ executablePath: req('electron'), args: [entry, '--force-device-scale-factor=1', '--user-data-dir=' + profile + '/data'], env, timeout: 60000 });
  const timer = setTimeout(() => { app.process().kill('SIGKILL'); process.exitCode = 1; }, 25 * 60000); timer.unref();
  const reports = [], errors = [];
  try {
    if (await app.evaluate(({ app }) => app.getPath('userData')) !== profile + '/data') throw Error('Profile was not isolated');
    const page = await app.firstWindow(); await page.setViewportSize({ width: 1440, height: 1000 });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await editor.waitFor();
    page.on('pageerror', e => errors.push(String(e)));
    for (const definition of definitions) {
      if (process.env.MINIMIZE_IDS && !process.env.MINIMIZE_IDS.split(',').includes(definition.id)) continue;
      const doc = structuredClone(original); doc.title = definition.id;
      if (!definition.full) {
        const frame = structuredClone(body);
        frame.content = definition.bodyAll || definition.id === 'body-without-cover' ? structuredClone(body.content) : structuredClone(body.content.slice(Math.max(0, begin - definition.prefix), end));
        const frames = [frame];
        if (definition.keepCover) { const cover = structuredClone(root.content[0]); if (definition.firstCoverBlock) cover.content = [cover.content[0]]; frames.unshift(cover); }
        doc.content = { ...structuredClone(root), content: frames };
      }
      const bytes = Buffer.from(JSON.stringify(doc)), file = profile + '/' + definition.id + '.wordstudio';
      await fs.writeFile(file, bytes); await fs.writeFile(stage + '/' + definition.id + '.wordstudio.gz', z.gzipSync(bytes));
      console.log(JSON.stringify({ id: definition.id, phase: 'open', bytes: bytes.length }));
      await app.evaluate(({ dialog }, file) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] }); }, file);
      await page.keyboard.press('ControlOrMeta+o');
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(definition.id);
      await expect(page.locator('.busy-indicator')).toBeHidden(); await page.evaluate(() => document.fonts.ready);
      const samples = []; let previous, stableSince = Date.now(); const start = Date.now();
      while (Date.now() - start < 90000) {
        await page.waitForTimeout(1500);
        const state = await editor.evaluate((root, targetText) => {
          const p = [...root.querySelectorAll('p')].find(p => !p.closest('.section-papers,.document-furniture') && p.textContent === targetText);
          if (!p) throw Error('Source target paragraph absent from derived input');
          const r = p.getBoundingClientRect(), widgets = [...p.querySelectorAll('.pagination-cell-gap,.pagination-gap,[data-flow-widget]')].map(w => ({ height: w.getBoundingClientRect().height, style: w.style.cssText }));
          const c = getComputedStyle(p), stack = root.closest('.document-stack');
          return { pages: Number(stack.dataset.pages), paragraphHeight: r.height, lineHeight: c.lineHeight, width: r.width, widgets, html: p.outerHTML };
        }, targetText);
        const signature = JSON.stringify([state.pages, state.paragraphHeight, state.widgets]);
        if (signature !== previous) { previous = signature; stableSince = Date.now(); }
        samples.push({ elapsedMs: Date.now() - start, stableMs: Date.now() - stableSince, ...state });
        if (Date.now() - stableSince >= 12000) break;
      }
      const final = samples.at(-1);
      reports.push({ definition, inputSHA256: hash(bytes), sourceTargetTextSHA256: hash(targetText), inputBytes: bytes.length,
        localTargetStableMs: final.stableMs, final, samples: samples.map(({ html, ...s }) => s) });
      await fs.writeFile(stage + '/results.json', JSON.stringify({ sourceCommit: build.sourceCommit, sourceUncompressedSHA256: hash(originalBytes),
        method: 'Read-only input minimization under frozen compiled native product. Full control plus explicit different diagnostic copies of the public source segment/prefix; fixed 29-document cohort and original files unchanged. Twelve-second LOCAL target-box convergence is diagnostic only, not the forty-second full layout acceptance test or Microsoft/page-accuracy result.',
        sourceBodyChildren: body.content.length, sourceTargetIndex: at, sourceSegment: [begin, end],
        targetSourceAttributes: target.attrs, targetSourceChildTypes: target.content.map(n => n.type), reports, errors }, null, 2));
      console.log(JSON.stringify({ id: definition.id, pages: final.pages, height: final.paragraphHeight, widgets: final.widgets, stableMs: final.stableMs }));
    }
  } finally { clearTimeout(timer); await app.close(); await fs.rm(profile, { recursive: true, force: true }); }
  if (errors.length) throw Error('Diagnostic renderer errors: ' + errors.join('; '));
})().catch(e => { console.error(e); process.exitCode = 1; });
