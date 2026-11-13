const fs = require('node:fs/promises'), path = require('node:path'), z = require('node:zlib'), crypto = require('node:crypto');
const { createRequire } = require('node:module');
const root = path.resolve(__dirname, '../../..'), base = path.dirname(__dirname), req = createRequire(root + '/package.json');
const { _electron: electron, expect } = req('@playwright/test');
const sha = value => crypto.createHash('sha256').update(value).digest('hex');
const canonical = value => Array.isArray(value) ? value.map(canonical) : value && typeof value === 'object'
  ? Object.fromEntries(Object.keys(value).sort().map(key => [key, canonical(value[key])])) : value;
const started = Date.now(), log = value => console.log(JSON.stringify({ elapsedMs: Date.now() - started, ...value }));
function semantics(doc) {
  const images = [], tables = [], notes = [], fields = [], breaks = [], sections = [doc.page];
  const visit = node => {
    if (node.type === 'sectionBreak') sections.push(JSON.parse(node.attrs.page));
    if (node.type === 'inlineImage' || node.type === 'image') {
      const attrs = node.attrs || {}, encoded = String(attrs.src || '').split(',')[1] || '';
      images.push({ sha256: sha(Buffer.from(encoded, 'base64')), width: attrs.width, height: attrs.height,
        floating: attrs.floating ? JSON.parse(attrs.floating) : null });
    }
    if (node.type === 'table') tables.push((node.content || []).map(row => (row.content || []).map(cell => [cell.attrs?.colspan || 1, cell.attrs?.rowspan || 1])));
    if (node.type === 'noteReference' || node.type === 'documentNote') notes.push([node.type, node.attrs?.kind, node.attrs?.label]);
    if (node.type === 'documentField') fields.push(node.attrs?.kind);
    if (node.type === 'pageBreak' || node.type === 'columnBreak') breaks.push(node.type);
    if (node.type === 'hardBreak') return '\n'; if (node.type === 'documentTab') return '\t';
    if (node.type === 'documentSymbol') return node.attrs?.glyph || '';
    return (node.text || '') + (node.content || []).map(visit).join('');
  };
  const text = visit(doc.content); return { text, images, tables, notes, fields, breaks, sections };
}
(async () => {
  const out = base + '/' + (process.env.NATIVE_QA_STAGE || 'native-sop-eee8aa0'); await fs.mkdir(out, { recursive: true });
  const manifest = JSON.parse(await fs.readFile(root + '/analysis/word-layout-8h-2026-11-10/native-government-manifest.json'));
  const row = manifest.rows.find(r => r.id === 'mois-sop-3'), input = await fs.readFile(row.path);
  if (sha(input) !== row.sha256) throw Error('Frozen SOP input changed');
  const profile = await fs.mkdtemp('/private/tmp/ws-layout-12h-sop-'), original = profile + '/mois-sop-3.docx';
  await fs.writeFile(original, input);
  const entry = profile + '/entry.cjs';
  await fs.writeFile(entry, `const{app,BrowserWindow,systemPreferences}=require('electron');const old=systemPreferences.setUserDefault.bind(systemPreferences);systemPreferences.setUserDefault=(k,...a)=>k==='NSOSPLastRootDirectory'?undefined:old(k,...a);app.whenReady().then(()=>app.setActivationPolicy('prohibited'));BrowserWindow.prototype.show=function(){};BrowserWindow.prototype.focus=function(){};app.on('browser-window-created',(_e,w)=>{w.setSkipTaskbar(true);w.webContents.setBackgroundThrottling(false)});app.whenReady().then(()=>app.dock?.hide());require(${JSON.stringify((process.env.NATIVE_QA_APP_DIR || base + '/native-app') + '/electron/main.cjs')});`);
  const env = { ...process.env, STUDIO_DISABLE_UPDATES: '1' }; delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const app = await electron.launch({ executablePath: req('electron'), args: [entry, '--force-device-scale-factor=1', '--user-data-dir=' + profile + '/data'], env, timeout: 60000 });
  const child = app.process(); const watchdog = setTimeout(() => { log({ failure: 'Owned native QA exceeded its total deadline' }); child.kill('SIGKILL'); process.exitCode = 1; }, Number(process.env.NATIVE_TOTAL_TIMEOUT_MS || 2400000)); watchdog.unref();
  let page, checkpointModel; const errors = [], phases = [], saves = []; let reference;
  try {
    const actualUserData = await app.evaluate(({ app }) => app.getPath('userData'));
    if (path.resolve(actualUserData) !== path.resolve(profile + '/data')) throw Error('QA native profile was not isolated');
    log({ phase: 'owned-native-profile-confirmed', actualUserData });
    page = await app.firstWindow(); await page.setViewportSize({ width: 1440, height: 1000 }); page.setDefaultTimeout(120000);
    page.on('pageerror', e => { errors.push(String(e)); log({ rendererError: String(e) }); });
    await app.evaluate(({ dialog }) => { const original = dialog.showMessageBoxSync.bind(dialog); dialog.showMessageBoxSync = (...args) => args.at(-1)?.title === 'Close Word Studio?' ? 1 : original(...args); });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await editor.waitFor();
    const modelHash = () => editor.evaluate(el => el.editor.getJSON()).then(model => sha(JSON.stringify(canonical(model))));
    const snapshot = () => page.evaluate(() => {
      const root = document.querySelector('.tiptap[contenteditable=true]'), stack = root.closest('.document-stack'), origin = root.getBoundingClientRect();
      let scale = 1; for (let a = root; a; a = a.parentElement) { const c = getComputedStyle(a); scale *= parseFloat(c.zoom) || 1; if (c.transform !== 'none') { const m = new DOMMatrixReadOnly(c.transform); scale *= Math.hypot(m.a, m.b); } }
      const number = v => Math.round(v / scale * 64) / 64;
      const rect = el => { const b = el.getBoundingClientRect(); return b.width < .1 && b.height < .1 ? [0, 0, 0, 0] : [number(b.x - origin.x), number(b.y - origin.y), number(b.width), number(b.height)]; };
      const excluded = '.section-papers,.paper-backgrounds,.pagination-headers,.document-furniture';
      const paragraphs = [...root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].filter(el => !el.closest(excluded)).map(rect);
      const images = [...root.querySelectorAll('img[data-image-floating]')].filter(el => !el.closest(excluded)).map(rect);
      const gaps = [...root.querySelectorAll('.pagination-gap,.pagination-cell-gap')].map(rect);
      const meaningful = [...root.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].filter(el => !el.closest(excluded) && (el.textContent.trim() || [...el.querySelectorAll('img')].some(i => i.getBoundingClientRect().height > .1)));
      const stride = Number(stack.dataset.pageHeight) + 24, physicalParagraphPages = Math.max(1, ...meaningful.map(p => Math.floor((rect(p)[1] + .5) / stride) + 1));
      const metric = root.editor.state.plugins.find(p => p.key.startsWith('hwpCellLineMetrics')), metricDecorations = metric?.getState(root.editor.state)?.find().length || 0;
      const sections = [...root.querySelectorAll('.document-section-content')].map(s => ({ pages: s.dataset.pages, top: s.dataset.margin, bottom: s.dataset.bottomMargin, height: s.dataset.pageHeight }));
      const data = JSON.stringify([stack.dataset.pages, number(origin.height), paragraphs, images, gaps, sections]); let hash = 2166136261;
      for (let i = 0; i < data.length; i++) hash = Math.imul(hash ^ data.charCodeAt(i), 16777619);
      const bodyPictures = [...root.querySelectorAll('img[data-image-floating]')].filter(el=>!el.closest(excluded));
      const floatingStarts = bodyPictures.flatMap((img,index)=>img.closest('p,h1,h2,h3,h4,h5,h6')?.hasAttribute('data-floating-start-page') ? [index] : []);
      return { floatingStarts, fontsStatus:document.fonts.status, pages: Number(stack.dataset.pages), physicalParagraphPages, metricDecorations, sections, height: number(origin.height), scale, signature: (hash >>> 0).toString(16), images, paragraphs: paragraphs.length, gaps: gaps.length, paragraphRects: paragraphs, gapRects: gaps };
    });
    async function stable(label) {
      const samples = [], start = Date.now(); let previous, stableSince = Date.now();
      while (Date.now() - start < 180000) {
        await page.waitForTimeout(1000); const state = await snapshot(), { paragraphRects, gapRects, ...compact } = state; samples.push({ elapsedMs: Date.now() - start, ...compact });
        if (state.signature !== previous) { previous = state.signature; stableSince = Date.now(); }
        if (samples.length % 10 === 0) log({ phase: label, sample: samples.length, pages: state.pages, signature: state.signature, stableMs: Date.now() - stableSince });
        if (Date.now() - stableSince >= Number(process.env.NATIVE_STABLE_MS || 40000)) {
          const result = { label, samples, final: state, stableMs: Date.now() - stableSince };
          phases.push(result); await fs.writeFile(out + '/' + label + '.json.gz', z.gzipSync(JSON.stringify(result)));
          if (errors.length) throw Error('Renderer error: ' + errors.join('; '));
          if (process.env.NATIVE_STRICT_FRAMES === '1' && (state.pages !== state.physicalParagraphPages || !state.metricDecorations))
            throw Error('Physical paper count or cell metrics are incomplete: ' + JSON.stringify({pages:state.pages, physical:state.physicalParagraphPages, metrics:state.metricDecorations}));
          return state;
        }
      }
      await fs.writeFile(out + '/' + label + '-failure.json.gz', z.gzipSync(JSON.stringify(samples))); throw Error('No 40-second geometry convergence: ' + label);
    }
    async function saveBackup(label) {
      const target = profile + '/' + label + '.wordstudio';
      await app.evaluate(({ dialog }, file) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: file }); }, target);
      await page.keyboard.press('ControlOrMeta+s'); await expect.poll(() => fs.stat(target).then(s => s.size).catch(() => 0), { timeout: 120000 }).toBeGreaterThan(0);
      await expect(page.locator('.busy-indicator')).toBeHidden(); const bytes = await fs.readFile(target);
      await fs.writeFile(out + '/' + label + '.wordstudio.gz', z.gzipSync(bytes)); return JSON.parse(bytes);
    }
    let file = original;
    for (let pass = 0; pass < Number(process.env.NATIVE_PASSES || 3); pass++) {
      await app.evaluate(({ dialog }, input) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [input] }); }, file);
      await page.keyboard.press('ControlOrMeta+o'); await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(file, '.docx'), { timeout: 120000 }); await page.waitForFunction(() => (document.querySelector('.tiptap[contenteditable=true]')?.textContent || '').length > 50000, {}, { timeout: 120000 });
      await expect(page.locator('.busy-indicator')).toBeHidden(); await page.evaluate(() => document.fonts.ready);
      const originalState = await stable('pass-' + pass + '-original'), document = await saveBackup('pass-' + pass + '-backup');
      const content = semantics(document); if (!reference) reference = content;
      if (content.text !== reference.text || JSON.stringify(content.tables) !== JSON.stringify(reference.tables) || JSON.stringify(content.notes) !== JSON.stringify(reference.notes) || JSON.stringify(content.fields) !== JSON.stringify(reference.fields) || JSON.stringify(content.breaks) !== JSON.stringify(reference.breaks) || content.sections.length !== reference.sections.length) throw Error('Source text/table/note/field/break preservation failed at pass ' + pass);
      if (JSON.stringify(content.images.map(i => i.sha256)) !== JSON.stringify(reference.images.map(i => i.sha256))) throw Error('Source image bytes/order changed at pass ' + pass);
      checkpointModel = await editor.evaluate(el => el.editor.getJSON());
      await fs.writeFile(out + '/pass-' + pass + '-before.json.gz', z.gzipSync(JSON.stringify(checkpointModel)));
      const before = await modelHash();
      await editor.evaluate((el, { target, segment }) => {
        let start = 0, limit = el.editor.state.doc.content.size, frameIndex = -1;
        if (target >= 0) el.editor.state.doc.forEach((node, pos) => { if (node.type.name === 'sectionContent' && ++frameIndex === target) {
          start = pos; limit = pos + node.nodeSize;
          if (segment >= 0) { let scope = 0, found = false; node.forEach((child, offset) => { if (child.type.name === 'pageBreak') scope++; if (!found && scope === segment && child.textContent.trim()) { start = pos + 1 + offset; limit = start + child.nodeSize; found = true; } }); if (!found) throw Error('Requested QA source segment missing'); }
        } });
        let end; el.editor.state.doc.descendants((node, pos) => { if (end === undefined && pos >= start && pos < limit && node.isTextblock && node.textContent.trim()) end = pos + 1 + node.content.size; });
        el.editor.commands.setTextSelection(end); el.editor.commands.focus();
      }, { target: Number(process.env.NATIVE_EDIT_SECTION ?? -1), segment: Number(process.env.NATIVE_EDIT_BREAK_SEGMENT ?? -1) });
      await page.keyboard.insertText(' NATIVE PAGE QA'); await expect(editor).toContainText(' NATIVE PAGE QA'); await stable('pass-' + pass + '-edited');
      await page.keyboard.press('ControlOrMeta+z'); await page.waitForTimeout(300);
      await fs.writeFile(out + '/pass-' + pass + '-after-undo.json.gz', z.gzipSync(JSON.stringify(await editor.evaluate(el => el.editor.getJSON()))));
      await expect.poll(modelHash, { timeout: 90000, intervals: [250, 500, 1000] }).toBe(before); const undone = await stable('pass-' + pass + '-undo');
      let undoMaxDelta = 0;
      if (undone.pages !== originalState.pages) throw Error('Undo changed page count at pass ' + pass);
      for (const key of ['paragraphRects', 'gapRects', 'images']) {
        if (undone[key].length !== originalState[key].length) throw Error('Undo changed geometry item count: ' + key);
        for (let index = 0; index < undone[key].length; index++) for (let axis = 0; axis < 4; axis++)
          undoMaxDelta = Math.max(undoMaxDelta, Math.abs(undone[key][index][axis] - originalState[key][index][axis]));
      }
      // Chromium may choose either adjacent 1/64px table-layout unit after
      // equivalent DOM reconstruction. Preserve every page and object, and
      // reject changes larger than two layout units.
      if (undoMaxDelta > 1/32) throw Error('Undo changed original geometry by ' + undoMaxDelta + 'px at pass ' + pass);
      log({ phase: 'undo-geometry-restored', pass, undoMaxDelta });
      if (process.env.NATIVE_UNDO_ONLY === '1') break;
      await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText(' NATIVE PAGE QA'); await stable('pass-' + pass + '-redo');
      await page.keyboard.press('ControlOrMeta+z'); await expect.poll(modelHash, { timeout: 90000, intervals: [250, 500, 1000] }).toBe(before); await stable('pass-' + pass + '-undo-again');
      for (const zoom of [50, 125, 100]) {
        const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
        const state = await stable('pass-' + pass + '-zoom-' + zoom);
        if (state.pages !== originalState.pages) throw Error('Zoom changed pagination: ' + pass + '/' + zoom + ': ' + originalState.pages + ' -> ' + state.pages);
      }
      await saveBackup('pass-' + pass + '-final');
      if (pass < 2) {
        file = out + '/textbox-save-sop-' + (pass + 1) + '.docx';
        await app.evaluate(({ dialog }, target) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: target }); }, file);
        await app.evaluate(({Menu}) => { const item = Menu.getApplicationMenu().items.flatMap(item => item.submenu?.items || []).find(item => item.label.startsWith('Export Word Document')); if (!item) throw Error('Native DOCX menu missing'); item.click(); }); const exportStartedAt = Date.now(); let pendingSamples = 0;
        await expect.poll(async () => {
          const bytes = await fs.stat(file).then(s => s.size).catch(() => 0);
          if (!bytes && ++pendingSamples % 20 === 0) {
            let probeTimer;
            const state = await Promise.race([
              page.evaluate(() => ({ busy: document.querySelector('.busy-indicator')?.textContent,
                fonts: document.fonts.status, measuringPages: document.querySelectorAll('.print-layout.measuring .print-page').length,
                measuringCells: document.querySelectorAll('.print-layout.measuring [data-hwp-export-line]').length,
                lastMeasuredCell: [...document.querySelectorAll('[data-hwp-export-line]')].at(-1)?.getAttribute('data-hwp-export-line') })),
              new Promise(resolve => { probeTimer = setTimeout(() => resolve({ diagnostic: 'Renderer did not answer the read-only probe in 2 seconds' }), 2000); }),
            ]).finally(() => clearTimeout(probeTimer));
            log({ phase: 'native-docx-export-pending', pass: pass + 1, exportMs: Date.now() - exportStartedAt, ...state });
          }
          return bytes;
        }, { timeout: Number(process.env.NATIVE_EXPORT_TIMEOUT_MS || 300000), intervals: [1000] }).toBeGreaterThan(0);
        await expect(page.locator('.busy-indicator')).toBeHidden(); const saved = await fs.readFile(file); if (saved.subarray(0, 2).toString() !== 'PK') throw Error('Native DOCX is not a ZIP');
        saves.push({ file, sha256: sha(saved), bytes: saved.length }); log({ phase: 'native-docx-saved', pass: pass + 1, bytes: saved.length });
      }
    }
    await page.getByRole('button', { name: '보기', exact: true }).click(); await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    await expect.poll(() => page.locator('.print-page').count(), { timeout: 180000 }).toBeGreaterThan(0);
    const printPages = await page.locator('.print-page').count();
    const pdf = await app.evaluate(async ({ BrowserWindow }) => (await BrowserWindow.getAllWindows()[0].webContents.printToPDF({ printBackground: true, preferCSSPageSize: true })).toString('base64'));
    await fs.writeFile(out + '/sop-save2-native.pdf', Buffer.from(pdf, 'base64'));
    await fs.writeFile(out + '/result.json', JSON.stringify({ sourceRef: process.env.NATIVE_BUILD_COMMIT || 'eee8aa0', editTarget: { frame:process.env.NATIVE_EDIT_SECTION, sourceBreakSegment:process.env.NATIVE_EDIT_BREAK_SEGMENT }, originalSHA256: sha(input), sourcePreserved: sha(await fs.readFile(row.path)) === sha(input), hidden: !(await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().some(w => w.isVisible()))), method: 'Isolated compiled Electron with actual native open/save IPC, keyboard editing, and the native application Export Word menu action. OS file chooser mocked; no user profile or source input overwritten. Each phase requires the configured continuous stability interval of identical physical paragraph/image/gap geometry; CSS class and attribute order are not geometry. Microsoft Word not newly measured.', requiredStableMs: Number(process.env.NATIVE_STABLE_MS || 40000), phases: phases.map(({ label, final, stableMs }) => ({ label, final, stableMs })), saves, printPages, bodyLength: reference.text.length, bodySHA256: sha(reference.text), images: reference.images.length, tables: reference.tables.length, sourceSections: reference.sections, errors }, null, 2));
    log({ complete: true, printPages, phases: phases.length, images: reference.images.length, saves });
  } catch (error) {
    if (page) await fs.writeFile(out + '/failure-model.json.gz', z.gzipSync(JSON.stringify(await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).evaluate(el => el.editor.getJSON()).catch(() => null))));
    await fs.writeFile(out + '/failure.txt', String(error) + '\n' + (error.stack || ''));
    if (page) { await page.screenshot({ path: out + '/failure.png' }).catch(() => {}); await fs.writeFile(out + '/failure-ui.txt', await page.locator('body').innerText().catch(() => '')); }
    throw error;
  } finally {
    clearTimeout(watchdog);
    let timer; try { await Promise.race([app.close(), new Promise((_, reject) => { timer = setTimeout(() => reject(Error('Owned native QA did not close')), 15000); })]); }
    catch { if (child.exitCode === null) child.kill('SIGKILL'); }
    finally { clearTimeout(timer); await fs.rm(profile, { recursive: true, force: true }); }
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
