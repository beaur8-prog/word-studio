const fs = require('node:fs/promises'), path = require('node:path'), crypto = require('node:crypto');
const R = path.resolve(__dirname, '../../..'), B = path.dirname(__dirname);
const req = require('node:module').createRequire(R + '/package.json');
const { _electron, expect: baseExpect } = req('@playwright/test');
const expect = baseExpect.configure({ timeout: 90000 });
const sha = data => crypto.createHash('sha256').update(data).digest('hex');
(async () => {
  const build = JSON.parse(await fs.readFile(B + '/native-build-final-5c3bd56.json'));
  const out = B + '/native-footer-two-saves-5c3bd56'; await fs.mkdir(out, { recursive: true });
  const profile = await fs.mkdtemp('/private/tmp/ws-native-footer-'), entry = profile + '/entry.cjs';
  await fs.writeFile(entry, `const{app,BrowserWindow,systemPreferences}=require('electron');const old=systemPreferences.setUserDefault.bind(systemPreferences);systemPreferences.setUserDefault=(k,...a)=>k==='NSOSPLastRootDirectory'?undefined:old(k,...a);app.whenReady().then(()=>app.setActivationPolicy('prohibited'));BrowserWindow.prototype.show=function(){};BrowserWindow.prototype.focus=function(){};app.on('browser-window-created',(_e,w)=>{w.setSkipTaskbar(true);w.webContents.setBackgroundThrottling(false)});app.whenReady().then(()=>app.dock?.hide());require(${JSON.stringify(build.app + '/electron/main.cjs')});`);
  const env = { ...process.env, STUDIO_DISABLE_UPDATES: '1' }; delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const app = await _electron.launch({ executablePath: req('electron'), args: [entry, '--user-data-dir=' + profile + '/data'], env, timeout: 60000 });
  const close = await req('./scripts/native-smoke-lifetime.cjs')(app, 'Native Word footer two saves');
  const errors = [], states = [], saves = []; let page;
  try {
    page = await app.firstWindow(); await page.setViewportSize({ width: 1440, height: 1000 });
    page.setDefaultTimeout(90000); page.on('pageerror', e => errors.push(String(e)));
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await editor.waitFor();
    const baseline = JSON.parse(await fs.readFile(B + '/baseline.json'));
    let file = R + '/' + baseline.rows.find(r => r.id === '07-characters').input;
    const originalBytes = await fs.readFile(file), originalSHA = sha(originalBytes);
    for (let pass = 0; pass < 3; pass++) {
      await app.evaluate(({ dialog }, file) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] }); }, file);
      await page.keyboard.press('ControlOrMeta+o');
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(path.basename(file, '.docx'));
      await expect(page.locator('.busy-indicator')).toBeHidden(); await page.evaluate(() => document.fonts.ready);
      await expect.poll(() => page.locator('.document-stack').first().getAttribute('data-pages')).toBe('4');
      const before = await editor.evaluate(el => el.editor.getJSON());
      await editor.locator('p').filter({ hasText: /\S/ }).first().click(); await page.keyboard.press('End'); await page.keyboard.insertText(' NATIVE FOOTER CHECK');
      await expect(editor).toContainText('NATIVE FOOTER CHECK'); await page.keyboard.press('ControlOrMeta+z');
      await expect.poll(() => editor.evaluate(el => el.editor.getJSON())).toEqual(before);
      await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor).toContainText('NATIVE FOOTER CHECK'); await page.keyboard.press('ControlOrMeta+z');
      await expect.poll(() => editor.evaluate(el => el.editor.getJSON())).toEqual(before);
      const geometry = [];
      for (const zoom of [100, 50, 125, 100]) {
        const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
        await expect.poll(() => page.locator('.document-stack').first().getAttribute('data-pages')).toBe('4');
        const blank = page.locator('.paper-footer .document-furniture p[data-word-paragraph-style]:empty').first();
        await expect.poll(() => blank.evaluate(el => el.getBoundingClientRect().height)).toBeGreaterThan(5 * zoom / 100);
        geometry.push({ zoom, blankHeight: await blank.evaluate(el => el.getBoundingClientRect().height), pages: 4 });
      }
      await page.getByRole('button', { name: '보기', exact: true }).click(); await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
      await expect.poll(() => page.locator('.print-page').count()).toBe(4);
      const furniture = await page.locator('.print-footer .document-furniture').first().evaluate(el => [...el.querySelectorAll('p')].map(p => ({ text: p.textContent, height: p.getBoundingClientRect().height })));
      expect(furniture).toHaveLength(2); expect(furniture[1].text).toBe(''); expect(furniture[1].height).toBeGreaterThan(5);
      await fs.writeFile(out + '/pass-' + pass + '.pdf', Buffer.from(await app.evaluate(async ({ BrowserWindow }) => (await BrowserWindow.getAllWindows()[0].webContents.printToPDF({ printBackground: true, preferCSSPageSize: true })).toString('base64')), 'base64'));
      states.push({ pass, geometry, furniture, editorPrintPages: 4, undoExact: true });
      await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
      if (pass < 2) {
        file = out + '/textbox-save-native-footer-' + (pass + 1) + '.docx';
        await app.evaluate(({ dialog }, file) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: file }); }, file);
        await app.evaluate(({ Menu }) => { const item = Menu.getApplicationMenu().items.flatMap(i => i.submenu?.items || []).find(i => i.label.startsWith('Export Word Document')); if (!item) throw Error('DOCX export menu missing'); item.click(); });
        await expect.poll(() => fs.stat(file).then(s => s.size).catch(() => 0)).toBeGreaterThan(0); await expect(page.locator('.busy-indicator')).toBeHidden();
        const bytes = await fs.readFile(file), zip = await req('jszip').loadAsync(bytes);
        const footer = await zip.file('word/footer1.xml').async('string');
        expect((footer.match(/<w:p[ >]/g) || []).length).toBe(2); expect(footer).toContain('PAGE');
        saves.push({ file: path.relative(B, file), sha256: sha(bytes), bytes: bytes.length, footerParagraphs: 2 });
      }
    }
    expect(errors).toEqual([]); expect(sha(await fs.readFile(R + '/' + baseline.rows.find(r => r.id === '07-characters').input))).toBe(originalSHA);
    const result = { sourceCommit: build.ref, method: 'Hidden isolated compiled native app; actual OS-open IPC and application menu DOCX export IPC. OS selectors mocked. Original/two real saves re-opened, editing/undo/redo, 100/50/125/100 percent, actual print PDFs. Not Microsoft Word.', originalSHA, states, saves, errors };
    await fs.writeFile(out + '/result.json', JSON.stringify(result, null, 2)); console.log(JSON.stringify({ states: states.length, saves: saves.length, errors }));
  } catch (e) { if (page) { await page.screenshot({ path: out + '/failure.png' }).catch(() => {}); await fs.writeFile(out + '/failure-ui.txt', await page.locator('body').innerText().catch(() => '')); } throw e; }
  finally { await close(); await fs.rm(profile, { recursive: true, force: true }); }
})().catch(e => { console.error(e); process.exitCode = 1; });
