"use strict";
const path = require('node:path'), fs = require('node:fs/promises'), assert = require('node:assert/strict');
const { _electron: electron, expect } = require('@playwright/test');
const { Document, Packer, Paragraph, Table, TableRow, TableCell } = require('docx');
const root = path.resolve(__dirname, '..');
async function run() {
  const directory = process.env.WORD_ZOOM_QA_OUTPUT ? path.resolve(process.env.WORD_ZOOM_QA_OUTPUT) : await fs.mkdtemp(path.join(root, '.cache', 'zoom-desktop-'));
  await fs.mkdir(directory, { recursive: true });
  // Keep the QA app hidden even on a locked Mac. Only this disposable process
  // and its own profile are affected; no product visibility setting is changed.
  const bootstrap = path.join(directory, 'hidden-main.cjs');
  await fs.writeFile(bootstrap, `const {app,BrowserWindow}=require('electron');BrowserWindow.prototype.show=function(){};BrowserWindow.prototype.showInactive=function(){};app.whenReady().then(()=>app.dock?.hide());require(${JSON.stringify(path.join(root, 'electron/main.cjs'))});`);
  const executablePath = path.join(root, 'node_modules/electron/dist', process.platform === 'darwin' ? 'Electron.app/Contents/MacOS/Electron' : process.platform === 'win32' ? 'electron.exe' : 'electron');
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE; delete env.WORD_STUDIO_DEV;
  const app = await electron.launch({ executablePath, args: [bootstrap, '--user-data-dir='+path.join(directory, 'profile')], env });
  const close = await require('./native-smoke-lifetime.cjs')(app, 'Word pointer zoom');
  try {
    await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().forEach(w => w.webContents.setBackgroundThrottling(false)));
    const page = await app.firstWindow(); await page.setViewportSize({ width: 1440, height: 1000 });
    const errors = []; page.on('pageerror', error => errors.push(error.message));
    const bytes = await Packer.toBuffer(new Document({ sections: [{ children: Array.from({ length: 3 }, (_, i) => [
      new Paragraph({ text: `NATIVE ZOOM PAGE ${i+1} 한글`, pageBreakBefore: i>0 }),
      new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph('KEPT TABLE CELL')] })] })] }),
    ]).flat() }] }));
    await page.locator('input[type=file]').first().setInputFiles({ name: 'native-zoom.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes });
    await expect(page.locator('.tiptap')).toContainText('NATIVE ZOOM PAGE 3');
    const editorScale = () => page.locator('.scaled-document').evaluate(el => new DOMMatrix(getComputedStyle(el).transform).a);
    async function wheel(view) {
      const r = await view.boundingBox(); await page.mouse.move(r.x+r.width/2,r.y+r.height/2);
      await page.keyboard.down('Control'); try { await page.mouse.wheel(0,-120); } finally { await page.keyboard.up('Control'); }
    }
    await wheel(page.locator('.document-workspace')); await expect.poll(editorScale).toBeGreaterThan(1.1);
    const cdp = await page.context().newCDPSession(page), r = await page.locator('.document-workspace').boundingBox(), before = await editorScale();
    await cdp.send('Input.synthesizePinchGesture', {x:r.x+r.width/2,y:r.y+r.height/2,scaleFactor:1.3,gestureSourceType:'mouse'}); await cdp.detach();
    await expect.poll(editorScale).toBeGreaterThan(before*1.1);
    assert.equal(await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].webContents.getZoomFactor()), 1);
    assert.equal(await page.evaluate(() => window.visualViewport.scale), 1);
    await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();
    const preview = page.locator('.print-layout.preview'); await expect(preview).toBeVisible();
    async function pdf(name) {
      await page.evaluate(async () => {await document.fonts.ready;await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));});
      const base64 = await app.evaluate(async ({ BrowserWindow }) => {
        const window=BrowserWindow.getAllWindows()[0], color=window.getBackgroundColor(); window.setBackgroundColor('#ffffff');
        const key=await window.webContents.insertCSS('@media print{:root{color-scheme:only light!important;background:#fff!important}}',{cssOrigin:'user'});
        try { return (await window.webContents.printToPDF({printBackground:true,preferCSSPageSize:true,displayHeaderFooter:false,margins:{top:0,right:0,bottom:0,left:0}})).toString('base64'); }
        finally {await window.webContents.removeInsertedCSS(key);window.setBackgroundColor(color);}
      });
      const buffer=Buffer.from(base64,'base64');assert.equal(buffer.subarray(0,4).toString(),'%PDF');await fs.writeFile(path.join(directory,name+'.pdf'),buffer);
      return (buffer.toString('latin1').match(/\/Type\s*\/Page\b/g)||[]).length;
    }
    const sourcePages=await preview.locator(':scope > .print-page').count(), first=await pdf('before-zoom');
    await wheel(preview);await expect.poll(()=>preview.getAttribute('data-display-scale')).not.toBe('1');
    const second=await pdf('after-zoom');assert.equal(first,sourcePages);assert.equal(second,sourcePages);
    await page.screenshot({path:path.join(directory,'print-preview-zoom.png')});
    await page.getByRole('button',{name:'편집으로 돌아가기',exact:true}).click();
    await page.locator('input[type=file]').first().setInputFiles(path.join(root,'tests/fixtures/hwp/landscape-pagination.hwp'));
    const hwp=page.getByRole('dialog',{name:'HWP 읽기',exact:true});await expect(hwp.getByTestId('hwp-page')).toBeVisible();
    const hwpScale=Number(await hwp.locator('.hwp-viewport').getAttribute('data-scale'));await wheel(hwp.locator('.hwp-viewport'));
    await expect.poll(async()=>Number(await hwp.locator('.hwp-viewport').getAttribute('data-scale'))).toBeGreaterThan(hwpScale*1.1);
    await page.screenshot({path:path.join(directory,'hwp-wheel-zoom.png')});
    assert.deepEqual(errors,[]);
    const hidden=await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows().every(w=>!w.isVisible()));assert.equal(hidden,true);
    const result={platform:process.platform,hidden,isolatedProfile:true,editorWheel:true,trackpadMousePinch:true,appZoomFactor:1,hwpWheel:true,previewPages:sourcePages,pdfPagesBefore:first,pdfPagesAfter:second,errors};
    await fs.writeFile(path.join(directory,'result.json'),JSON.stringify(result,null,2)+'\n');console.log(JSON.stringify(result));
  } finally {await close();}
}
run().catch(error=>{console.error(error);process.exitCode=1;});
