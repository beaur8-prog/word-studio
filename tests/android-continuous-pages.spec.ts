import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, ImageRun, SectionType, PageOrientation } from 'docx';
import JSZip from 'jszip';
import path from 'node:path';
test.beforeEach(({page})=>{page.on('pageerror',error=>console.log('PAGE ERROR',error.stack));});
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=', 'base64');
async function bytes() { return Packer.toBuffer(new Document({ sections: [{ children: [
  new Paragraph('CONTINUOUS PAGE 1'), new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph('CELL LEFT')] }), new TableCell({ children: [new Paragraph('CELL RIGHT')] })] })] }),
  new Paragraph({ children: [new ImageRun({ type: 'png', data: png, transformation: { width: 30, height: 30 } })] }),
  ...[2,3,4,5,6].map(n => new Paragraph({ text: `CONTINUOUS PAGE ${n}`, pageBreakBefore: true }))
] }] })); }
async function open(page: Page) { await page.goto('/'); await page.locator('.android-shell > input[type=file]').setInputFiles({ name: 'continuous.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await bytes() }); await expect(page.locator('.android-paper-viewport')).toBeVisible(); }
async function viewer(page: Page) { await page.locator('.android-topbar').getByRole('button', { name: '전체 화면 뷰어', exact: true }).click(); await expect(page.locator('.android-presentation-page')).toContainText('CONTINUOUS PAGE 1'); }
async function visiblePages(page: Page, selector: string, viewport: string) { return page.locator(selector).evaluateAll((elements, area) => { const r = document.querySelector(area as string)!.getBoundingClientRect(); return elements.filter(el => { const p = el.getBoundingClientRect(); return p.bottom > r.top + 4 && p.top < r.bottom - 4; }).length; }, viewport); }
async function pinch(page: Page, selector: string) {
  const r = (await page.locator(selector).boundingBox())!, x = r.x + r.width / 2, y = r.y + Math.min(r.height / 2, 160);
  const c = await page.context().newCDPSession(page);
  await c.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: x-35, y, id: 1 }, { x: x+35, y, id: 2 }] });
  await c.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: x-70, y: y-15, id: 1 }, { x: x+70, y: y+15, id: 2 }] });
  await c.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] }); await c.detach();
}

test('viewer keeps an absolute scale while paging and shows following pages when zoomed out', async ({ page }, info) => {
  await open(page); await viewer(page);
  await page.keyboard.press('-'); await page.keyboard.press('-');
  const scale = Number(await page.locator('.android-presentation').getAttribute('data-scale'));
  expect(await visiblePages(page, '.android-presentation-slot', '.android-presentation-viewport')).toBeGreaterThanOrEqual(2);
  await page.keyboard.press('ArrowRight'); await expect(page.locator('.android-presentation-page')).toContainText('CONTINUOUS PAGE 2');
  expect(Number(await page.locator('.android-presentation').getAttribute('data-scale'))).toBeCloseTo(scale, 6);
  await page.keyboard.press('ArrowRight'); await expect(page.locator('.android-presentation-page')).toContainText('CONTINUOUS PAGE 3');
  expect(Number(await page.locator('.android-presentation').getAttribute('data-scale'))).toBeCloseTo(scale, 6);
  await page.keyboard.press('ArrowLeft'); await expect(page.locator('.android-presentation-page')).toContainText('CONTINUOUS PAGE 2');
  await page.screenshot({ path: info.outputPath('viewer-continuous-phone.png') });
  await page.keyboard.press('End'); await expect(page.locator('.android-presentation-page')).toContainText('CONTINUOUS PAGE 6');
  await page.keyboard.press('Home'); await expect(page.locator('.android-presentation-page')).toContainText('CELL LEFT');
  await expect(page.locator('.android-presentation-page img')).toHaveCount(1);
  for(let n=0;n<4;n++) await page.keyboard.press('-');
  const c=await page.context().newCDPSession(page);
  await c.send('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[{x:205,y:650,id:1}]});
  for(let n=1;n<=4;n++) await c.send('Input.dispatchTouchEvent',{type:'touchMove',touchPoints:[{x:205,y:650-100*n,id:1}]});
  await c.send('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});await c.detach();
  await expect.poll(async()=>Number(await page.locator('.android-presentation').getAttribute('data-page'))).toBeGreaterThanOrEqual(3);

});

test('Escape exits the viewer once and never closes the Android app or the document', async ({ page }) => {
  await page.addInitScript(() => { (window as any).closes=0; (window as any).wordAndroid={close:()=>{(window as any).closes++},onBack:(fn:any)=>{(window as any).backNative=fn;return()=>{}}}; });
  await page.goto('/'); await page.getByRole('button', { name: '새 문서', exact: true }).click();
  const editor=page.getByLabel('Android 문서 편집 영역'); await editor.click(); await page.keyboard.type('PRESERVE ESC DOCUMENT');
  await page.locator('.android-topbar').getByRole('button', { name: '전체 화면 뷰어', exact:true }).click();
  await expect(page.locator('.android-presentation-page')).toContainText('PRESERVE ESC DOCUMENT');
  await page.keyboard.down('Escape');
  await page.evaluate(()=>{window.dispatchEvent(new Event('studio-native-escape'));(window as any).backNative();});
  await page.keyboard.up('Escape'); await expect(editor).toBeVisible(); await expect(editor).toContainText('PRESERVE ESC DOCUMENT');
  await page.evaluate(()=>window.dispatchEvent(new Event('studio-native-escape')));
  expect(await page.evaluate(()=>(window as any).closes)).toBe(0); await expect(page.locator('.android-shell')).toBeVisible();
});

test('phone editing zoom leaves logical paper and text unchanged and reveals multiple pages', async ({ page }, info) => {
  await open(page); await page.locator('.android-topbar').getByRole('button',{name:'편집 모드',exact:true}).click();
  const editor=page.getByLabel('Android 문서 편집 영역'); await expect(editor).toBeVisible();
  await expect.poll(()=>page.locator('.android-editor-stage .paper').count()).toBeGreaterThanOrEqual(6);
  const before=await editor.locator('p').first().evaluate(el=>getComputedStyle(el).fontSize);
  for(let n=0;n<3;n++) await page.getByRole('button',{name:'편집 축소',exact:true}).click();
  expect(await visiblePages(page,'.android-editor-stage .paper','.android-editor-scroll')).toBeGreaterThanOrEqual(2);
  const width=await page.locator('.android-editor-stage .document-stack').evaluate(el=>(el as HTMLElement).offsetWidth); expect(width).toBeGreaterThan(790);
  expect(await editor.locator('p').first().evaluate(el=>getComputedStyle(el).fontSize)).toBe(before);
  await page.screenshot({path:info.outputPath('editor-continuous-phone.png')});
  const scale=Number(await page.locator('.android-editor-scroll').getAttribute('data-scale')); await pinch(page,'.android-editor-scroll');
  await expect.poll(async()=>Number(await page.locator('.android-editor-scroll').getAttribute('data-scale'))).toBeGreaterThan(scale*1.6);
  await editor.click(); await page.keyboard.press('ControlOrMeta+End'); await page.keyboard.type(' PHONE ZOOM EDIT'); await expect(editor).toContainText('PHONE ZOOM EDIT');
});

test('tablet uses the actual PC ribbon and retains its editor history, multiple-page zoom and native DOCX saves', async ({ page }, info) => {
  await page.setViewportSize({width:1280,height:1000});
  await page.addInitScript(()=>{ (window as any).savedCopies=[]; (window as any).wordAndroid={saveFile:async(file:any)=>{(window as any).savedCopies.push({name:file.name,data:Array.from(file.data)});return 'completed';}}; });
  await open(page); if(await page.getByRole('button',{name:'알림 닫기',exact:true}).count()) await page.getByRole('button',{name:'알림 닫기',exact:true}).click();
  await page.locator('.android-topbar').getByRole('button',{name:'편집 모드',exact:true}).click();
  const pc=page.locator('.office-shell[data-android-desktop=true]'),editor=pc.getByLabel('문서 편집 영역');
  await expect(pc).toBeVisible(); await expect(editor).toContainText('CONTINUOUS PAGE 1');
  await expect(pc.locator('.ruler')).toBeVisible();
  for(const name of ['홈','삽입','디자인','레이아웃','검토','보기']) await expect(pc.getByRole('button',{name,exact:true}).first()).toBeVisible();
  await editor.locator('p').first().click(); await page.keyboard.press('End'); await page.keyboard.type(' TABLET EDIT');
  await pc.getByLabel('확대 축소',{exact:true}).fill('30');
  expect(await visiblePages(page,'.android-tablet-editor .paper','.android-tablet-editor .document-workspace')).toBeGreaterThanOrEqual(2);
  await page.screenshot({path:info.outputPath('tablet-pc-multiple-pages.png')});
  const scaleBefore=await pc.locator('.scaled-document').evaluate(el=>el.getBoundingClientRect().width);
  await pinch(page,'.android-tablet-editor .document-workspace');
  await expect.poll(()=>pc.locator('.scaled-document').evaluate(el=>el.getBoundingClientRect().width)).toBeGreaterThan(scaleBefore*1.6);
  await pc.getByRole('button',{name:'전체 화면 뷰어',exact:true}).click(); await expect(page.locator('.android-presentation-page')).toContainText('TABLET EDIT');
  await page.keyboard.press('Escape'); await expect(pc).toBeVisible(); await expect(editor).toHaveAttribute('contenteditable','true');
  await pc.getByRole('button',{name:'실행 취소 (Ctrl+Z)',exact:true}).click(); await expect(editor).not.toContainText('TABLET EDIT');
  await pc.getByRole('button',{name:'다시 실행 (Ctrl+Y)',exact:true}).click(); await expect(editor).toContainText('TABLET EDIT');
  await pc.getByRole('button',{name:'Word 문서 사본 저장 (.docx)',exact:true}).click();
  await expect.poll(()=>page.evaluate(()=>(window as any).savedCopies.length)).toBe(1);
  const saved=await page.evaluate(()=>(window as any).savedCopies[0]);const zip=await JSZip.loadAsync(Buffer.from(saved.data));
  expect(await zip.file('word/document.xml')!.async('string')).toContain('TABLET EDIT');
  expect(Object.keys(zip.files).filter(p=>p.startsWith('word/media/')&&!zip.files[p].dir)).toHaveLength(1);
  // Reopen the actual native-save bytes and save again with the reused desktop UI.
  await page.locator('.android-shell > input[type=file]').setInputFiles({name:'roundtrip.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer:Buffer.from(saved.data)});
  await expect(page.locator('.android-paper-viewport')).toBeVisible(); await expect(page.locator('.android-paper-layer')).toContainText('TABLET EDIT');
  await page.locator('.android-topbar').getByRole('button',{name:'편집 모드',exact:true}).click(); await expect(pc).toBeVisible();
  await pc.getByRole('button',{name:'Word 문서 사본 저장 (.docx)',exact:true}).click(); await expect.poll(()=>page.evaluate(()=>(window as any).savedCopies.length)).toBe(2);
  const savedAgain=await page.evaluate(()=>(window as any).savedCopies[1]);const zipAgain=await JSZip.loadAsync(Buffer.from(savedAgain.data));
  const xmlAgain=await zipAgain.file('word/document.xml')!.async('string');
  expect((xmlAgain.match(/<w:drawing[ >]/g)||[])).toHaveLength(1);
  await expect(editor.locator('table')).toHaveCount(1); await expect(editor.locator('img[src]')).toHaveCount(1);
});

test('HWP continuous viewer keeps scale and Escape returns to the original reader', async ({page})=>{
  await page.goto('/');await page.locator('.android-shell > input[type=file]').setInputFiles(path.resolve('tests/fixtures/hwp/landscape-pagination.hwp'));
  const reader=page.getByRole('dialog',{name:'HWP 읽기',exact:true});await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await reader.getByRole('button',{name:'한글 전체 화면 뷰어',exact:true}).click();await expect(reader.locator('.android-presentation-page svg')).toBeVisible();
  await page.keyboard.press('-');await page.keyboard.press('-');const scale=await reader.locator('.android-presentation').getAttribute('data-scale');
  expect(await visiblePages(page,'.android-presentation-slot','.android-presentation-viewport')).toBeGreaterThanOrEqual(2);
  await page.keyboard.press('ArrowRight');await expect(reader.locator('.android-presentation')).toHaveAttribute('data-page','2');
  expect(await reader.locator('.android-presentation').getAttribute('data-scale')).toBe(scale);
  await page.evaluate(()=>window.dispatchEvent(new Event('studio-native-escape')));
  await expect(reader.locator('.android-presentation')).toHaveCount(0);await expect(reader.getByTestId('hwp-page')).toBeVisible();
  await expect(reader.getByRole('region',{name:'한글 편집 변환 안내',exact:true})).toHaveCount(0);
});


test('mixed portrait and landscape pages keep the same physical zoom', async ({page})=>{
  await page.goto('/');
  const buffer=await Packer.toBuffer(new Document({sections:[
    {properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.PORTRAIT}}},children:[new Paragraph('PORTRAIT MARKER')]},
    {properties:{type:SectionType.NEXT_PAGE,page:{size:{width:16838,height:11906,orientation:PageOrientation.LANDSCAPE}}},children:[new Paragraph('LANDSCAPE MARKER')]}
  ]}));
  await page.locator('.android-shell > input[type=file]').setInputFiles({name:'mixed.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer});
  await expect(page.locator('.android-paper-viewport')).toBeVisible();
  await page.locator('.android-topbar').getByRole('button',{name:'전체 화면 뷰어',exact:true}).click();
  await expect(page.locator('.android-presentation-page')).toContainText('PORTRAIT MARKER');
  await page.keyboard.press('+');const scale=Number(await page.locator('.android-presentation').getAttribute('data-scale'));
  await page.keyboard.press('ArrowRight'); await expect(page.locator('.android-presentation-page')).toContainText('LANDSCAPE MARKER');
  expect(Number(await page.locator('.android-presentation').getAttribute('data-scale'))).toBeCloseTo(scale,6);
  const dimensions=await page.locator('.android-presentation-page').evaluate(el=>({logical:(el as HTMLElement).offsetWidth,display:el.getBoundingClientRect().width}));
  expect(dimensions.logical).toBeGreaterThan(1100); expect(dimensions.display/dimensions.logical).toBeCloseTo(scale,2);
});


test('a 768px native tablet keeps PC controls visible and unlocks after native save cancellation', async ({page})=>{
  await page.setViewportSize({width:768,height:1024});
  await page.addInitScript(()=>{
    (window as any).__studioHostConfig={studio:'word',tablet:true};
    (window as any).wordAndroid={saveFile:()=>new Promise(resolve=>{(window as any).finishTabletSave=()=>resolve('cancelled')})};
  });
  await page.goto('/');await page.getByRole('button',{name:'새 문서',exact:true}).click();
  const pc=page.locator('.office-shell[data-android-desktop=true]');await expect(pc).toBeVisible();
  const modes=pc.locator('.android-mode-switch');const r=(await modes.boundingBox())!;expect(r.x).toBeGreaterThanOrEqual(0);expect(r.x+r.width).toBeLessThanOrEqual(768);
  expect(await pc.evaluate(el=>el.scrollWidth<=el.clientWidth+1)).toBe(true);
  const editor=pc.getByLabel('문서 편집 영역');await editor.click();await page.keyboard.type('SAVE CANCEL SAFE');
  await pc.getByRole('button',{name:'Word 문서 사본 저장 (.docx)',exact:true}).click();
  await expect.poll(()=>page.evaluate(()=>typeof (window as any).finishTabletSave)).toBe('function');
  await expect(pc.getByRole('button',{name:'전체 화면 뷰어',exact:true})).toBeDisabled();
  await expect(editor).toHaveAttribute('contenteditable','false');
  await page.evaluate(()=>(window as any).finishTabletSave());
  await expect(pc.getByRole('button',{name:'전체 화면 뷰어',exact:true})).toBeEnabled();await expect(editor).toHaveAttribute('contenteditable','true');await expect(editor).toContainText('SAVE CANCEL SAFE');
});


test('tablet touch can edit a small scaled table cell without starting mouse row resizing',async({page})=>{
  await page.setViewportSize({width:1066,height:1100});await open(page);
  if(await page.getByRole('button',{name:'알림 닫기',exact:true}).count())await page.getByRole('button',{name:'알림 닫기',exact:true}).click();
  await page.locator('.android-topbar').getByRole('button',{name:'편집 모드',exact:true}).click();
  const pc=page.locator('.office-shell[data-android-desktop=true]'),editor=pc.getByLabel('문서 편집 영역');await expect(editor).toContainText('CELL LEFT');
  await pc.getByLabel('확대 축소',{exact:true}).fill('25');
  const cell=editor.locator('td').filter({hasText:'CELL LEFT'});await cell.tap();await page.keyboard.type(' TABLET TOUCH');
  await expect(cell).toContainText('TABLET TOUCH');await expect(editor).not.toHaveClass(/table-row-resizing/);
  await pc.getByRole('button',{name:'실행 취소 (Ctrl+Z)',exact:true}).click();await expect(cell).not.toContainText('TABLET TOUCH');
  await pc.getByRole('button',{name:'다시 실행 (Ctrl+Y)',exact:true}).click();await expect(cell).toContainText('TABLET TOUCH');
});


test('a slow vertical drag turns a fully fitted tablet page and preserves scale',async({page})=>{
  await page.setViewportSize({width:1280,height:900});await open(page);await viewer(page);
  await page.getByRole('button',{name:'쪽 전체',exact:true}).click();
  const scale=Number(await page.locator('.android-presentation').getAttribute('data-scale'));
  const c=await page.context().newCDPSession(page);
  await c.send('Input.dispatchTouchEvent',{type:'touchStart',touchPoints:[{x:640,y:675,id:1}]});
  for(let n=1;n<=3;n++){
    await c.send('Input.dispatchTouchEvent',{type:'touchMove',touchPoints:[{x:640,y:675-150*n,id:1}]});
    await page.waitForTimeout(300);
  }
  await c.send('Input.dispatchTouchEvent',{type:'touchEnd',touchPoints:[]});await c.detach();
  await expect(page.locator('.android-presentation')).toHaveAttribute('data-page','2');
  expect(Number(await page.locator('.android-presentation').getAttribute('data-scale'))).toBeCloseTo(scale,6);
});
