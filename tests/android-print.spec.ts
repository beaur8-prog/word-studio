import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, ImageRun } from 'docx';
import path from 'node:path';
const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=','base64');
async function start(page:Page,tablet=false){
  if(tablet)await page.setViewportSize({width:1280,height:1000});
  await page.addInitScript(()=>{(window as any).printJobs=[];(window as any).wordAndroid={printDocument:(payload:any)=>new Promise((resolve,reject)=>{(window as any).printJobs.push({...payload,html:new TextDecoder().decode(payload.data)});(window as any).finishPrint=resolve;(window as any).failPrint=reject;})};});
  await page.goto('/');
}
async function docx(page:Page){
 const buffer=await Packer.toBuffer(new Document({sections:[{children:[new Paragraph('PRINT FIRST PAGE'),new Table({rows:[new TableRow({children:[new TableCell({columnSpan:2,children:[new Paragraph('PRINT MERGED CELL 123456789')]})]})]}),new Paragraph({children:[new ImageRun({type:'png',data:png,transformation:{width:80,height:40}})]}),new Paragraph({text:'PRINT LAST PAGE',pageBreakBefore:true})]}]}));
 await page.locator('.android-shell > input[type=file]').setInputFiles({name:'print.docx',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',buffer});await expect(page.locator('.android-paper-viewport')).toBeVisible();
 if(await page.getByRole('button',{name:'알림 닫기',exact:true}).count())await page.getByRole('button',{name:'알림 닫기',exact:true}).click();
}
async function jobs(page:Page,count:number){await expect.poll(()=>page.evaluate(()=>(window as any).printJobs.length)).toBe(count);return page.evaluate(()=>(window as any).printJobs.at(-1));}
async function finish(page:Page,result='cancelled'){await page.evaluate(value=>(window as any).finishPrint(value),result);await expect(page.locator('.android-operation')).toHaveCount(0);}

test('phone menu prints full unscaled pages with table/picture and releases the guard after cancellation',async({page})=>{
 await start(page);await docx(page);await page.getByRole('button',{name:'문서 메뉴',exact:true}).click();await page.getByRole('button',{name:'인쇄·PDF 저장',exact:true}).click();
 const payload=await jobs(page,1);expect(payload.pages).toBe(2);expect(payload.pageWidthMils).toBeCloseTo(8268,0);expect(payload.type).toBe('text/html');
 const summary=await page.evaluate(html=>{const d=new DOMParser().parseFromString(html,'text/html');return {text:d.body.textContent,chrome:!!d.body.querySelector('.android-topbar,.office-shell,.android-reader-tools'),pages:d.querySelectorAll('.android-print-sheet').length,table:d.querySelector('td[colspan="2"]')?.textContent,images:d.querySelectorAll('img').length,width:d.querySelector<HTMLElement>('.android-print-source')?.style.width,scripts:d.querySelectorAll('script').length};},payload.html);
 expect(summary).toMatchObject({chrome:false,pages:2,table:'PRINT MERGED CELL 123456789',images:1,scripts:0});expect(summary.width).toMatch(/^793/);expect(summary.text).toContain('PRINT LAST PAGE');
 await expect(page.getByRole('button',{name:'취소',exact:true})).toHaveCount(0);await finish(page);await expect(page.locator('.android-paper-layer')).toContainText('PRINT FIRST PAGE');
 await page.getByRole('button',{name:'문서 메뉴',exact:true}).click();await page.getByRole('button',{name:'인쇄·PDF 저장',exact:true}).click();await jobs(page,2);await finish(page,'closed');
 await expect(page.getByText('인쇄창을 닫았습니다.',{exact:true})).toBeVisible();
});

test('full-screen viewer print and Ctrl+P retain its page and precise zoom',async({page})=>{
 await start(page);await docx(page);await page.locator('.android-topbar').getByRole('button',{name:'전체 화면 뷰어',exact:true}).click();
 await page.keyboard.press('-');await page.keyboard.press('ArrowRight');await expect(page.locator('.android-presentation')).toHaveAttribute('data-page','2');const scale=await page.locator('.android-presentation').getAttribute('data-scale');
 await page.locator('.android-presentation').getByRole('button',{name:'인쇄·PDF 저장',exact:true}).click();await jobs(page,1);await expect(page.locator('.android-operation')).toBeVisible();await finish(page);
 await expect(page.locator('.android-presentation')).toHaveAttribute('data-page','2');expect(await page.locator('.android-presentation').getAttribute('data-scale')).toBe(scale);
 await page.keyboard.press('Control+p');await jobs(page,2);await finish(page,'submitted');await expect(page.locator('.android-presentation-page')).toContainText('PRINT LAST PAGE');
});

test('tablet Ctrl+P and PDF export use current edits, preserve undo and recover after native error',async({page})=>{
 await start(page,true);await docx(page);await page.locator('.android-topbar').getByRole('button',{name:'편집 모드',exact:true}).click();
 const pc=page.locator('.office-shell[data-android-desktop=true]'),editor=pc.getByLabel('문서 편집 영역');await expect(editor).toHaveAttribute('contenteditable','true');await editor.locator('p').first().click();await page.keyboard.press('End');await page.keyboard.type(' PRINT EDIT');
 await page.keyboard.press('Control+p');const payload=await jobs(page,1);expect(payload.html).toContain('PRINT EDIT');await expect(editor).toHaveAttribute('contenteditable','false');await finish(page);await expect(editor).toHaveAttribute('contenteditable','true');await expect(editor).toContainText('PRINT EDIT');
 await pc.getByRole('button',{name:'실행 취소 (Ctrl+Z)',exact:true}).click();await expect(editor).not.toContainText('PRINT EDIT');await pc.getByRole('button',{name:'다시 실행 (Ctrl+Y)',exact:true}).click();await expect(editor).toContainText('PRINT EDIT');
 await pc.getByRole('button',{name:'내보내기',exact:true}).click();await page.getByRole('button',{name:/PDF 문서/}).click();await jobs(page,2);await page.evaluate(()=>(window as any).failPrint(new Error('PRINT FAILURE')));await expect(page.locator('.android-operation')).toHaveCount(0);await expect(editor).toHaveAttribute('contenteditable','true');await expect(editor).toContainText('PRINT EDIT');
 await page.keyboard.press('Control+p');await jobs(page,3);await finish(page,'closed');
});

test('original HWP menu and Ctrl+P print all SVG pages without converting or leaving its reader',async({page})=>{
 await start(page);await page.locator('.android-shell > input[type=file]').setInputFiles(path.resolve('tests/fixtures/hwp/landscape-pagination.hwp'));
 const reader=page.getByRole('dialog',{name:'HWP 읽기',exact:true});await expect(reader.getByTestId('hwp-page')).toBeVisible();await reader.getByRole('button',{name:'한글 문서 메뉴',exact:true}).click();await page.getByRole('button',{name:'인쇄·PDF 저장',exact:true}).click();
 const payload=await jobs(page,1);expect(payload.pages).toBe(4);expect(payload.pageWidthMils).toBeGreaterThan(payload.pageHeightMils);
 const summary=await page.evaluate(html=>{const d=new DOMParser().parseFromString(html,'text/html');return {svg:d.querySelectorAll('.android-print-source > svg').length,text:d.body.textContent};},payload.html);expect(summary.svg).toBe(4);expect(summary.text?.replace(/\s/g,'')).toContain('ROW-24-COLUMN-2');
 await finish(page);await expect(reader.getByTestId('hwp-page')).toBeVisible();await expect(reader.getByRole('region',{name:'한글 편집 변환 안내',exact:true})).toHaveCount(0);
 await page.keyboard.press('Control+p');await jobs(page,2);await finish(page,'closed');await expect(reader).toBeVisible();
});
