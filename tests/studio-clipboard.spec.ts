import {test,expect} from '@playwright/test';
import fs from 'node:fs';
const rich=JSON.parse(fs.readFileSync('tests/fixtures/studio-table.json','utf8'));
test('Home paste retains an editable Sheet Studio table, merges, text and borders',async({page})=>{
 await page.goto('/');await page.locator('input[type=file]').first().setInputFiles({name:'clipboard-empty.html',mimeType:'text/html',buffer:Buffer.from('<p>붙여넣기 위치</p>')});await expect(page.locator('.tiptap')).toHaveText('붙여넣기 위치');await page.locator('.tiptap p').click();await page.keyboard.press('End');
 await page.evaluate(rich=>{Object.defineProperty(navigator.clipboard,'read',{configurable:true,value:async()=>[{types:['text/html','text/plain'],getType:async(type:string)=>new Blob([type==='text/html'?rich.html:rich.text],{type})}]});},rich);
 await page.getByRole('button',{name:'붙여넣기',exact:true}).first().click();
 await expect(page.locator('.tiptap table')).toHaveCount(1);const cells=page.locator('.tiptap td');await expect(cells).toHaveCount(3);await expect(cells.first()).toHaveAttribute('colspan','2');await expect(cells.first()).toHaveCSS('background-color','rgb(255, 242, 204)');await expect(cells.first()).toHaveCSS('border-bottom-color','rgb(255, 0, 0)');await expect(cells.first()).toContainText('보고서');await expect(cells.nth(1)).toContainText('줄1');await expect(cells.nth(1)).toContainText('줄2');await expect(cells.nth(2)).toHaveText('=literal');
 await page.keyboard.press('ControlOrMeta+z');await expect(page.locator('.tiptap table')).toHaveCount(0);await page.keyboard.press('ControlOrMeta+Shift+z');await expect(page.locator('.tiptap table')).toHaveCount(1);
 // A normal keyboard paste goes through the same sanitizer and table schema.
 await page.locator('.tiptap').click();await page.keyboard.press('ControlOrMeta+a');await page.evaluate(rich=>{const data=new DataTransfer();data.setData('text/html',rich.html);data.setData('text/plain',rich.text);document.querySelector('.tiptap')!.dispatchEvent(new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data}));},rich);await expect(page.locator('.tiptap table')).toHaveCount(1);
 const documentHtml=await page.locator('.tiptap').innerHTML();expect(documentHtml).toContain('colspan="2"');await expect(page.locator('.tiptap td').first()).toHaveCSS('border-bottom-width','3px');expect(await page.locator('.tiptap table').first().evaluate(el=>el.getBoundingClientRect().width)).toBeCloseTo(250,0);
 await page.screenshot({path:'test-results/studio-table-paste.png'});
});
