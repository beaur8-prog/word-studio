import {test,expect} from '@playwright/test';
import {GROUP_PNG} from './helpers/grouped-drawing';

for(const kind of ['table','floating']as const)for(const direction of ['Right','Left']as const)test(`find then Arrow${direction} collapses an LTR ${kind} text selection before typing`,async({page})=>{
 await page.goto('/');const target=kind==='table'?'가나다 selected cell text':'Selected words across several wrapped lines for a real text caret';
 const doc=await page.evaluate(async({kind,target,png})=>{
  const{defaultPage}=await import('/src/types.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');
  const g=JSON.stringify({horizontalPosition:{relative:'page',offset:276*9525},verticalPosition:{relative:'page',offset:96*9525},wrap:{type:1}});
  const html=kind==='table'?`<table><tbody><tr><td><p>${target}</p></td><td><p>NEIGHBOR CELL</p></td></tr></tbody></table>`:`<p style="font-family:Arial;font-size:12pt;line-height:1;margin:0"><img data-inline-image="true" data-image-floating='${g}' src="data:image/png;base64,${png}" width="160" height="140">${target} ${'Following body text wraps around the same picture. '.repeat(18)}</p>`;
  return{format:'word-studio',version:1,id:'selection-navigation',title:'Selection navigation',updatedAt:new Date().toISOString(),page:defaultPage,comments:[],content:parseDocumentHtml(html)};
 },{kind,target,png:GROUP_PNG});
 await page.locator('input[type=file]').first().setInputFiles({name:'selection-navigation.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});
 const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText(target);if(kind==='floating')await expect(editor.locator('[data-flow-widget]').first()).toBeAttached();
 const before=await editor.evaluate((el:any)=>el.editor.getJSON());await page.keyboard.press('ControlOrMeta+f');const find=page.getByRole('textbox',{name:'문서에서 찾기',exact:true});await find.fill(target);await find.press('Enter');await expect(editor).toBeFocused();
 await expect.poll(()=>editor.evaluate((el:any)=>{const s=el.editor.state.selection;return el.editor.state.doc.textBetween(s.from,s.to,'','\ufffc');})).toBe(target);
 await page.keyboard.press('Arrow'+direction);
 await expect.poll(()=>editor.evaluate((el:any)=>{const s=el.editor.state.selection;return s.empty?el.editor.state.doc.textBetween(s.$from.start(),s.from,'','\ufffc').replace(/\ufffc/g,''):null;})).toBe(direction==='Right'?target:'');
 await page.keyboard.insertText('X');const p=kind==='table'?editor.locator('td p').first():editor.locator('p').first();await expect(p).toContainText(direction==='Right'?target+'X':'X'+target);if(kind==='table')await expect(editor.locator('td').nth(1)).toHaveText('NEIGHBOR CELL');
 await page.keyboard.press('ControlOrMeta+z');expect(await editor.evaluate((el:any)=>el.editor.getJSON())).toEqual(before);await page.keyboard.press('ControlOrMeta+Shift+z');await expect(p).toContainText(direction==='Right'?target+'X':'X'+target);
});
