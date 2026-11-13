import {test,expect} from '@playwright/test';
for(const direction of ['end','forward-range','backward-range'])test(`Tab uses the live native ${direction} before a delayed selectionchange`,async({page})=>{
 await page.goto('/');const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toBeVisible();
 await page.locator('input[type=file]').first().setInputFiles({name:'typing.txt',mimeType:'text/plain',buffer:Buffer.from('Label')});await expect(editor).toHaveText('Label');
 const result=await editor.evaluate((root:any,direction)=>{
  const ed=root.editor;ed.commands.setTextSelection(1);root.focus();
  const text=root.querySelector('p').firstChild,selection=document.getSelection();
  const [anchor,head]=direction==='end'?[5,5]:direction==='forward-range'?[1,4]:[4,1];
  selection.setBaseAndExtent(text,anchor,text,head);
  const before={model:ed.state.selection.toJSON(),native:[selection.anchorOffset,selection.focusOffset]};
  const event=new KeyboardEvent('keydown',{key:'Tab',code:'Tab',keyCode:9,bubbles:true,cancelable:true});root.dispatchEvent(event);
  return{before,handled:event.defaultPrevented,model:ed.getJSON()};
 },direction);
 expect(result.before.model).toEqual({type:'text',anchor:1,head:1});expect(result.handled).toBe(true);
 await page.keyboard.insertText('Value');await expect(editor).toBeFocused();await expect(editor.locator('[data-docx-tab]')).toHaveCount(1);
 await expect(editor).toHaveText(direction==='end'?'Label\tValue':'L\tValuel');
 await page.keyboard.press('ControlOrMeta+z');await expect(editor).toHaveText('Label');await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor).toHaveText(direction==='end'?'Label\tValue':'L\tValuel');
});

test('a live native caret inside a table still delegates Tab to the next cell',async({page})=>{
 await page.goto('/');const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toBeVisible();
 const result=await editor.evaluate((root:any)=>{
  const ed=root.editor;ed.commands.setContent('<p>BODY</p><table><tr><td><p>ONE</p></td><td><p>TWO</p></td></tr></table>');ed.commands.setTextSelection(1);root.focus();
  const text=root.querySelector('td p').firstChild;document.getSelection().setBaseAndExtent(text,3,text,3);
  root.dispatchEvent(new KeyboardEvent('keydown',{key:'Tab',code:'Tab',keyCode:9,bubbles:true,cancelable:true}));
  return{selection:ed.state.selection.$from.parent.textContent,tabs:root.querySelectorAll('[data-docx-tab]').length,text:ed.state.doc.textContent};
 });
 expect(result).toEqual({selection:'TWO',tabs:0,text:'BODYONETWO'});
});

test('Tab from a trailing table gap retains the gap instead of reading its old cell DOM selection',async({page})=>{
 await page.goto('/');const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toBeVisible();
 const result=await editor.evaluate(async(root:any)=>{
  const {GapCursor}=await import('/node_modules/@tiptap/pm/gapcursor/index.ts');const ed=root.editor;
  ed.commands.setContent('<table><tr><td><p>CELL</p></td></tr></table>');root.focus();ed.view.dispatch(ed.state.tr.setSelection(new GapCursor(ed.state.doc.resolve(ed.state.doc.content.size))));
  const before=ed.state.selection.toJSON();root.dispatchEvent(new KeyboardEvent('keydown',{key:'Tab',code:'Tab',keyCode:9,bubbles:true,cancelable:true}));
  return{before,cell:root.querySelector('td').textContent,body:ed.getJSON()};
 });
 expect(result.before.type).toBe('gapcursor');expect(result.cell).toBe('CELL');
 expect(result.body.content.at(-1).type).toBe('paragraph');expect(result.body.content.at(-1).content).toEqual([{type:'documentTab'}]);
});
