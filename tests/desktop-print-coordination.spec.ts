import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

test('desktop print completes its cold snapshot while editor layout requests are pending, then resumes editing', async ({page})=>{
  await page.goto('/');
  const doc=await page.evaluate(async()=>{const{defaultPage}=await import('/src/types.ts'),{parseDocumentHtml}=await import('/src/extensions.ts');return{format:'word-studio',version:1,id:'print-coordination',title:'Print coordination',updatedAt:new Date().toISOString(),comments:[],page:{...defaultPage,size:'Custom',width:140,height:150,margin:12,pageNumbers:false},content:parseDocumentHtml(Array.from({length:25},(_,i)=>`<p style="font-family:'WS Desktop Cold';font-size:11pt;line-height:normal;margin:0">MARKER${String(i).padStart(2,'0')} A cold snapshot must retain every paragraph while the editing surface waits.</p>`).join(''))};});
  await page.locator('input[type=file]').first().setInputFiles({name:'print.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(doc))});
  const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true});await expect(editor).toContainText('MARKER24');await expect(page.locator('.busy-indicator')).toBeHidden();
  let release!:()=>void;const gate=new Promise<void>(resolve=>release=resolve);
  await page.route('**/texgyreadventor-regular.otf?desktop-cold',async route=>{await gate;await route.continue().catch(()=>{});});
  await page.evaluate(()=>{const face=new FontFace('WS Desktop Cold','url("/src/assets/fonts/texgyreadventor-regular.otf?desktop-cold")');document.fonts.add(face);});
  try{
    await page.getByRole('button',{name:'보기',exact:true}).click();await page.getByRole('button',{name:'인쇄 미리보기',exact:true}).click();
    await expect(page.locator('.print-layout.measuring')).toHaveCount(1);
    const reads=await page.evaluate(async()=>{
      const root=document.querySelector('.tiptap[contenteditable=true]')!,original=Element.prototype.getBoundingClientRect;let reads=0;
      Element.prototype.getBoundingClientRect=function(){if(this===root||root.contains(this))reads++;return original.call(this);};
      try{for(let i=0;i<8;i++){document.dispatchEvent(new Event('studio-editor-layout-resume'));await new Promise<void>(r=>requestAnimationFrame(()=>requestAnimationFrame(()=>r())));}return reads;}
      finally{Element.prototype.getBoundingClientRect=original;}
    });
    await writeFile(test.info().outputPath('print-editor-layout-reads.json'),JSON.stringify({reads},null,2));
    expect(reads).toBeLessThan(5);
    release();await expect(page.getByRole('toolbar',{name:'인쇄 미리보기',exact:true})).toBeVisible();
    expect(await page.locator('.print-page').count()).toBeGreaterThan(1);
    const text=await page.locator('.print-content').allTextContents();for(let i=0;i<25;i++)expect(text.join('')).toContain('MARKER'+String(i).padStart(2,'0'));
    await page.getByRole('button',{name:'편집으로 돌아가기',exact:true}).click();
    const first=editor.locator('p').first(),before=await editor.evaluate((root:any)=>root.editor.getJSON());await first.fill('EDIT AFTER PRINT');await expect(first).toHaveText('EDIT AFTER PRINT');await page.keyboard.press('ControlOrMeta+z');await expect.poll(()=>editor.evaluate((root:any)=>root.editor.getJSON())).toEqual(before);
  }finally{release();}
});
