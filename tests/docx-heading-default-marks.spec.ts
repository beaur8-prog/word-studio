import { test, expect } from '@playwright/test';
import JSZip from 'jszip';

// Synthetic source-style probes, not additions to the frozen Word originals.
async function fixture() {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  z.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:strike/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b w:val="0"/><w:i w:val="0"/><w:u w:val="none"/><w:strike w:val="0"/></w:rPr></w:style></w:styles>`);
  const off = '<w:b w:val="0"/><w:i w:val="0"/><w:u w:val="none"/><w:strike w:val="0"/>';
  z.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Body paragraph</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Styled heading</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">Inherited </w:t></w:r><w:r><w:rPr>${off}</w:rPr><w:t>Plain override</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/><w:rPr>${off}</w:rPr></w:pPr><w:r><w:t>Mark override</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Regular heading</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: 'nodebuffer' });
}
async function files(page: any) {
  return page.evaluate(async (bytes: number[]) => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const output = []; let file = new File([new Uint8Array(bytes)], 'heading-defaults.docx');
    for (let stage = 0; stage < 3; stage++) {
      const imported = await importFile(file), content = parseDocumentHtml(imported.html!);
      output.push({ bytes: [...new Uint8Array(await file.arrayBuffer())], headings: content.content!.filter(n => n.type === 'heading') });
      if (stage < 2) file = new File([await exportDocx({ ...imported.document, content })], 'heading-defaults.docx');
    }
    return output;
  }, [...await fixture()]);
}
const enabled = { bold: true, italic: true, underline: true, strike: true };
const disabled = { bold: false, italic: false, underline: false, strike: false };

test('heading typing marks and explicit normal bold/italic runs survive two DOCX saves', async ({ page }) => {
  await page.goto('/');
  for (const [stage, data] of (await files(page)).entries()) {
    expect(data.headings.map((n: any) => JSON.parse(n.attrs.wordHeadingMarks || 'null')), `default marks stage${stage}`).toEqual([enabled, enabled, disabled, disabled]);
    const mixed = data.headings[1];
    expect(mixed.content![0].marks!.map((m: any) => m.type)).toEqual(expect.arrayContaining(['bold', 'italic', 'underline', 'strike']));
    expect(mixed.content![1].marks?.some((m: any) => ['bold','italic'].includes(m.type)) || false).toBe(false);
  }
});

test('empty imported headings restore their defaults for new typing and respect a user toggle', async ({ page }) => {
  await page.goto('/');
  for (const [stage, data] of (await files(page)).entries()) {
    await page.locator('input[type=file]').first().setInputFiles({ name: `defaults-${stage}.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(data.bytes) });
    const editor = page.locator('.tiptap'); await expect(editor).toContainText('Styled heading');
    const sync = async () => expect.poll(() => editor.evaluate((el: HTMLElement & {editor:any}) => {
      const s=getSelection()!,a=el.editor.view.posAtDOM(s.anchorNode,s.anchorOffset),b=el.editor.view.posAtDOM(s.focusNode,s.focusOffset);
      return el.editor.state.selection.from===Math.min(a,b)&&el.editor.state.selection.to===Math.max(a,b);
    })).toBe(true);
    for (const [index, text, styled] of [[0,'Styled heading',true],[2,'Mark override',false],[3,'Regular heading',false]] as const) {
      const heading = editor.locator('h2,h3').nth(index); const box = await heading.boundingBox();
      await heading.click({ position: { x: 1, y: box!.height / 2 } }); await page.keyboard.down('Shift');
      for (let i=0;i<text.length;i++) await page.keyboard.press('ArrowRight'); await page.keyboard.up('Shift');
      await expect.poll(() => page.evaluate(() => getSelection()?.toString())).toBe(text);
      await sync(); await page.keyboard.press('Backspace'); await expect(heading).toBeEmpty();
      await editor.locator('p').first().click(); await sync(); await heading.click(); await sync();
      if (styled) await expect(page.getByRole('button', {name:'굵게 (Ctrl+B)',exact:true})).toHaveAttribute('aria-pressed','true');
      await page.keyboard.insertText('Retyped');
      const marks = await heading.evaluate(el => { const w=document.createTreeWalker(el,NodeFilter.SHOW_TEXT);w.nextNode();const s=getComputedStyle(w.currentNode.parentElement!);return{weight:s.fontWeight,italic:s.fontStyle,decoration:s.textDecorationLine}; });
      expect(marks.weight, `stage${stage} heading${index}`).toBe(styled?'700':'400');
      expect(marks.italic).toBe(styled?'italic':'normal');
      // Decorations may be painted by an ancestor; the model checks above
      // verify source defaults and bold/italic resets. Existing negative underline
      // and strike inheritance is a separate recorded import limitation.
      if (styled) {
        // A quick replacement is one undo group: undo restores the original
        // title, and redo restores the replacement rather than an empty node.
        await page.keyboard.press('ControlOrMeta+z'); await expect(heading).toHaveText(text);
        await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(heading).toHaveText('Retyped');
        const rect=await heading.boundingBox(); await heading.click({position:{x:1,y:rect!.height/2}});
        await page.keyboard.down('Shift'); for(let i=0;i<'Retyped'.length;i++)await page.keyboard.press('ArrowRight'); await page.keyboard.up('Shift');
        await expect.poll(()=>page.evaluate(()=>getSelection()?.toString())).toBe('Retyped'); await sync();
        await page.keyboard.press('Backspace'); await expect(heading).toBeEmpty();
        await editor.locator('p').first().click(); await sync(); await heading.click(); await sync();
        await page.keyboard.press('ControlOrMeta+b');
        await editor.locator('p').first().click(); await sync(); await heading.click(); await sync();
        await page.keyboard.insertText('User normal');
        const weight=await heading.evaluate(el=>{const w=document.createTreeWalker(el,NodeFilter.SHOW_TEXT);w.nextNode();return getComputedStyle(w.currentNode.parentElement!).fontWeight;});
        expect(weight).toBe('400');
        await page.keyboard.press('ControlOrMeta+z'); await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(heading).toHaveText('User normal');
      }
    }
  }
});
