import { test, expect } from '@playwright/test';
import JSZip from 'jszip';

// Small synthesized regression fixture; the separately frozen publisher DOCX
// files supply the actual Word-authored layout evidence.
async function headingFixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="240" w:lineRule="auto" w:before="0" w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/><w:spacing w:before="120" w:after="40"/></w:pPr><w:rPr><w:sz w:val="26"/></w:rPr></w:style>
    <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/><w:spacing w:before="60" w:after="40"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
  </w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:p><w:r><w:t>Opening paragraph</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Regular heading</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t xml:space="preserve">Inherited bold </w:t></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>Explicit normal</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>
  </w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('imported headings use the source spacing and weight in editor and print', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'headings.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await headingFixture() });
  const heading = page.locator('.tiptap h2').filter({ hasText: 'Regular heading' });
  await expect(heading).toHaveCSS('padding-top', '0px');
  await expect(heading).toHaveCSS('margin-top', '8px');
  await expect(heading).toHaveCSS('margin-bottom', '2.66667px');
  await expect(heading).toHaveCSS('font-weight', '400');
  await page.getByRole('button', { name: '보기', exact: true }).click();
  await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
  const printed = page.locator('.print-content h2').filter({ hasText: 'Regular heading' });
  await expect(printed).toHaveCSS('padding-top', '0px');
  await expect(printed).toHaveCSS('font-weight', '400');
  await expect(page.locator('.print-content')).toContainText('Inherited bold Explicit normal');
});

test('two DOCX saves preserve regular headings and direct bold overrides', async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    let file = new File([new Uint8Array(bytes)], 'headings.docx');
    const states = [];
    for (let stage = 0; stage < 3; stage++) {
      const opened = await importFile(file), content = parseDocumentHtml(opened.html!);
      const headings = content.content!.filter(n => n.type === 'heading').map(n => ({
        level: n.attrs!.level,
        runs: (n.content || []).filter(n => n.type === 'text').map(n => ({ text: n.text, bold: !!n.marks?.some(m => m.type === 'bold') })),
      }));
      states.push({ headings, html: opened.html });
      if (stage < 2) file = new File([await exportDocx({ ...opened.document, content })], 'again.docx');
    }
    return states;
  }, [...await headingFixture()]);
  await test.info().attach("roundtrip-states", { body: Buffer.from(JSON.stringify(states, null, 2)), contentType: "application/json" });
  for (const [stage, state] of states.entries()) expect(state.headings, `save stage ${stage}`).toEqual([
    { level: 2, runs: [{ text: 'Regular heading', bold: false }] },
    { level: 3, runs: [{ text: 'Inherited bold ', bold: true }, { text: 'Explicit normal', bold: false }] },
  ]);
});

test('ordinary authored headings retain Studio heading defaults', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'authored.html', mimeType: 'text/html', buffer: Buffer.from('<p>Opening</p><h2>Authored heading</h2>') });
  const heading = page.locator('.tiptap h2').filter({ hasText: 'Authored heading' });
  await expect(heading).not.toHaveAttribute('data-word-line-rule');
  await expect(heading).toHaveCSS('font-weight', '600');
  await expect(heading).toHaveCSS('padding-top', '12px');
});

// A saved imported heading must not pick up docx's authored Heading2 palette
// as its default for newly typed text while existing black runs hide the drift.
test('two DOCX saves keep imported heading defaults for retyping', async ({ page }) => {
  await page.goto('/');
  const inputs = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const files = []; let file = new File([new Uint8Array(bytes)], 'headings.docx');
    for (let stage = 0; stage < 3; stage++) {
      files.push([...new Uint8Array(await file.arrayBuffer())]);
      const imported = await importFile(file), content = parseDocumentHtml(imported.html!);
      if (stage < 2) file = new File([await exportDocx({ ...imported.document, content })], 'headings.docx');
    }
    return files;
  }, [...await headingFixture()]);
  for (const [stage, bytes] of inputs.entries()) {
    await page.locator('input[type=file]').first().setInputFiles({ name: `heading-color-${stage}.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) });
    const editor = page.locator('.tiptap');
    await expect(editor.locator('h2').filter({ hasText: 'Regular heading' })).toBeVisible();
    const heading = editor.locator('h2').first();
    const box = await heading.boundingBox();
    await heading.click({ position: { x: 1, y: box!.height / 2 } });
    await page.keyboard.down('Shift');
    for (let i = 0; i < 'Regular heading'.length; i++) await page.keyboard.press('ArrowRight');
    await page.keyboard.up('Shift');
    await expect.poll(() => page.evaluate(() => getSelection()?.toString())).toBe('Regular heading');
    await expect.poll(() => editor.evaluate((el: HTMLElement & { editor: any }) => {
      const selection = getSelection()!;
      const from = el.editor.view.posAtDOM(selection.anchorNode, selection.anchorOffset);
      const to = el.editor.view.posAtDOM(selection.focusNode, selection.focusOffset);
      return el.editor.state.selection.from === Math.min(from, to) && el.editor.state.selection.to === Math.max(from, to);
    })).toBe(true);
    await page.keyboard.press('Backspace');
    await expect(heading).toBeEmpty();
    // Clear the previous run's stored marks by visiting another paragraph.
    await editor.locator('p').first().click();
    await heading.click();
    await expect.poll(() => editor.evaluate((el: HTMLElement & { editor: any }) => {
      const selection = getSelection()!;
      return el.editor.state.selection.from === el.editor.view.posAtDOM(selection.anchorNode, selection.anchorOffset);
    })).toBe(true);
    await page.keyboard.insertText('Retyped heading');
    await expect(heading).toHaveText('Retyped heading');
    const color = await heading.evaluate(el => {
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT); walker.nextNode();
      return getComputedStyle(walker.currentNode.parentElement!).color;
    });
    expect(color, `actual newly typed ink after DOCX save ${stage}`).toBe('rgb(0, 0, 0)');
  }
});

test('paragraph-mark color does not replace real run colors or the nonempty font strut', async ({ page }) => {
  const zip = await JSZip.loadAsync(await headingFixture());
  let source = await zip.file('word/document.xml')!.async('string');
  source = source.replace('<w:pStyle w:val="Heading2"/>', '<w:pStyle w:val="Heading2"/><w:rPr><w:rFonts w:ascii="Georgia"/><w:sz w:val="144"/><w:color w:val="008000"/></w:rPr>');
  source = source.replace('<w:t>Regular heading</w:t>', '<w:rPr><w:color w:val="C00000"/></w:rPr><w:t>Regular heading</w:t>');
  zip.file('word/document.xml', source);
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'heading-mark-color.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.tiptap h2').filter({ hasText: 'Regular heading' })).toBeVisible();
  const heading = page.locator('.tiptap h2').first();
  await expect(heading).toHaveCSS('color', 'rgb(0, 128, 0)');
  await expect(heading).toHaveCSS('font-size', '17.3333px');
  await expect(heading).toHaveCSS('font-family', 'Arial');
  const actualRunColor = await heading.evaluate(el => {
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT); walker.nextNode();
    return getComputedStyle(walker.currentNode.parentElement!).color;
  });
  expect(actualRunColor).toBe('rgb(192, 0, 0)');
});
