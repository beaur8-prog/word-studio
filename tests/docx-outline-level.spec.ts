import { test, expect, type Page } from '@playwright/test';
import { readFile, writeFile } from 'node:fs/promises';

// Authored OOXML controls, separate from the original Microsoft Word corpus.
async function input(page: Page, headingReset = false) {
  await page.goto('/');
  return page.evaluate(async headingReset => {
    const { exportDocx } = await import('/src/document-io.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    const texts = ['ZERO', 'TWO', 'BODY', 'INHERITED', 'RESET', 'PLAIN'];
    const blob = await exportDocx({ format: 'word-studio', version: 1, title: 'OUTLINE', updatedAt: new Date().toISOString(),
      page: { ...defaultPage }, comments: [], content: { type: 'doc', content: texts.map(text => ({ type: 'paragraph', content: [{ type: 'text', text }] })) } });
    const zip = await JSZip.loadAsync(await blob.arrayBuffer()), W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const doc = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const paragraphs = [...doc.getElementsByTagNameNS(W, 'body')[0].children].filter(n => n.localName === 'p');
    for (const [index, level] of [0, 2, 9].entries()) {
      const props = paragraphs[index].getElementsByTagNameNS(W, 'pPr')[0];
      const outline = doc.createElementNS(W, 'w:outlineLvl'); outline.setAttributeNS(W, 'w:val', String(level)); props.insertBefore(outline, [...props.children].find(n => ['rPr', 'sectPr', 'pPrChange'].includes(n.localName)) ?? null);
    }
    if (headingReset) {
      const props = paragraphs[2].getElementsByTagNameNS(W, 'pPr')[0];
      props.getElementsByTagNameNS(W, 'pStyle')[0]?.remove();
      const style = doc.createElementNS(W, 'w:pStyle'); style.setAttributeNS(W, 'w:val', 'Heading1'); props.insertBefore(style, props.firstChild);
    }
    const styles = new DOMParser().parseFromString(await zip.file('word/styles.xml')!.async('string'), 'application/xml');
    for (const [id, parent, level] of [['OutlineBody', 'Normal', '2'], ['OutlineReset', 'OutlineBody', '9']]) {
      const style = styles.createElementNS(W, 'w:style'); style.setAttributeNS(W, 'w:type', 'paragraph'); style.setAttributeNS(W, 'w:styleId', id);
      const name = styles.createElementNS(W, 'w:name'); name.setAttributeNS(W, 'w:val', id); style.append(name);
      const based = styles.createElementNS(W, 'w:basedOn'); based.setAttributeNS(W, 'w:val', parent); style.append(based);
      const props = styles.createElementNS(W, 'w:pPr'), outline = styles.createElementNS(W, 'w:outlineLvl'); outline.setAttributeNS(W, 'w:val', level); props.append(outline); style.append(props); styles.documentElement.append(style);
    }
    for (const [index, id] of [[3, 'OutlineBody'], [4, 'OutlineReset']] as const) {
      const props = paragraphs[index].getElementsByTagNameNS(W, 'pPr')[0];
      const old = props.getElementsByTagNameNS(W, 'pStyle')[0]; if (old) old.remove();
      const style = doc.createElementNS(W, 'w:pStyle'); style.setAttributeNS(W, 'w:val', id); props.insertBefore(style, props.firstChild);
    }
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc)); zip.file('word/styles.xml', new XMLSerializer().serializeToString(styles));
    return [...await zip.generateAsync({ type: 'uint8array' })];
  }, headingReset);
}

async function savedLevels(page: Page, bytes: number[]) {
  return page.evaluate(async bytes => {
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js'); const zip = await JSZip.loadAsync(new Uint8Array(bytes));
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const doc = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'), 'application/xml');
    const styles = new DOMParser().parseFromString(await zip.file('word/styles.xml')!.async('string'), 'application/xml');
    const byId = new Map([...styles.getElementsByTagNameNS(W, 'style')].map(s => [s.getAttributeNS(W, 'styleId'), s]));
    const styleLevel = (id: string, seen = new Set<string>()): number | null => {
      if (seen.has(id)) throw new Error('Cyclic output style'); seen.add(id);
      const style = byId.get(id), props = style && [...style.children].find(n => n.localName === 'pPr');
      const level = props && [...props.children].find(n => n.localName === 'outlineLvl');
      if (level) return Number(level.getAttributeNS(W, 'val'));
      const parent = style && [...style.children].find(n => n.localName === 'basedOn')?.getAttributeNS(W, 'val');
      return parent ? styleLevel(parent, seen) : null;
    };
    return [...doc.getElementsByTagNameNS(W, 'body')[0].children].filter(n => n.localName === 'p').map(p => {
      const props = [...p.children].find(n => n.localName === 'pPr'), level = props && [...props.children].find(n => n.localName === 'outlineLvl');
      const style = props && [...props.children].find(n => n.localName === 'pStyle')?.getAttributeNS(W, 'val');
      return level ? Number(level.getAttributeNS(W, 'val')) : styleLevel(style || 'Normal');
    });
  }, bytes);
}

test('direct and inherited outline levels survive two saves without creating heading keeps', async ({ page }, info) => {
  const bytes = await input(page, true);
  const result = await page.evaluate(async bytes => {
    const { importFile, exportDocx } = await import('/src/document-io.ts'); const { parseDocumentHtml } = await import('/src/extensions.ts');
    let file = new File([new Uint8Array(bytes)], 'outline.docx'); const states = [], saves = [];
    for (let i = 0; i < 3; i++) {
      const loaded = await importFile(file), doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) };
      states.push(doc.content.content!.map(n => ({ type: n.type, level: n.attrs?.wordOutlineLevel ?? null, keep: n.attrs?.wordKeepNext === true, text: n.content?.map(t => t.text ?? '').join('') })));
      if (i < 2) { const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]); file = new File([blob], 'again.docx'); }
    }
    return { states, saves };
  }, bytes);
  for (const [i, file] of [bytes, ...result.saves].entries()) await writeFile(info.outputPath(`textbox-save-outline-${i}.docx`), Buffer.from(file));
  for (const state of result.states) {
    expect(state.map(n => n.level)).toEqual([0, 2, 9, 2, 9, null]);
    expect(state.map(n => n.type)).toEqual(['paragraph', 'paragraph', 'heading', 'paragraph', 'paragraph', 'paragraph']);
    expect(state.every(n => !n.keep)).toBe(true);
    expect(state.map(n => n.text)).toEqual(['ZERO', 'TWO', 'BODY', 'INHERITED', 'RESET', 'PLAIN']);
  }
  for (const file of result.saves) expect(await savedLevels(page, file)).toEqual([0, 2, 9, 2, 9, null]);
});

test('typing, undo and real Word exports retain outline metadata and the ordinary paragraph view', async ({ page }, info) => {
  const bytes = await input(page);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'outline.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(bytes) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }); await expect(editor).toContainText('INHERITED');
  await expect(editor.locator('p')).toHaveCount(6); await expect(editor.locator('h1,h2,h3,h4,h5,h6')).toHaveCount(0);
  await editor.locator('p').first().click(); await page.keyboard.press('End'); await page.keyboard.insertText(' EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(editor.locator('p').first()).toHaveText('ZERO');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(editor.locator('p').first()).toHaveText('ZERO EDIT');
  for (let save = 1; save <= 2; save++) {
    await page.getByRole('button', { name: '파일', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
    const download = await pending, file = [...await readFile((await download.path())!)];
    await writeFile(info.outputPath(`textbox-save-outline-ui-${save}.docx`), Buffer.from(file));
    expect(await savedLevels(page, file)).toEqual([0, 2, 9, 2, 9, null]);
    await page.locator('input[type=file]').first().setInputFiles({ name: `outline-save-${save}.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(file) });
    const title = page.getByRole('textbox', { name: '문서 제목', exact: true });
    await expect.poll(async () => await title.inputValue() === `outline-save-${save}` || await page.getByRole('dialog', { name: '변경 내용 보관' }).isVisible()).toBe(true);
    if (await page.getByRole('dialog', { name: '변경 내용 보관' }).isVisible()) await page.getByRole('button', { name: '복구본 보관 후 계속' }).click();
    await expect(title).toHaveValue(`outline-save-${save}`);
    await expect(editor.locator('p').first()).toHaveText('ZERO EDIT');
  }
  await editor.locator('p').first().click();
  await page.getByRole('button', { name: '가나다Aa 제목 2', exact: true }).click();
  await expect(editor.locator('h2')).toHaveText('ZERO EDIT');
  await expect(editor.locator('h2')).not.toHaveAttribute('data-word-outline-level', '0');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(editor.locator('p').first()).toHaveAttribute('data-word-outline-level', '0');
  await page.keyboard.press('ControlOrMeta+Shift+z');
  await expect(editor.locator('h2')).toHaveText('ZERO EDIT');
  await page.getByRole('button', { name: '파일', exact: true }).click(); const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
  const download = await pending, file = [...await readFile((await download.path())!)];
  await writeFile(info.outputPath('textbox-save-outline-heading-edit.docx'), Buffer.from(file));
  expect(await savedLevels(page, file)).toEqual([1, 2, 9, 2, 9, null]);
});

test('outline metadata accepts only integer levels zero through nine', async ({ page }) => {
  await page.goto('/');
  const rejected = await page.evaluate(async () => {
    const { validateDocument } = await import('/src/document-io.ts'); const { defaultPage } = await import('/src/types.ts');
    return [-1, 10, 1.5, '0', {}, true].map(wordOutlineLevel => {
      try { validateDocument({ format: 'word-studio', version: 1, title: 'BOUND', updatedAt: new Date().toISOString(), page: defaultPage, comments: [], content: { type: 'doc', content: [{ type: 'paragraph', attrs: { wordOutlineLevel } }] } }); return false; } catch { return true; }
    });
  });
  expect(rejected).toEqual([true, true, true, true, true, true]);
});
