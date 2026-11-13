import { test, expect, type Page } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

const editor = (page: Page) => page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
async function geometry(page: Page) {
  return editor(page).evaluate(async root => {
    const { documentScale } = await import('/src/document-scale.ts');
    const scale = documentScale(root), image = root.querySelector<HTMLImageElement>('img[alt="CELL ART"]')!, cell = image.closest('td,th')!;
    const blocks = [...cell.children].filter(el => el.matches('p,h1,h2,h3,h4,h5,h6'));
    const rect = (el: Element) => { const r = el.getBoundingClientRect(); return { top: r.top / scale, bottom: r.bottom / scale, height: r.height / scale, width: r.width / scale }; };
    return { image: rect(image), cell: rect(cell), first: rect(blocks[0]), last: rect(blocks.at(-1)!), blocks: blocks.length,
      pages: Number(root.closest<HTMLElement>('.document-stack')!.dataset.pages),
      text: cell.textContent, geometry: JSON.parse(image.dataset.imageFloating!), src: image.src };
  });
}

for (const align of ['top', 'middle', 'bottom']) for (const multiple of [false, true]) test(`cell-relative artwork includes its extent in a ${align} aligned ${multiple ? 'multi-paragraph' : 'single-paragraph'} cell, through editing and two DOCX saves`, async ({ page }) => {
  test.setTimeout(120000);
  const errors: string[] = []; page.on('pageerror', error => errors.push(String(error)));
  await page.goto('/'); await expect(editor(page)).toBeVisible();
  const result = await page.evaluate(async ({ align, multiple }) => {
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
    const { defaultPage } = await import('/src/types.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { withoutEditorReflow } = await import('/src/editor-layout-task.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const floating = { horizontalPosition: { relative: 'column', offset: 0 }, verticalPosition: { relative: 'page', offset: 70 * 9525 }, wrap: { type: 1, side: 'bothSides' }, layoutInCell: true, behindDocument: false, allowOverlap: true, margins: { top: 0, right: 0, bottom: 0, left: 0 } };
    const p = (text: string) => `<p style="font-family:Arial;font-size:11pt;line-height:20px;margin:0">${text}</p>`;
    const image = `<img data-inline-image="true" alt="CELL ART" src="data:image/png;base64,${png}" width="140" height="80" data-image-floating="${JSON.stringify(floating).replaceAll('"', '&quot;')}">`;
    let doc: any = { format: 'word-studio', version: 1, id: 'cell-relative', title: 'Cell artwork', comments: [], updatedAt: new Date().toISOString(), page: { ...defaultPage, pageNumbers: false },
      content: parseDocumentHtml(p('BEFORE TABLE') + `<table style="width:500px"><colgroup><col style="width:250px"><col style="width:250px"></colgroup><tr style="height:220px" data-word-row-rule="atLeast"><td colwidth="250" style="width:250px;padding:8px;vertical-align:${align}">${multiple ? p('FIRST BLOCK') : ''}${p(image)}${multiple ? p('LAST BLOCK') : ''}</td><td colwidth="250" style="width:250px;padding:8px;vertical-align:top">${p('NEIGHBOR CELL')}</td></tr></table>` + p('AFTER TABLE')) };
    const states = [], saves = [];
    for (let pass = 0; pass < 3; pass++) {
      const printed = await withoutEditorReflow(async () => {
        const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
        try {
          const image = layout.element.querySelector<HTMLImageElement>('img[alt="CELL ART"]')!, cell = image.closest('td,th')!, ps = [...cell.children].filter(el => el.matches('p,h1,h2,h3,h4,h5,h6'));
          const r = image.getBoundingClientRect(), c = cell.getBoundingClientRect(), first = ps[0].getBoundingClientRect();
          return { pages: layout.pages, offset: r.top - first.top, bottomInside: c.bottom - r.bottom, width: r.width, height: r.height, cellHeight: c.height, text: cell.textContent, geometry: JSON.parse(image.dataset.imageFloating!), src: image.src };
        } finally { layout.dispose(); }
      });
      states.push({ doc, printed });
      if (pass < 2) { const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]); const opened = await importFile(new File([blob], 'cell-art.docx')); doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; }
    }
    return { states, saves };
  }, { align, multiple });
  for (const [index, bytes] of result.saves.entries()) await writeFile(test.info().outputPath(`textbox-save-cell-relative-${align}-${multiple}-${index + 1}.docx`), Buffer.from(bytes));
  for (const state of result.states) {
    expect(state.printed.pages).toBe(1); expect(state.printed.offset).toBeCloseTo(70, 0); expect(state.printed.bottomInside).toBeGreaterThanOrEqual(7);
    expect(state.printed.width).toBeCloseTo(140, 0); expect(state.printed.height).toBeCloseTo(80, 0);
    expect(state.printed.geometry).toEqual(result.states[0].printed.geometry); expect(state.printed.src).toEqual(result.states[0].printed.src);
    expect(state.printed.text).toEqual(result.states[0].printed.text);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'cell-relative.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(state.doc)) });
    await expect(editor(page)).toContainText('AFTER TABLE');
    for (const zoom of [100, 50, 125, 100]) {
      const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
      await expect.poll(async () => { const g = await geometry(page); return Math.abs(g.image.top - g.first.top - 70); }).toBeLessThan(.5);
      const g = await geometry(page); expect(g.cell.bottom - g.image.bottom).toBeGreaterThanOrEqual(7); expect(g.pages).toBe(1);
    }
    const before = await editor(page).evaluate((root: any) => root.editor.getJSON()), initial = await geometry(page);
    await editor(page).focus();
    await editor(page).evaluate((root: any) => { let at = 0; root.editor.state.doc.descendants((node: any, pos: number) => { if (node.isTextblock && node.content.content.some((n: any) => n.type.name === 'inlineImage')) at = pos + 1 + node.content.size; }); root.editor.commands.setTextSelection(at); });
    await page.waitForTimeout(100);
    await page.keyboard.insertText(' EDITED ' + 'Cell text must wrap without displacing or losing its artwork. '.repeat(6));
    await expect(editor(page)).toContainText('Cell text must wrap');
    await page.keyboard.press('Enter'); await expect(editor(page)).toContainText('Cell text must wrap');
    await page.keyboard.insertText('NEW CELL BLOCK');
    await expect(editor(page)).toContainText('NEW CELL BLOCK');
    await expect.poll(async () => { const g = await geometry(page); return Math.abs(g.image.top - g.first.top - 70); }).toBeLessThan(.5);
    const edited = await geometry(page); expect(edited.blocks).toBeGreaterThan(initial.blocks); expect(edited.cell.height).toBeGreaterThan(initial.cell.height);
    expect(edited.cell.bottom - edited.image.bottom).toBeGreaterThanOrEqual(7);
    const overlaps = await editor(page).evaluate(root => {
      const image = root.querySelector('img[alt="CELL ART"]')!, box = image.getBoundingClientRect(), cell = image.closest('td,th')!, walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT);
      let overlaps = 0;
      while (walker.nextNode()) {
        const text = walker.currentNode as Text;
        if (text.parentElement?.closest('[data-flow-widget],.ProseMirror-widget')) continue;
        for (const part of new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text.data)) {
          if (!part.segment.trim()) continue;
          const range = document.createRange(); range.setStart(text, part.index); range.setEnd(text, part.index + part.segment.length);
          for (const ink of range.getClientRects()) if (Math.min(ink.right, box.right) - Math.max(ink.left, box.left) > .5 && Math.min(ink.bottom, box.bottom) - Math.max(ink.top, box.top) > .5) overlaps++;
        }
      }
      return overlaps;
    });
    expect(overlaps).toBe(0);
    let undos = 0;
    while (undos < 6 && JSON.stringify(await editor(page).evaluate((root: any) => root.editor.getJSON())) !== JSON.stringify(before)) {
      await page.keyboard.press('ControlOrMeta+z'); undos++; await page.waitForTimeout(120);
    }
    await expect.poll(() => editor(page).evaluate((root: any) => root.editor.getJSON())).toEqual(before);
    await expect.poll(async () => Math.abs((await geometry(page)).cell.height - initial.cell.height)).toBeLessThan(.5);
    for (let n = 0; n < undos; n++) await page.keyboard.press('ControlOrMeta+Shift+z');
    await expect(editor(page)).toContainText('NEW CELL BLOCK');
    await expect.poll(async () => { const g = await geometry(page); return Math.abs(g.image.top - g.first.top - 70); }).toBeLessThan(.5);
    for (let n = 0; n < undos; n++) await page.keyboard.press('ControlOrMeta+z');
    await expect.poll(() => editor(page).evaluate((root: any) => root.editor.getJSON())).toEqual(before);
    expect((await geometry(page)).src).toBe(state.printed.src);
  }
  expect(errors).toEqual([]);
});

for (const kind of ['merged', 'nested']) test(`cell-relative drawing grows and shrinks a ${kind} cell without keeping an old reservation`, async ({ page }) => {
  await page.goto('/');
  const doc = await page.evaluate(async kind => {
    const { parseDocumentHtml } = await import('/src/extensions.ts'), { defaultPage } = await import('/src/types.ts');
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const floating = { horizontalPosition: { relative: 'column', offset: 0 }, verticalPosition: { relative: 'page', offset: 70 * 9525 }, wrap: { type: 1 }, layoutInCell: true };
    const p = (text: string) => `<p style="font-family:Arial;font-size:11pt;line-height:20px;margin:0">${text}</p>`;
    const content = p('FIRST BLOCK') + p(`<img data-inline-image="true" alt="CELL ART" src="data:image/png;base64,${png}" width="140" height="80" data-image-floating="${JSON.stringify(floating).replaceAll('"', '&quot;')}">`) + p('LAST BLOCK');
    const inner = `<table style="width:400px"><colgroup><col style="width:200px"><col style="width:200px"></colgroup><tr style="height:180px" data-word-row-rule="atLeast"><td ${kind === 'merged' ? 'colspan="2" colwidth="200,200"' : 'colwidth="200"'} style="padding:8px;vertical-align:middle">${content}</td>${kind === 'merged' ? '' : `<td colwidth="200">${p('NEIGHBOR')}</td>`}</tr><tr><td colwidth="200">${p('LOWER LEFT')}</td><td colwidth="200">${p('LOWER RIGHT')}</td></tr></table>`;
    const html = p('BEFORE TABLE') + (kind === 'nested' ? `<table><tr><td>${p('OUTER BEFORE')}${inner}${p('OUTER AFTER')}</td></tr></table>` : inner) + p('AFTER TABLE');
    return { format: 'word-studio', version: 1, id: 'cell-resize', title: 'Cell resize', comments: [], updatedAt: new Date().toISOString(), page: { ...defaultPage, pageNumbers: false }, content: parseDocumentHtml(html) };
  }, kind);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'cell-resize.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(doc)) });
  await expect(editor(page)).toContainText('AFTER TABLE');
  await expect.poll(async () => { const g = await geometry(page); return Math.abs(g.image.top - g.first.top - 70); }).toBeLessThan(.5);
  const initial = await geometry(page), model = await editor(page).evaluate((root: any) => root.editor.getJSON());
  const resize = async (height: number) => editor(page).evaluate((root: any, height) => {
    const ed = root.editor; let at = -1; ed.state.doc.descendants((node: any, pos: number) => { if (node.type.name === 'inlineImage' && node.attrs.alt === 'CELL ART') at = pos; });
    if (at < 0) throw Error('Drawing missing'); const node = ed.state.doc.nodeAt(at); ed.view.dispatch(ed.state.tr.setNodeMarkup(at, undefined, { ...node.attrs, height }));
  }, height);
  await resize(180);
  await expect.poll(async () => (await geometry(page)).cell.height).toBeGreaterThan(initial.cell.height + 50);
  let state = await geometry(page); expect(state.image.top - state.first.top).toBeCloseTo(70, 0); expect(state.cell.bottom - state.image.bottom).toBeGreaterThanOrEqual(7);
  await resize(40);
  await expect.poll(async () => (await geometry(page)).cell.height).toBeLessThan(initial.cell.height + .5);
  state = await geometry(page); expect(state.image.height).toBeCloseTo(40, 0); expect(state.image.top - state.first.top).toBeCloseTo(70, 0);
  await editor(page).focus(); await page.keyboard.press('ControlOrMeta+z'); await page.keyboard.press('ControlOrMeta+z');
  await expect.poll(() => editor(page).evaluate((root: any) => root.editor.getJSON())).toEqual(model);
  await expect.poll(async () => Math.abs((await geometry(page)).cell.height - initial.cell.height)).toBeLessThan(.5);
  const printed = await page.evaluate(async () => {
    const root = document.querySelector<HTMLElement>('.tiptap[contenteditable=true]')! as any;
    const { createPrintLayout } = await import('/src/print-layout.ts'), { withoutEditorReflow } = await import('/src/editor-layout-task.ts'), { applyFloatingLayout } = await import('/src/floating-layout.ts');
    return withoutEditorReflow(async () => {
      const layout = await createPrintLayout(root.editor.getHTML(), (await import('/src/types.ts')).defaultPage);
      try {
        const image = layout.element.querySelector<HTMLImageElement>('img[alt="CELL ART"]')!, cell = image.closest('td,th')!, first = cell.querySelector('p')!;
        const snapshots = [];
        for (let pass = 0; pass < 3; pass++) { const im = image.getBoundingClientRect(), c = cell.getBoundingClientRect(); snapshots.push({ offset: im.top - first.getBoundingClientRect().top, cellHeight: c.height, bottomInside: c.bottom - im.bottom }); applyFloatingLayout(image.closest<HTMLElement>('.print-content')!); }
        return { pages: layout.pages, snapshots };
      } finally { layout.dispose(); }
    });
  });
  expect(printed.pages).toBe(1);
  for (const state of printed.snapshots) { expect(state.offset).toBeCloseTo(70, 0); expect(state.bottomInside).toBeGreaterThanOrEqual(7); expect(state.cellHeight).toBeCloseTo(printed.snapshots[0].cellHeight, 0); }
});
