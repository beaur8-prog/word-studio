import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

for (const font of ['inherit', 'Studio Text']) test(`${font}: final glyphs fit while trailing leading stays in document flow through zoom, undo and two DOCX saves`, async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async font => {
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const paper = { ...defaultPage, size: 'Custom', width: 150, height: 180, margin: 10, pageNumbers: false };
    if (font !== 'inherit') await document.fonts.load(`40pt "${font}"`);
    const body = 160 * 96 / 25.4;
    const html = `<p style="font-size:1px;margin:0;line-height:${body - 70}px"></p><p style="font-family:${font};font-size:40pt;line-height:80px;margin:0">FINAL</p>`;
    let doc = { format: 'word-studio', version: 1, id: 'fit-test', title: 'Fit', updatedAt: new Date().toISOString(), page: paper, comments: [], content: parseDocumentHtml(html) };
    const original = doc, states = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), paper);
      try {
        const last = [...layout.element.querySelectorAll('p')].find(p => p.textContent === 'FINAL')!;
        const content = last.closest('.print-content')!;
        states.push({ pages: layout.pages, text: last.textContent, html: last.outerHTML, height: last.getBoundingClientRect().height,
          leadingOutside: last.getBoundingClientRect().bottom - content.getBoundingClientRect().bottom });
      } finally { layout.dispose(); }
      if (save < 2) { const loaded = await importFile(new File([await exportDocx(doc)], 'again.docx')); doc = { ...loaded.document, content: parseDocumentHtml(loaded.html!) }; }
    }
    const guards = [];
    for (const tail of ['<p style="font-size:40pt;line-height:80px">FINAL</p><p>AFTER</p>', '<p style="font-size:40pt;line-height:80px"><mark>FINAL</mark></p>', '<p style="font-size:40pt;line-height:80px">FINAL<br></p>']) {
      const layout = await createPrintLayout(generateHTML(parseDocumentHtml(html.replace(/<p style="font-family:[\s\S]*$/, tail)), editorExtensions), paper);
      try { guards.push({ pages: layout.pages, first: layout.element.querySelector('.print-page')!.textContent }); } finally { layout.dispose(); }
    }
    return { original, states, guards };
  }, font);
  if (process.env.HWP_LAYOUT_DIAGNOSTICS) console.log(JSON.stringify(result));
  for (const state of result.states) {
    expect(state.pages).toBe(1); expect(state.text).toBe('FINAL');
    expect(state.height).toBeCloseTo(80, 0); expect(state.leadingOutside).toBeGreaterThan(9);
  }
  expect(result.guards[0].pages).toBe(2); expect(result.guards[0].first).toContain('FINAL'); expect(result.guards[0].first).not.toContain('AFTER');
  expect(result.guards[1].pages).toBe(2); expect(result.guards[2].pages).toBe(2);
  await page.locator('input[type=file]').first().setInputFiles({ name: 'fit.wordstudio', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(result.original)) });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역' });
  for (const zoom of [50, 100, 125, 200, 100]) {
    const slider = page.getByRole('slider', { name: '확대 축소' }); await slider.fill(String(zoom)); await slider.dispatchEvent('input'); await slider.dispatchEvent('change');
    await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
  }
  const last = editor.locator('p').filter({ hasText: /^FINAL$/ });
  await last.fill('WRAP '.repeat(100)); await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).not.toBe('1');
  await page.keyboard.press('ControlOrMeta+z'); await expect(last).toHaveText('FINAL'); await expect.poll(() => page.locator('.document-stack').getAttribute('data-pages')).toBe('1');
  const download = page.waitForEvent('download'); await page.keyboard.press('ControlOrMeta+s');
  const saved = JSON.parse(await fs.readFile((await (await download).path())!, 'utf8'));
  expect(saved.content).toEqual(result.original.content);
});

test('the first print loads a requested bundled face before measuring pages', async ({ page }) => {
  await page.route('**/NotoSansMono.ttf', async route => { await new Promise(resolve => setTimeout(resolve, 150)); await route.continue(); });
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts'), { createPrintLayout } = await import('/src/print-layout.ts');
    const html = '<p style="font-family:Studio Text;font-size:40pt;line-height:80px">FIRST PRINT</p>';
    const states = [];
    for (let i = 0; i < 2; i++) {
      const layout = await createPrintLayout(html, defaultPage);
      try { states.push({ loaded: document.fonts.check('40pt "Studio Text"'), pages: layout.pages, height: layout.element.querySelector('.print-content p')!.getBoundingClientRect().height }); }
      finally { layout.dispose(); }
      await document.fonts.ready;
    }
    return states;
  });
  expect(result[0].loaded).toBe(true); expect(result[0]).toEqual(result[1]);
});

test('mixed HWP body lines use their run heights at the page edge through two DOCX saves', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const pxPerMm = 96 / 25.4, margin = 10;
    const text = 'Large opener, then smaller words that wrap to another line with descenders gyp.';
    const html = '<p style="margin:0;font-size:10pt;line-height:56px">Following the source text:</p>' +
      '<p data-hwp-body-line="1 180 20 36" style="margin:0;font-family:Studio Text;font-size:15pt;line-height:36px">' +
      '<span style="font-size:15pt">Large opener,</span>' +
      '<span style="font-size:12pt"> then smaller words that wrap to another line with descenders gyp.</span></p>';
    const paper = { ...defaultPage, size: 'Custom' as const, width: 100, height: 130, margin, pageNumbers: false };
    const content = parseDocumentHtml(html);
    const sample = await createPrintLayout(generateHTML(content, editorExtensions), paper);
    let height: number;
    try {
      const tail = sample.element.querySelector<HTMLElement>('[data-hwp-body-layout="true"]')!;
      const sheetTop = tail.closest('.print-page')!.getBoundingClientRect().top;
      const range = document.createRange(); range.selectNodeContents(tail);
      const textBottom = Math.max(...[...range.getClientRects()].map(box => box.bottom)) - sheetTop;
      // The body is taller than even the text ranges with a small reserve.
      // The full final line box extends beyond that boundary.
      height = (textBottom + 1.5) / pxPerMm + margin;
    } finally { sample.dispose(); }
    let doc = {
      format: 'word-studio', version: 1, id: 'mixed-hwp-bottom-trim', title: 'Mixed HWP body',
      updatedAt: new Date().toISOString(), comments: [], page: { ...paper, height }, content,
    };
    const trimmed = [];
    for (let save = 0; save < 3; save++) {
      const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
      try {
        const tail = layout.element.querySelector<HTMLElement>('[data-hwp-body-layout="true"]')!;
        trimmed.push({ pages:layout.pages, text:tail.textContent,
          leadingOutside:tail.getBoundingClientRect().bottom-tail.closest('.print-content')!.getBoundingClientRect().bottom });
      } finally { layout.dispose(); }
      if (save < 2) {
        const imported = await importFile(new File([await exportDocx(doc)], 'mixed-body-save.docx'));
        doc = { ...imported.document, content: parseDocumentHtml(imported.html!) } as typeof doc;
      }
    }
    const guarded = await createPrintLayout(generateHTML(content, editorExtensions), { ...paper, height, suppressBottomSpacing: false });
    try { return { text, trimmed, reservedPages:guarded.pages }; }
    finally { guarded.dispose(); }
  });
  for (const stage of result.trimmed) {
    expect(stage.pages).toBe(1);
    expect(stage.text).toBe(result.text);
    expect(stage.leadingOutside).toBeGreaterThan(.5);
  }
  expect(result.reservedPages).toBe(2);
});

test('verified HWP top-aligned ink fits within the body while a descender outside moves to the next page', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { defaultPage } = await import('/src/types.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { paragraphFitBottom } = await import('/src/paragraph-fit.ts');
    const pxPerMm = 96 / 25.4, margin = 10;
    const text = 'Final gyp.';
    const html = '<p style="font-size:1px;line-height:320px;margin:0"></p>' +
      '<p data-hwp-body-line="2 150 20 30" data-word-line-rule="exact" ' +
      'style="font:20px/30px Studio Text;margin:0">' + text + '</p>';
    const paper = { ...defaultPage, size: 'Custom' as const, width: 120, height: 120, margin, pageNumbers: false };
    await document.fonts.load('20px "Studio Text"');
    const sample = await createPrintLayout(html, paper);
    let height: number, flowHeight: number;
    try {
      const tail = sample.element.querySelector<HTMLElement>('[data-hwp-body-layout="top"]')!;
      flowHeight = tail.getBoundingClientRect().height;
      const font = document.createElement('canvas').getContext('2d')!;
      font.font = '20px "Studio Text"';
      const metrics = font.measureText(text);
      const range = document.createRange();
      range.selectNodeContents(tail.querySelector('[data-hwp-body-run]')!);
      const rect = range.getBoundingClientRect();
      if (rect.height !== metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent) {
        throw new Error('Fixture requires matching DOM and canvas font boxes');
      }
      const ink = rect.top + metrics.fontBoundingBoxAscent + metrics.actualBoundingBoxDescent;
      const top = tail.closest('.print-page')!.getBoundingClientRect().top;
      height = (ink - top + .15) / pxPerMm + margin;
    } finally { sample.dispose(); }
    const states = [];
    for (const reduction of [0, 2]) {
      const layout = await createPrintLayout(html, { ...paper, height: height - reduction / pxPerMm });
      try {
        const tail = layout.element.querySelector<HTMLElement>('[data-hwp-body-layout="top"]')!;
        const body = tail.closest('.print-content')!.getBoundingClientRect();
        states.push({
          page: Number(tail.closest<HTMLElement>('.print-page')!.dataset.page),
          text: tail.textContent, flowHeight: tail.getBoundingClientRect().height,
          inkRoom: body.bottom - paragraphFitBottom(tail),
        });
      } finally { layout.dispose(); }
    }
    return { text, states, flowHeight };
  });
  expect(result.states.map(state => state.page)).toEqual([1, 2]);
  expect(result.states.map(state => state.text)).toEqual([result.text, result.text]);
  expect(result.states.map(state => state.flowHeight)).toEqual([result.flowHeight, result.flowHeight]);
  expect(result.states[0].inkRoom).toBeGreaterThan(0);
  expect(result.states[0].inkRoom).toBeLessThan(.4);
});

test('an authored HWP line top measures the glyphs on each existing visual line', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { paragraphFitBottom } = await import('/src/paragraph-fit.ts');
    const first = '읊 몫 좁 꽃 숲 붓 읊 몫 좁 꽃 숲 붓';
    const last = '가나다라';
    const font = '17.3333px "Noto Serif KR"';
    await document.fonts.load(font, first + last);
    const paragraph = document.createElement('p');
    paragraph.style.cssText = 'position:absolute;left:12px;top:12px;line-height:27.7px;margin:0;white-space:nowrap;width:max-content';
    paragraph.style.font = font;
    paragraph.style.lineHeight = '27.7px';
    paragraph.dataset.hwpBodyLayout = 'top';
    paragraph.textContent = first;
    document.body.append(paragraph);
    try {
      const width = paragraph.getBoundingClientRect().width;
      paragraph.style.width = `${width + .25}px`;
      paragraph.style.whiteSpace = 'normal';
      const literal = first + ' ' + last;
      paragraph.textContent = literal;
      const range = document.createRange();
      range.selectNodeContents(paragraph);
      const all = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
      range.setStart(paragraph.firstChild!, first.length + 1);
      const second = range.getBoundingClientRect();
      const context = document.createElement('canvas').getContext('2d')!;
      context.font = font;
      const metrics = context.measureText(literal);
      const fullRun = Math.max(second.top + metrics.fontBoundingBoxAscent, second.bottom - metrics.fontBoundingBoxDescent) +
        metrics.actualBoundingBoxDescent;
      const measured = paragraphFitBottom(paragraph);
      return { lines: all.length, literal: paragraph.textContent, expected: literal, fullRun, measured, secondTop: second.top };
    } finally { paragraph.remove(); }
  });
  expect(result.lines).toBe(2);
  expect(result.literal).toBe(result.expected);
  expect(result.measured).toBeGreaterThan(result.secondTop);
  expect(result.measured).toBeLessThan(result.fullRun - .125);
});
