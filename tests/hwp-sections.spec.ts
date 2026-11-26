import { test, expect } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: edited section paper, margins and plain furniture survive two saves`, async ({ page }) => {
  await page.goto('/');
  const states = await page.evaluate(async format => {
    const { exportEditableHwp, importEditableHwp } = await import('/src/hwp/editing.ts');
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const initial = { size: 'A4', orientation: 'portrait', margin: 20, header: 'FIRST HEADER', footer: 'FIRST FOOTER', pageNumbers: false };
    const second = { ...initial, size: 'Letter', orientation: 'landscape', margins: { top: 25, right: 15, bottom: 23, left: 17 }, header: 'SECOND HEADER', footer: 'SECOND FOOTER' };
    const third = { ...initial, size: 'Custom', width: 180, height: 240, margin: 18, header: '', footer: '' };
    const section = settings => `<div data-section-page="${JSON.stringify(settings).replaceAll('"', '&quot;')}"></div>`;
    let doc: any = { format: 'word-studio', version: 1, id: crypto.randomUUID(), title: 'Sections', comments: [], updatedAt: new Date().toISOString(), page: initial,
      content: parseDocumentHtml('<p>FIRST BODY</p>' + section(second) + '<p>SECOND BODY</p><div data-page-break="true"></div><p>SECOND NEXT PAGE</p>' + section(third) + '<p>THIRD BODY</p>') };
    const states = [];
    for (let round = 0; round < 3; round++) {
      const html = generateHTML(doc.content, editorExtensions), layout = await createPrintLayout(html, doc.page);
      try { states.push({ pages: [...layout.element.querySelectorAll<HTMLElement>('.print-page')].map(sheet => ({ width: parseFloat(sheet.style.width) * 25.4 / 96, height: parseFloat(sheet.style.height) * 25.4 / 96, body: sheet.querySelector('.print-content')!.textContent, header: sheet.querySelector('.print-header')!.textContent, footer: sheet.querySelector('.print-footer')!.textContent, left: parseFloat((sheet.querySelector('.print-content') as HTMLElement).style.left) * 25.4 / 96 })) }); }
      finally { layout.dispose(); }
      if (round === 2) break;
      const out = await exportEditableHwp(doc, format), session = new HwpReaderSession();
      if (out.warnings.some(w => /첫 용지 설정으로 통합|내용 손실/.test(w))) throw Error(out.warnings.join('\n'));
      try { await session.open(new File([out.blob], 'sections.' + format)); doc = importEditableHwp(await session.editable(), 'sections.' + format); }
      finally { session.close(); }
    }
    return states;
  }, format);
  for (const state of states) {
    expect(state.pages.map(p => p.body)).toEqual(['FIRST BODY', 'SECOND BODY', 'SECOND NEXT PAGE', 'THIRD BODY']);
    expect(state.pages.map(p => p.header)).toEqual(['FIRST HEADER', 'SECOND HEADER', 'SECOND HEADER', '']);
    expect(state.pages.map(p => p.footer)).toEqual(['FIRST FOOTER', 'SECOND FOOTER', 'SECOND FOOTER', '']);
    for (const [index, [width, height, left]] of [[210, 297, 20], [279.4, 215.9, 17], [279.4, 215.9, 17], [180, 240, 18]].entries()) {
      expect(Math.abs(state.pages[index].width - width)).toBeLessThan(.02);
      expect(Math.abs(state.pages[index].height - height)).toBeLessThan(.02);
      expect(Math.abs(state.pages[index].left - left)).toBeLessThan(.02);
    }
  }
});

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: retains the paper definition on a blank leading paragraph`, async ({ page }) => {
  await page.goto('/');
  const bytes = await page.evaluate(async format => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { defaultPage } = await import('/src/types.ts');
    const first = { ...defaultPage, margin: 14, header: 'FIRST HEADER', pageNumbers: false };
    const second = { ...first, size: 'Letter', orientation: 'landscape', margin: 19, header: 'SECOND HEADER' };
    const boundary = JSON.stringify(second).replaceAll('"', '&quot;');
    const html = `<p>FIRST BODY</p><div data-section-page="${boundary}"></div><p></p><p>SECOND BODY</p>`;
    const session = new HwpReaderSession();
    try {
      const saved = await session.export(html, first, format);
      return [...saved.bytes];
    } finally { session.close(); }
  }, format);
  const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = new HwpDocument(Uint8Array.from(bytes));
  try {
    expect(doc.getSectionCount()).toBe(2);
    expect(doc.pageCount()).toBe(2);
    expect(doc.getParagraphLength(1, 0)).toBe(0);
    expect(doc.getTextRange(1, 1, 0, 100)).toBe('SECOND BODY');
    expect(JSON.parse(doc.getPageDef(1))).toMatchObject({
      width: 61200, height: 79200, landscape: true,
      marginLeft: Math.round(19 * 7200 / 25.4), marginRight: Math.round(19 * 7200 / 25.4),
    });
    const section = strFromU8(unzipSync(doc.exportHwpx())['Contents/section1.xml']);
    expect(section.match(/<hp:secPr\b/g)).toHaveLength(1);
    expect(section.match(/<hp:p\b[^>]*>[\s\S]*?<\/hp:p>/)?.[0]).toContain('<hp:secPr');
  } finally { doc.free(); }
});

for (const format of ['hwp', 'hwpx'] as const) test(`${format}: a new section keeps an explicit two-column group after its full-width heading`, async ({ page }) => {
  await page.goto('/');
  const bytes = await page.evaluate(async format => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { defaultPage } = await import('/src/types.ts');
    const first = { ...defaultPage, pageNumbers: false, columns: { count: 2, space: 18 } };
    const next = { ...first, columns: { count: 1, space: 0 } };
    const boundary = JSON.stringify(next).replaceAll('"', '&quot;');
    const html = '<p>COVER</p><div data-section-page="' + boundary + '"></div>' +
      '<p>FULL WIDTH</p><div data-column-count="2" data-column-space="24" data-column-balance="false">' +
      '<p>LEFT COLUMN</p><div data-column-break="true"></div><p>RIGHT COLUMN</p></div><p>WIDE TAIL</p>';
    const session = new HwpReaderSession();
    try { return [...(await session.export(html, first, format)).bytes]; }
    finally { session.close(); }
  }, format);
  const require = createRequire(import.meta.url), core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(path.join(core, 'rhwp_bg.wasm')) });
  const doc = new HwpDocument(Uint8Array.from(bytes));
  try {
    expect(doc.getSectionCount()).toBe(2);
    expect(doc.pageCount()).toBe(2);
    const { runs } = JSON.parse(doc.getPageTextLayout(1));
    expect(runs.map(run => run.text).join('')).toBe('FULL WIDTHLEFT COLUMNRIGHT COLUMNWIDE TAIL');
    const left = runs.find(run => run.text === 'LEFT COLUMN'), right = runs.find(run => run.text === 'RIGHT COLUMN');
    expect(right.x - left.x).toBeGreaterThan(200);
    expect(right.y).toBeCloseTo(left.y, 0);
    const section = strFromU8(unzipSync(doc.exportHwpx())['Contents/section1.xml']);
    expect([...section.matchAll(/<hp:colPr\b[^>]*\bcolCount="(\d+)"/g)].map(match => Number(match[1]))).toEqual([1, 2, 1]);
  } finally { doc.free(); }
});
