import { test, expect } from '@playwright/test';
import { Bookmark, Document, HeadingLevel, HorizontalPositionRelativeFrom, ImageRun, LineRuleType, Packer, Paragraph, TextRun, TextWrappingType, VerticalPositionRelativeFrom } from 'docx';
import { writeFile } from 'node:fs/promises';

const levels = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6];
const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';

// Independent source documents; no source Word samples or oracle keys added.
async function fixture(level: number, kind: string) {
  const run = { font: 'Arial', size: 52, bold: false };
  const defaults: any = { document: { run } };
  for (let n = 1; n <= 6; n++) defaults['heading' + n] = { run, paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO } } };
  const children = kind === 'bookmark' ? [new Bookmark({ id: 'empty_heading_anchor', children: [] })] : kind === 'floating' ? [new ImageRun({
    type: 'png', data: Buffer.from(png, 'base64'), transformation: { width: 40, height: 40 },
    floating: { horizontalPosition: { relative: HorizontalPositionRelativeFrom.COLUMN, offset: 0 }, verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 }, wrap: { type: TextWrappingType.SQUARE } },
  })] : [];
  return [...await Packer.toBuffer(new Document({ styles: { default: defaults }, sections: [{ children: [
    new Paragraph({ heading: levels[level - 1], spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO }, children }),
    new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'FOLLOW MARKER', font: 'Arial', size: 24 })] }),
  ] }] }))];
}

for (const level of [1, 2, 3, 4, 5, 6]) for (const kind of ['empty', 'bookmark', 'floating']) {
  test(`H${level} ${kind}: print retains the authored empty line through preview and two DOCX saves`, async ({ page }, info) => {
    await page.goto('/'); const input = await fixture(level, kind);
    await page.locator('input[type=file]').first().setInputFiles({ name: 'empty-heading.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(input) });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor).toContainText('FOLLOW MARKER'); await page.evaluate(() => document.fonts.ready);
    const before = await editor.evaluate((root: any) => root.editor.getJSON());
    const measured = await editor.evaluate((root, level) => {
      const h = root.querySelector('h' + level)!; const next = [...root.querySelectorAll('p')].find(p => p.textContent === 'FOLLOW MARKER')!;
      return { height: h.getBoundingClientRect().height, followingGap: next.getBoundingClientRect().top - h.getBoundingClientRect().top };
    }, level);
    expect(measured.height).toBeGreaterThan(20);
    await page.getByRole('button', { name: '보기', exact: true }).click(); await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const printed = page.locator('.print-content h' + level).first(); await expect(printed).toBeAttached();
    await expect.poll(() => printed.evaluate(h => h.getBoundingClientRect().height)).toBeGreaterThan(20);
    const print = await page.locator('.print-layout').evaluate((root, level) => {
      const h = root.querySelector('.print-content h' + level)!; const next = [...root.querySelectorAll('.print-content p')].find(p => p.textContent === 'FOLLOW MARKER')!;
      return { height: h.getBoundingClientRect().height, followingGap: next.getBoundingClientRect().top - h.getBoundingClientRect().top };
    }, level);
    expect(Math.abs(print.height - measured.height)).toBeLessThan(1);
    expect(Math.abs(print.followingGap - measured.followingGap)).toBeLessThan(1);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
    expect(await editor.evaluate((root: any) => root.editor.getJSON())).toEqual(before);
    const result = await page.evaluate(async ({ bytes, level }) => {
      const { importFile, exportDocx } = await import('/src/document-io.ts'); const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js'); const { createPrintLayout } = await import('/src/print-layout.ts');
      let file = new File([new Uint8Array(bytes)], 'empty-heading.docx'); const states = [], saves = [];
      for (let save = 0; save < 3; save++) {
        const opened = await importFile(file), doc = { ...opened.document, content: parseDocumentHtml(opened.html!) }; const html = generateHTML(doc.content, editorExtensions), source = document.createElement('div'); source.innerHTML = html;
        const layout = await createPrintLayout(html, doc.page);
        try {
          const h = layout.element.querySelector<HTMLElement>('.print-content h' + level)!;
          states.push({ height: h.getBoundingClientRect().height, text: source.textContent, breaks: source.querySelectorAll('br').length, pictures: [...source.querySelectorAll('img')].map(img => img.getAttribute('src')), bookmarks: [...source.querySelectorAll('[data-docx-bookmark]')].map(n => n.getAttribute('data-docx-bookmark')) });
        } finally { layout.dispose(); }
        if (save < 2) { const blob = await exportDocx(doc); saves.push([...new Uint8Array(await blob.arrayBuffer())]); file = new File([blob], 'again.docx'); }
      }
      return { states, saves };
    }, { bytes: input, level });
    for (const state of result.states) {
      expect(Math.abs(state.height - measured.height)).toBeLessThan(1); expect(state.text).toBe(result.states[0].text); expect(state.breaks).toBe(0);
      expect(state.pictures).toEqual(result.states[0].pictures); expect(state.bookmarks).toEqual(result.states[0].bookmarks);
    }
    for (const [i, bytes] of result.saves.entries()) await writeFile(info.outputPath(`textbox-save-empty-heading-${level}-${kind}-${i + 1}.docx`), Buffer.from(bytes));
  });
}

test('empty heading struts preserve hard breaks, inline objects and explicitly collapsed layout', async ({ page }) => {
  await page.goto('/');
  const rows = await page.evaluate(async png => {
    const { prepareParagraphEnds } = await import('/src/paragraph-fragments.ts');
    const host = document.createElement('div'); host.className = 'tiptap';
    host.style.cssText = 'position:fixed;left:-10000px;top:0;font:26px Arial;line-height:32px;width:600px';
    const cases = [
      ['ordinary-empty-paragraph', '<p></p>'],
      ['nonempty-heading', '<h1>Visible title</h1>'],
      ['inline-picture', `<h1><img width="40" height="40" src="data:image/png;base64,${png}"></h1>`],
      ['inline-drawing', '<h1><span data-drawing-group-anchor="inline" style="display:inline-block;width:40px;height:80px"></span></h1>'],
      ['svg-picture', '<h1><svg width="40" height="80"></svg></h1>'],
      ['explicit-zero-height', '<h1 style="height:0;min-height:0;line-height:0;font-size:0"></h1>'],
    ];
    document.body.append(host);
    try {
      return cases.map(([name, html]) => {
        host.innerHTML = html; const block = host.firstElementChild as HTMLElement;
        block.style.margin = '0'; const before = block.getBoundingClientRect().height, content = block.innerHTML;
        prepareParagraphEnds(host); prepareParagraphEnds(host);
        return { name, before, after: block.getBoundingClientRect().height, contentUnchanged: block.innerHTML === content };
      }).concat((() => {
        host.innerHTML = '<h1 style="font:26px Arial;line-height:32px;margin:0"><br><br class="ProseMirror-trailingBreak"></h1>';
        const editorHeight = host.firstElementChild!.getBoundingClientRect().height;
        host.querySelector('.ProseMirror-trailingBreak')!.remove();
        prepareParagraphEnds(host); prepareParagraphEnds(host);
        return [{ name: 'authored-final-break', before: editorHeight, after: host.firstElementChild!.getBoundingClientRect().height, contentUnchanged: host.querySelectorAll('br').length === 1 }];
      })());
    } finally { host.remove(); }
  }, png);
  for (const row of rows) {
    expect(Math.abs(row.after - row.before), row.name).toBeLessThan(.1);
    expect(row.contentUnchanged, row.name).toBe(true);
  }
});

test('preparing an empty heading twice keeps one line and releases it when content changes', async ({ page }) => {
  await page.goto('/');
  const heights = await page.evaluate(async () => {
    const { prepareParagraphEnds } = await import('/src/paragraph-fragments.ts');
    const host = document.createElement('div');
    host.className = 'tiptap';
    host.style.cssText = 'position:fixed;left:-10000px;top:0;width:600px';
    const heading = document.createElement('h2');
    heading.style.cssText = 'font:26px Arial;line-height:32px;margin:0';
    host.append(heading);
    document.body.append(host);
    try {
      const measure = () => ({
        height: heading.getBoundingClientRect().height,
        reservesLine: heading.hasAttribute('data-print-empty-heading'),
        breaks: heading.querySelectorAll('br').length,
      });

      const before = measure();
      prepareParagraphEnds(host);
      const empty = measure();
      prepareParagraphEnds(host);
      const twice = measure();

      heading.textContent = 'Title';
      prepareParagraphEnds(host);
      const title = measure();

      heading.replaceChildren();
      prepareParagraphEnds(host);
      const emptyAgain = measure();

      return { before, empty, twice, title, emptyAgain };
    } finally {
      host.remove();
    }
  });

  expect(heights.before).toEqual({ height: 0, reservesLine: false, breaks: 0 });
  expect(heights.empty).toEqual({ height: 32, reservesLine: true, breaks: 0 });
  expect(heights.twice).toEqual(heights.empty);
  expect(heights.title).toEqual({ height: 32, reservesLine: false, breaks: 0 });
  expect(heights.emptyAgain).toEqual(heights.empty);
});
