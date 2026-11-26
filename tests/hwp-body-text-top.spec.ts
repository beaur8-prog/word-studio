import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const wrapped = 'The paragraph wraps across the printable width and keeps its descenders gyp.';
const short = 'The final body line keeps its descenders gyp.';
const cases = [
  { scope: 'within a paragraph', lines: ['FIRST', short, wrapped], lineCounts: [1, 1, 2], top: 84, height: 64 },
  { scope: 'across a paragraph boundary', lines: ['FIRST', wrapped, short], lineCounts: [1, 2, 1], top: 116, height: 32 },
  { scope: 'with a smaller run on a measured body line', lines: ['FIRST', short, wrapped], lineCounts: [1, 1, 2], top: 84, height: 64, mixed: true },
  { scope: 'after consecutive one-line paragraphs', lines: ['FIRST', wrapped, 'A short continuation', short],
    lineCounts: [1, 2, 1, 1], top: 148, height: 32, paperHeight: 16050 },
];

for (const { scope, lines, lineCounts, top, height, mixed = false, paperHeight = 12000 } of cases) test(`HWP percentage leading ${scope} keeps the text box on its source page through two DOCX saves`, async ({ page }) => {
  const requireLocal = createRequire(import.meta.url);
  const library = dirname(requireLocal.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(join(library, 'rhwp.js')).href);
  await init({ module_or_path: await readFile(join(library, 'rhwp_bg.wasm')) });
  const source = HwpDocument.createEmpty();
  let bytes: Uint8Array;
  try {
    source.createBlankDocument();
    const html = lines.map((line, index) => {
      const content = mixed && index === 2 ? line.replace('paragraph', '<span style="font-size:13pt">paragraph</span>') : line;
      return `<p style="font-family:Arial;font-size:15pt;line-height:160%;margin:0"><span style="font-family:Arial;font-size:15pt">${content}</span></p>`;
    }).join('');
    source.pasteHtml(0, 0, 0, '<p></p>' + html + '<p></p>');
    source.deleteParagraph(0, source.getParagraphCount(0) - 1);
    source.deleteParagraph(0, 0);
    source.setPageDef(0, JSON.stringify({
      width: 37500, height: paperHeight, landscape: false, marginTop: 1500, marginBottom: 1500,
      marginLeft: 1500, marginRight: 1500, marginHeader: 0, marginFooter: 0, marginGutter: 0,
    }));
    const rendered = new HwpDocument(source.exportHwp());
    try {
      expect(rendered.pageCount()).toBe(1);
      const { runs } = JSON.parse(rendered.getPageTextLayout(0));
      const files = unzipSync(source.exportHwpx()), section = 'Contents/section0.xml';
      // Save the synthetic writer's measured line geometry in a normal HWPX
      // cache. The paragraph text still contains no manual line breaks.
      const pxToUnits = (px: number) => Math.round(px * 75);
      let index = 0;
      const xml = strFromU8(files[section]).replace(/<\/hp:p>/g, () => {
        const { lines: [line] } = JSON.parse(rendered.getParagraphLineMetrics(0, index, '[]'));
        const body = runs.filter((run: { paraIdx: number }) => run.paraIdx === index);
        const firstOnEachLine = body.filter((run: { y: number }, i: number) => !i || body[i - 1].y !== run.y);
        expect(firstOnEachLine).toHaveLength(lineCounts[index++]);
        const flags = line.flags & 0x7fffffff; // The output now includes stored metrics.
        const segments = firstOnEachLine.map((run: { y: number; charStart: number }) =>
          `<hp:lineseg textpos="${run.charStart}" vertpos="${pxToUnits(run.y - 20)}" vertsize="${pxToUnits(line.height)}" textheight="${pxToUnits(line.textHeight)}" baseline="${pxToUnits(line.baseline)}" spacing="${pxToUnits(line.spacing)}" horzpos="0" horzsize="${pxToUnits(line.width)}" flags="${flags}"/>`
        );
        return `<hp:linesegarray>${segments.join('')}</hp:linesegarray></hp:p>`;
      });
      expect(index).toBe(lines.length);
      files[section] = strToU8(xml);
      bytes = zipSync(files);
    } finally {
      rendered.free();
    }
  } finally {
    source.free();
  }
  const original = new HwpDocument(bytes);
  try {
    expect(original.pageCount()).toBe(1);
    expect(JSON.parse(original.getParagraphLineMetrics(0, lines.length - 1, '[]')).stored).toBe(true);
    const runs = JSON.parse(original.getPageTextLayout(0)).runs;
    expect(runs.find((run: { paraIdx: number }) => run.paraIdx === lines.length - 1).y).toBeCloseTo(top, 0);
  } finally {
    original.free();
  }

  await page.goto('/');
  const result = await page.evaluate(async ({ input, lines, mixed }) => {
    const { HwpReaderSession } = await import('/src/hwp/reader.ts');
    const { importEditableHwp } = await import('/src/hwp/editing.ts');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const reader = new HwpReaderSession();
    let document;
    try {
      await reader.open(new File([new Uint8Array(input)], 'original.hwpx'));
      document = importEditableHwp(await reader.editable(), 'original.hwpx');
    } finally {
      reader.close();
    }
    const startingDocument = document;
    const render = async (content: typeof document.content) => {
      const layout = await createPrintLayout(generateHTML(content, editorExtensions), document.page);
      try {
        const body = [...layout.element.querySelectorAll('.print-content p')];
        const last = body.find(block => block.textContent === lines.at(-1)) as HTMLElement | undefined;
        return {
          pages: layout.pages, text: body.map(block => block.textContent),
          mode: last?.dataset.hwpBodyLayout,
          height: last?.getBoundingClientRect().height,
          top: last && last.getBoundingClientRect().top - last.closest('.print-page')!.getBoundingClientRect().top,
        };
      } finally {
        layout.dispose();
      }
    };
    const control = structuredClone(document.content);
    control.content!.forEach(node => {
      if (node.attrs) node.attrs.hwpBodyLine = mixed ? node.attrs.hwpBodyLine?.replace(/^3 /, '1 ') : null;
    });
    const withoutSourceGeometry = await render(control);
    const stages = [];
    for (let save = 0; save < 3; save++) {
      stages.push(await render(document.content));
      if (save < 2) {
        const opened = await importFile(new File([await exportDocx(document)], 'saved.docx'));
        document = { ...opened.document, content: parseDocumentHtml(opened.html!) };
      }
    }
    return { startingDocument, withoutSourceGeometry, stages };
  }, { input: [...bytes], lines, mixed });
  expect(result.withoutSourceGeometry.pages).toBe(paperHeight === 12000 ? 2 : 1);
  if (paperHeight !== 12000) expect(result.withoutSourceGeometry.mode).toBeUndefined();
  for (const stage of result.stages) {
    expect(stage.pages).toBe(1);
    expect(stage.mode).toBe(mixed ? 'mixed-top' : 'top');
    expect(stage.text).toEqual(lines);
    expect(stage.top).toBeCloseTo(top, 0);
    expect(stage.height).toBeCloseTo(height, 0);
  }
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'native-line.wordstudio', mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify(result.startingDocument)),
  });
  const edge = page.getByRole('textbox', { name: '문서 편집 영역' }).locator('p').filter({ hasText: lines.at(-1) });
  await expect(edge).toHaveAttribute('data-hwp-body-layout', mixed ? 'mixed-top' : 'top');
  await expect(page.locator('.document-stack')).toHaveAttribute('data-pages', '1');
});
