import { expect, test, type Locator, type Page } from '@playwright/test';
import {
  Document, ImageRun, LineRuleType, Packer, PageBreak, Paragraph, TextRun,
  TextWrappingSide, TextWrappingType,
} from 'docx';

test.use({ deviceScaleFactor: 2 });

const fallbackFamily = 'Studio Apple Gothic space probe';
const fallbackAsset = '/src/assets/fonts/texgyreadventor-regular.otf';
const run = { font: 'Calibri', size: 24 };
const normalSpacing = { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO };
const image = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==', 'base64',
);

const scenarios = [
  { name: 'wrapped Latin', wrapped: true, mixed: false, lineHeight: 19, finalTextPage: 1 },
  { name: 'inline image', wrapped: false, mixed: false, lineHeight: 19.5, finalTextPage: 2 },
  { name: 'wrapped mixed script', wrapped: true, mixed: true, lineHeight: 19.5, finalTextPage: 2 },
] as const;

type Scenario = typeof scenarios[number];

function targetText(mixed: boolean) {
  return ['A', 'B', 'C'].map(letter => letter + '’s' + 'W'.repeat(32)).join(' ') + (mixed ? '한' : '');
}

function fixture({ wrapped, mixed }: Scenario) {
  const body = (style: string, text: string) => new Paragraph({
    style, widowControl: false, children: [new TextRun({ ...run, text })],
  });
  return Packer.toBuffer(new Document({
    creator: 'Independent layout QA author',
    compatibility: { wrapTrailSpaces: true },
    styles: {
      default: { document: { run } },
      paragraphStyles: ['DprBody', 'DprTarget'].map(id => ({
        id, name: id, basedOn: 'Normal', run, paragraph: { spacing: normalSpacing },
      })),
    },
    sections: [{
      // The Letter body is 864 CSS px high. The 45 lines occupy 855 px
      // at the reference height and 877.5 px at the native DPR 2 height.
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
      },
      children: [
        ...Array.from({ length: 42 }, (_, i) =>
          body('DprBody', 'Line ' + (i + 1) + ' leaves the printer’s measure unchanged.' + (mixed ? ' 한' : ''))),
        body('DprTarget', targetText(mixed)),
        new Paragraph({ spacing: { line: 15, lineRule: LineRuleType.EXACT }, children: [new PageBreak()] }),
        new Paragraph({ children: [new ImageRun({
          type: 'png', data: image, transformation: { width: 100, height: 40 },
          ...(wrapped ? {
            floating: {
              horizontalPosition: { relative: 'page', offset: 540 * 9525 },
              verticalPosition: { relative: 'page', offset: 116 * 9525 },
              wrap: { type: TextWrappingType.SQUARE, side: TextWrappingSide.BOTH_SIDES },
            },
          } : {}),
        })] }),
        body('DprBody', 'Picture’s second-sheet caption.' + (mixed ? ' 한' : '')),
      ],
    }],
  }));
}

async function emulateDensitySensitiveFallback(page: Page) {
  await page.addInitScript(({ family, asset }) => {
    const NativeFace = window.FontFace;
    window.FontFace = class extends NativeFace {
      constructor(name: string, source: string | BufferSource, options?: FontFaceDescriptors) {
        if (name === 'Studio Calibri presence probe') source = 'local("Word Studio QA missing Calibri")';
        if (name === family) source = 'url("' + asset + '")';
        super(name, source, options);
      }
    };
    const rect = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function () {
      if (!(this instanceof HTMLSpanElement)) return rect.call(this);
      const { style } = this;
      if (style.position !== 'fixed' || style.visibility !== 'hidden' || style.left !== '-10000px' ||
          style.whiteSpace !== 'pre' || style.fontFamily.replace(/["']/g, '') !== 'Calibri') return rect.call(this);
      if (style.fontSize === '64px' && ['Wm', 'i', '0', ' '].includes(this.textContent ?? '')) {
        style.fontFamily = JSON.stringify(family);
        try { return rect.call(this); }
        finally { style.fontFamily = 'Calibri'; }
      }
      const bounds = rect.call(this);
      if (style.lineHeight !== 'normal' || style.fontSize !== '16px') return bounds;
      // Model only the fallback probe. Real CSS zoom, line wrapping, glyph
      // bounds and page placement still use the browser.
      const zoom = this.currentCSSZoom;
      const density = devicePixelRatio * zoom;
      return new DOMRect(bounds.x, bounds.y, bounds.width, (density === 1 ? 19 : 19.5) * zoom);
    };
  }, { family: fallbackFamily, asset: fallbackAsset });
}

function readTarget(root: Locator, mode: 'editor' | 'print') {
  return root.locator('p[data-word-paragraph-style="DprTarget"]').evaluateAll((rows, view) => rows.map(row => {
    const editor = row.closest('.tiptap[contenteditable]')!;
    const stack = row.closest<HTMLElement>('.document-stack');
    const walker = document.createTreeWalker(row, NodeFilter.SHOW_TEXT);
    let last: Text | undefined;
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      if (node.data.trim() && !node.parentElement?.closest('[data-flow-widget],.pagination-cell-gap')) last = node;
    }
    if (!last) throw Error('Expected an authored text fragment');
    const range = document.createRange();
    range.setStart(last, last.data.length - 1);
    range.setEnd(last, last.data.length);
    const bottom = range.getBoundingClientRect().top;
    const paper = view === 'print'
      ? Number(row.closest<HTMLElement>('.print-page')?.dataset.page)
      : 1 + Math.floor((bottom - editor.getBoundingClientRect().top + .5) / (Number(stack?.dataset.pageHeight) + 24));
    return { paper, text: row.textContent ?? '', height: parseFloat(getComputedStyle(row).lineHeight) };
  }), mode);
}

for (const scenario of scenarios) {
  test('Word auto line height at DPR 2 with ' + scenario.name, async ({ page }) => {
    await emulateDensitySensitiveFallback(page);
    await page.goto('/');
    expect(await page.evaluate(() => devicePixelRatio)).toBe(2);
    await expect.poll(() => page.evaluate(family => [...document.fonts].some(face =>
      face.family.replace(/["']/g, '') === family && face.status === 'loaded'), fallbackFamily)).toBe(true);

    await page.locator('input[type=file]').first().setInputFiles({
      name: 'word-reference-density.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await fixture(scenario),
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue('word-reference-density');
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect.poll(() => readTarget(editor, 'editor')).toEqual([
      { paper: scenario.finalTextPage, text: targetText(scenario.mixed), height: scenario.lineHeight },
    ]);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const print = page.locator('.print-layout .print-content');
    await expect.poll(async () => (await readTarget(print, 'print')).at(-1)?.paper).toBe(scenario.finalTextPage);
    const fragments = await readTarget(print, 'print');
    expect(fragments.map(fragment => fragment.text).join('')).toBe(targetText(scenario.mixed));
    expect(fragments.map(fragment => fragment.paper)).toEqual(scenario.finalTextPage === 1 ? [1] : [1, 2]);
    for (const fragment of fragments) expect(fragment.height).toBe(scenario.lineHeight);
  });
}
