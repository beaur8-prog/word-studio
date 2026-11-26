import { test, expect } from '@playwright/test';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createRequire } from 'node:module';

const { createLongMergeFixture } = createRequire(import.meta.url)('../hwp-merge-row-break-fixture.cjs');
const cases = [
  { count: 46, rowHeight: 1022, sourceEnds: [13, 32, 46] },
  { count: 84, rowHeight: 1850, sourceEnds: [13, 32, 51, 70, 84] },
];

for (const fixture of cases) test('Measure last glyphs in a ' + fixture.sourceEnds.length + '-page continued merge', async ({ page }) => {
  test.setTimeout(90_000);
  const directory = await mkdtemp(join(tmpdir(), 'studio-merge-witness-'));
  try {
    const files = await createLongMergeFixture(directory, fixture);
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles(files.hwp);
    await page.getByRole('button', { name: 'Word 화면에서 편집', exact: true }).click();
    await page.getByRole('button', { name: '편집본 만들기', exact: true }).click();
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(page.locator('.document-stack')).toHaveAttribute('data-pages', String(fixture.sourceEnds.length));
    await page.evaluate(() => document.fonts.ready);
    const measure = async (printed: boolean) => page.evaluate(async ({ count, printed }) => {
      const { measuredHwpLineInk } = await import('/src/hwp-line-ink.ts');
      const root = document.querySelector<HTMLElement>(printed ? '.print-layout.preview' : '.tiptap.ProseMirror')!;
      const sheet = root.closest<HTMLElement>('.document-stack');
      const origin = root.getBoundingClientRect().top;
      const scale = sheet ? sheet.getBoundingClientRect().width / parseFloat(getComputedStyle(sheet).width) : 1;
      const height = Number(sheet?.dataset.pageHeight), stride = height + 24;
      const picks = [count - 2, count - 1, count];
      return [...root.querySelectorAll<HTMLParagraphElement>('p')].flatMap(node => {
        const match = /^TARGET RIGHT (\d+)$/.exec(node.textContent ?? '');
        if (!match || !picks.includes(+match[1])) return [];
        const range = document.createRange(); range.selectNodeContents(node);
        const rect = range.getBoundingClientRect(), cell = node.closest('td')!.getBoundingClientRect();
        const p = node.getBoundingClientRect(), css = getComputedStyle(node);
        const paper = printed ? Number(node.closest<HTMLElement>('.print-page')!.dataset.page) :
          Math.floor(((rect.top - origin) / scale + .5) / stride) + 1;
        const body = printed ? node.closest('.print-content')!.getBoundingClientRect() : {
          top: origin + ((paper - 1) * stride + Number(sheet!.dataset.margin)) * scale,
          bottom: origin + ((paper - 1) * stride + height - Number(sheet!.dataset.bottomMargin)) * scale,
        };
        const baseline = measuredHwpLineInk(node, scale, parseFloat(css.fontSize));
        const compact = (top: number, bottom: number) => [top - p.top, bottom - p.top].map(n => +n.toFixed(3));
        return [{ marker: +match[1], paper, font: css.fontFamily, lineHeight: css.lineHeight, pHeight: p.height,
          range: compact(rect.top, rect.bottom), cell: compact(cell.top, cell.bottom), body: compact(body.top, body.bottom),
          ink: baseline && compact(baseline.top, baseline.bottom), baseline: baseline && +(baseline.baseline - p.top).toFixed(3) }];
      });
    }, { count: fixture.count, printed });
    const live = await measure(false);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    await expect(page.locator('.print-layout.preview .print-page')).toHaveCount(fixture.sourceEnds.length);
    const printed = await measure(true);
    expect(live).toHaveLength(3);
    expect(printed).toHaveLength(3);
    console.log('WINDOWS_LONG_MERGE_WITNESS ' + JSON.stringify({ count: fixture.count, live, printed }));
  } finally { await rm(directory, { recursive: true, force: true }); }
});
