import { test, expect } from '@playwright/test';
import { Document, LineRuleType, Packer, Paragraph, TextRun } from 'docx';
import { writeFile } from 'node:fs/promises';

const lineTokens = ['LINE0', 'LINE1', 'LINE2', 'LINE3'];

async function fixture(capacity: number, widowControl: boolean, wrapped: boolean) {
  // A 400px sheet with two 50px margins gives 300px of body space.
  // Leave five spare pixels beyond the chosen number of exact 20px lines.
  const leadHeight = 300 - capacity * 20 - 5;
  const paragraph = (tokens: string[], exactHeight: number, widowControl: boolean) => new Paragraph({
    children: tokens.map((text, index) => new TextRun({
      text, font: 'Arial', size: 18, break: index > 0 ? 1 : 0,
    })),
    spacing: { before: 0, after: 0, line: exactHeight * 15, lineRule: LineRuleType.EXACT },
    widowControl,
  });
  const target = wrapped
    // Each fixed-width word fits inside 300px, but two cannot share a line.
    ? new Paragraph({
      children: [new TextRun({
        text: lineTokens.map(token => token + 'x'.repeat(30)).join(' '), font: 'Courier New', size: 18,
      })],
      spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.EXACT },
      widowControl,
    })
    : paragraph(lineTokens, 20, widowControl);
  return Packer.toBuffer(new Document({
    sections: [{
      properties: { page: { size: { width: 6000, height: 6000 }, margin: { top: 750, right: 750, bottom: 750, left: 750 } } },
      children: [paragraph(['LEAD'], leadHeight, false), target],
    }],
  }));
}

const cases = [
  ...[1, 2, 3].flatMap(capacity => [true, false].map(enabled => ({ capacity, enabled, wrapped: false }))),
  ...[1, 3].map(capacity => ({ capacity, enabled: true, wrapped: true })),
];

for (const { capacity, enabled, wrapped } of cases) {
  test(`Word widow control ${enabled ? 'on' : 'off'} when ${capacity} of 4 ${wrapped ? 'wrapped' : 'explicit'} lines fit the first page`, async ({ page }, info) => {
    const input = await fixture(capacity, enabled, wrapped);
    await page.goto('/');
    await page.locator('input[type=file]').first().setInputFiles({
      name: 'widow-split.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: input,
    });
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    const target = editor.locator('p').filter({ hasText: 'LINE0' });
    await expect(target).toContainText('LINE3');
    await expect(target).toHaveCSS('orphans', enabled ? '2' : '1');
    await expect(target).toHaveCSS('widows', enabled ? '2' : '1');
    await expect(target.locator('br:not(.ProseMirror-trailingBreak)')).toHaveCount(wrapped ? 0 : 3);
    await page.evaluate(() => document.fonts.ready);

    const measureEditor = () => editor.evaluate(root => {
      const sheet = root.closest<HTMLElement>('.document-stack')!;
      const height = Number(sheet.dataset.pageHeight);
      const topMargin = Number(sheet.dataset.margin);
      const bottomMargin = Number(sheet.dataset.bottomMargin);
      const origin = root.getBoundingClientRect().top;
      const stride = height + 24;
      const pages: string[][] = [];
      const outsideBody: string[] = [];
      const distinctTargetTops = new Set<number>();
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const text = walker.currentNode as Text;
        if (text.parentElement?.closest('[data-flow-widget]')) continue;
        for (const match of text.data.matchAll(/LEAD|LINE[0-3]/g)) {
          const range = document.createRange();
          range.setStart(text, match.index!);
          range.setEnd(text, match.index! + match[0].length);
          const bounds = range.getBoundingClientRect();
          const y = bounds.top - origin;
          const pageIndex = Math.max(0, Math.floor((y + .5) / stride));
          if (match[0] !== 'LEAD') distinctTargetTops.add(Math.round(y));
          const bodyY = y - pageIndex * stride;
          if (bodyY < topMargin - 1 || bodyY + bounds.height > height - bottomMargin + 1) outsideBody.push(match[0]);
          (pages[pageIndex] ??= []).push(match[0]);
        }
      }
      return { pages, outsideBody, targetLineCount: distinctTargetTops.size };
    });
    await expect.poll(async () => (await measureEditor()).pages.length).toBe(2);
    const before = await measureEditor();

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const print = page.locator('.print-content');
    await expect(print).toHaveCount(2);
    const printed = (await print.allTextContents()).map(text => text.match(/LEAD|LINE[0-3]/g) ?? []);
    const actual = { editor: before, print: printed };
    await writeFile(info.outputPath('line-placement.json'), JSON.stringify(actual, null, 2));

    const split = enabled ? (capacity === 1 ? 0 : 2) : capacity;
    const expected = [['LEAD', ...lineTokens.slice(0, split)], lineTokens.slice(split)];
    expect.soft(before.targetLineCount, 'every target token occupies a different visual line').toBe(4);
    expect.soft(before.pages, 'editable pages').toEqual(expected);
    expect.soft(before.outsideBody, 'no glyphs outside the page body').toEqual([]);
    expect(printed, 'printed pages').toEqual(expected);
  });
}
