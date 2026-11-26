import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { BorderStyle, Document, LineRuleType, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx';
import JSZip from 'jszip';
import { PreservedTableCell } from '../src/docx-table-cell';
import { hwpWrapFingerprint } from '../src/hwp-text-advances';

const noBorder = { style: BorderStyle.NONE, size: 0 } as const;
const labels = ['구분 가', '구분 나'];
const longCell = Array.from({ length: 18 }, (_, i) => 'CONTENT ' + i);
const modes = ['observed', 'unmarked', 'changed-label', 'changed-width', 'single-observed', 'single-unmarked'] as const;
const cellLabels = (mode: typeof modes[number]) => mode.startsWith('single-') ? ['07 안내 항목'] :
  mode === 'changed-label' ? [labels[0], '수정 나'] : labels;
const paragraph = (text: string, line = 20) => new Paragraph({
  children: [new TextRun({ text, font: 'Arial', size: 18 })],
  spacing: { before: 0, after: 0, line: line * 15, lineRule: LineRuleType.EXACT },
});
const margins = { top: 0, right: 0, bottom: 0, left: 0 };

async function fixture(mode: typeof modes[number]) {
  const leading = mode.endsWith('unmarked') ? undefined :
    (mode === 'changed-width' ? 88 : 90) + ' ' + hwpWrapFingerprint(JSON.stringify(mode.startsWith('single-') ? cellLabels(mode) : labels));
  const first = new PreservedTableCell({
    children: cellLabels(mode).map(text => paragraph(text)),
    verticalAlign: 'center', margins,
  }, false, leading);
  const table = new Table({
    width: { size: 4500, type: WidthType.DXA }, layout: 'fixed', columnWidths: [1350, 3150],
    borders: { top: noBorder, right: noBorder, bottom: noBorder, left: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [new TableRow({ cantSplit: false, children: [first,
      new TableCell({ children: longCell.map(line => paragraph(line)), verticalAlign: 'top', margins })] })],
  });
  return Packer.toBuffer(new Document({ sections: [{ properties: { page: {
    size: { width: 6000, height: 6000 }, margin: { top: 750, right: 750, bottom: 750, left: 750 },
  } }, children: [paragraph('LEAD', 197), table] }] }));
}

for (const mode of modes) {
  test(`keep HWP labels in the first printed row fragment only with their valid source measurement: ${mode}`, async ({ page }, info) => {
    let bytes = await fixture(mode);
    await page.goto('/');
    const unchanged = mode.endsWith('observed');
    const expected = cellLabels(mode);
    for (let stage = 0; stage <= (unchanged ? 2 : 0); stage++) {
      const title = 'hwp-centered-labels-' + mode + '-' + stage;
      await page.locator('input[type=file]').first().setInputFiles({
        name: title + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes,
      });
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(title);
      await page.getByRole('button', { name: '보기', exact: true }).click();
      await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
      const preview = page.locator('.print-layout.preview');
      await expect(preview.locator('.print-page')).toHaveCount(2);
      const physical = await preview.locator('.print-content td > p').evaluateAll(nodes => nodes.map(p => {
        const td = p.closest('td')!, cell = td.getBoundingClientRect(), r = document.createRange();
        r.selectNodeContents(p);
        const boxes = [...r.getClientRects()].filter(box => box.width > 0 && box.height > 0);
        const combined = r.getBoundingClientRect();
        return { text: p.textContent, page: Number(p.closest<HTMLElement>('.print-page')!.dataset.page),
          align: getComputedStyle(td).verticalAlign, centered: (combined.top + combined.bottom - cell.top - cell.bottom) / 2,
          inside: boxes.length > 0 && boxes.every(box => box.left >= cell.left - .5 && box.right <= cell.right + .5 &&
            box.top >= cell.top - .5 && box.bottom <= cell.bottom + .5) };
      }));
      const seen = physical.filter(line => line.text !== null && expected.includes(line.text));
      expect(physical.filter(line => line.text?.startsWith('CONTENT')).map(line => line.text)).toEqual(longCell);
      expect(seen).toHaveLength(expected.length);
      expect(seen.every(line => line.align === 'middle' && line.inside)).toBe(true);
      expect(seen.map(line => line.page)).toEqual(expected.map(() => unchanged ? 1 : 2));
      expect(Math.abs(seen.reduce((sum, line) => sum + line.centered, 0) / expected.length)).toBeLessThan(4);
      await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
      if (!unchanged || stage === 2) break;
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const target = info.outputPath(`textbox-save-hwp-centered-labels-${stage + 1}.docx`);
      await (await pending).saveAs(target);
      bytes = await readFile(target);
      const zip = await JSZip.loadAsync(bytes), xml = await zip.file('word/document.xml')!.async('string');
      const cell = [...xml.matchAll(/<w:tc(?:\s[^>]*?)?>[\s\S]*?<\/w:tc>/g)].find(match => match[0].includes(expected[0]))?.[0];
      expect(cell).toContain('wcl:leading=');
      expect(cell).toMatch(/<w:vAlign w:val="center"\s*\/>/);
    }
  });
}
