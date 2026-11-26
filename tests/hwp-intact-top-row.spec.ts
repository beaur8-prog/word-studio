import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { BorderStyle, Document, HeightRule, LineRuleType, Packer, Paragraph, TableCell, TableRow, TextRun, WidthType } from 'docx';
import JSZip from 'jszip';
import { PreservedTable, PreservedTableRow } from '../src/docx-table-cell';
import { hwpWrapFingerprint } from '../src/hwp-text-advances';

const noBorder = { style: BorderStyle.NONE, size: 0 } as const;
const paragraph = (text: string, line = 20) => new Paragraph({
  children: [new TextRun({ text, font: 'Arial', size: 18 })],
  spacing: { before: 0, after: 0, line: line * 15, lineRule: LineRuleType.EXACT },
});
const cell = (lines: string[]) => new TableCell({
  children: lines.map(line => paragraph(line)), verticalAlign: 'top',
  margins: { top: 0, right: 0, bottom: 0, left: 0 },
});
const labels = ['LEFT', 'RIGHT'].map(side => Array.from({ length: 4 }, (_, i) => side + '-' + i));
const fingerprint = hwpWrapFingerprint(JSON.stringify(labels));
const modes = ['observed', 'unmarked', 'changed-text', 'changed-width', 'too-tall'] as const;

async function fixture(mode: typeof modes[number]) {
  const height = mode === 'too-tall' ? 285 : 80;
  const rendered = mode === 'changed-text' ? [['LEFT-0', 'LEFT-1', 'LEFT-2 updated', 'LEFT-3'], labels[1]] : labels;
  const placement = mode === 'unmarked' ? undefined : JSON.stringify({
    height, widths: mode === 'changed-width' ? [149, 150] : [150, 150], fingerprint,
  });
  const tableHint = JSON.stringify({ row: 1, rows: 3, bodyWidth: 300, bodyHeight: 300, tableTop: 204,
    cells: ['EARLY-A', 'EARLY-B'].map(hwpWrapFingerprint) });
  const table = new PreservedTable({
    width: { size: 4500, type: WidthType.DXA }, layout: 'fixed', columnWidths: [2250, 2250],
    borders: { top: noBorder, right: noBorder, bottom: noBorder, left: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [
      new TableRow({ tableHeader: true, children: [cell(['HEAD-A']), cell(['HEAD-B'])] }),
      new TableRow({ children: [cell(['EARLY-A']), cell(['EARLY-B'])] }),
      new PreservedTableRow({ cantSplit: false, height: { value: height * 15, rule: HeightRule.ATLEAST },
        children: rendered.map(cell) }, undefined, false, undefined, false, placement),
    ],
  }, undefined, undefined, undefined, false, undefined, tableHint);
  return Packer.toBuffer(new Document({ sections: [{ properties: { page: {
    size: { width: 6000, height: 6000 }, margin: { top: 750, right: 750, bottom: 750, left: 750 },
  } }, children: [paragraph('LEAD', 205), table] }] }));
}

for (const mode of modes) {
  test(`source row placement uses the editor fit without changing DOCX split permission: ${mode}`, async ({ page }, info) => {
    let bytes = await fixture(mode);
    await page.goto('/');
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    for (let stage = 0; stage <= (mode === 'observed' ? 2 : 0); stage++) {
      const title = 'hwp-intact-top-row-' + mode + '-' + stage;
      await page.locator('input[type=file]').first().setInputFiles({
        name: title + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes,
      });
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(title);
      const first = editor.getByText('LEFT-0', { exact: true });
      await expect.poll(() => first.evaluate(p => {
        const root = p.closest('.tiptap')!, stack = root.closest<HTMLElement>('.document-stack')!;
        return 1 + Math.floor((p.getBoundingClientRect().top - root.getBoundingClientRect().top + .5) /
          (Number(stack.dataset.pageHeight) + 24));
      })).toBe(mode === 'observed' ? 2 : 1);
      if (mode !== 'observed' || stage === 2) break;
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const saved = info.outputPath(`textbox-save-hwp-intact-top-row-${stage + 1}.docx`);
      await (await pending).saveAs(saved);
      bytes = await readFile(saved);
      const zip = await JSZip.loadAsync(bytes), xml = await zip.file('word/document.xml')!.async('string');
      const row = [...xml.matchAll(/<w:tr(?:\s[^>]*?)?>[\s\S]*?<\/w:tr>/g)].find(match => match[0].includes('LEFT-0'))?.[0];
      expect(row).toContain('wsr:intactTop=');
      expect(row).toMatch(/<w:cantSplit w:val="(?:off|0|false)"\s*\/>/);
      expect(row).toMatch(/<w:trHeight w:val="1200" w:hRule="atLeast"\s*\/>/);
    }
  });
}
