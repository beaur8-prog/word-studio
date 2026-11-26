import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { BorderStyle, Document, LineRuleType, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx';
import JSZip from 'jszip';

const namespace = 'urn:wordstudio:hwp-grid-row-height:1';
const compatibility = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const noBorder = { style: BorderStyle.NONE, size: 0 } as const;
const width = { size: 4500, type: WidthType.DXA } as const;
const modes = ['observed', 'unmarked', 'top-aligned', 'oversized'] as const;

async function fixture(mode: typeof modes[number]) {
  const p = (text: string, line = 20) => new Paragraph({
    children: [new TextRun({ text, font: 'Arial', size: 18 })],
    spacing: { before: 0, after: 0, line: line * 15, lineRule: LineRuleType.EXACT },
  });
  const cell = (children: Paragraph[]) => new TableCell({
    children, verticalAlign: mode === 'top-aligned' ? 'top' : 'center',
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
  });
  const count = mode === 'oversized' ? 21 : 4;
  const table = new Table({
    width, layout: 'fixed', columnWidths: [2250, 2250],
    borders: { top: noBorder, right: noBorder, bottom: noBorder, left: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [
      new TableRow({ tableHeader: true, children: [cell([p('HEAD-A')]), cell([p('HEAD-B')])] }),
      new TableRow({ children: [cell([p('EARLY-A')]), cell([p('EARLY-B')])] }),
      new TableRow({ cantSplit: false, children: [cell(Array.from({ length: count }, (_, i) => p('LEFT-' + i))),
        cell(Array.from({ length: count }, (_, i) => p('RIGHT-' + i)))] }),
    ],
  });
  const source = await Packer.toBuffer(new Document({ sections: [{ properties: { page: {
    size: { width: 6000, height: 6000 }, margin: { top: 750, right: 750, bottom: 750, left: 750 },
  } }, children: [p('LEAD', 205), table] }] }));
  if (mode === 'unmarked') return source;
  const zip = await JSZip.loadAsync(source), xml = await zip.file('word/document.xml')!.async('string');
  const marker = `<w:tr xmlns:wsr="${namespace}" xmlns:mc="${compatibility}" mc:Ignorable="wsr" mc:PreserveAttributes="wsr:intact" wsr:intact="1">`;
  const last = xml.lastIndexOf('<w:tr>');
  zip.file('word/document.xml', xml.slice(0, last) + marker + xml.slice(last + '<w:tr>'.length));
  return zip.generateAsync({ type: 'nodebuffer' });
}

for (const mode of modes) {
  test(`a centered row uses the observed HWP placement only when it fits fresh paper: ${mode}`, async ({ page }, info) => {
    let bytes = await fixture(mode);
    const count = mode === 'oversized' ? 21 : 4;
    const tokens = ['LEFT-', 'RIGHT-'].flatMap(prefix => Array.from({ length: count }, (_, i) => prefix + i));
    const expectedFirst = mode === 'observed' ? 2 : 1;
    await page.goto('/');
    for (let stage = 0; stage <= (mode === 'observed' ? 2 : 0); stage++) {
      const title = 'hwp-centered-row-' + mode + '-' + stage;
      await page.locator('input[type=file]').first().setInputFiles({
        name: title + '.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: bytes,
      });
      await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(title);
      const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
      await expect(editor.locator('td p').filter({ hasText: /^(?:LEFT-|RIGHT-)/ })).toHaveCount(tokens.length);
      const editorRow = editor.locator('tr[data-row-split="allow"]').filter({ hasText: 'LEFT-0' });
      if (mode === 'unmarked') await expect(editor.locator('tr[data-hwp-row-intact]').filter({ hasText: 'LEFT-0' })).toHaveCount(0);
      else await expect(editorRow).toHaveAttribute('data-hwp-row-intact', 'true');
      const first = editor.getByText('LEFT-0', { exact: true });
      await expect.poll(() => first.evaluate(p => {
        const editor = p.closest('.tiptap')!, stack = editor.closest<HTMLElement>('.document-stack')!;
        return 1 + Math.floor((p.getBoundingClientRect().top - editor.getBoundingClientRect().top + .5) /
          (Number(stack.dataset.pageHeight) + 24));
      })).toBe(expectedFirst);

      await page.getByRole('button', { name: '보기', exact: true }).click();
      await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
      const preview = page.locator('.print-layout');
      await expect(preview.locator('.print-page')).toHaveCount(mode === 'oversized' ? 3 : 2);
      const placements = await preview.locator('.print-content td > p').evaluateAll(paragraphs => paragraphs
        .filter(p => /^(?:LEFT-|RIGHT-)/.test(p.textContent || ''))
        .map(p => {
          const ink = document.createRange();
          ink.selectNodeContents(p);
          const box = ink.getBoundingClientRect(), paper = p.closest('.print-page')!.getBoundingClientRect();
          return { text: p.textContent, page: Number(p.closest<HTMLElement>('.print-page')!.dataset.page),
            physical: box.width > 0 && box.height > 0 && box.top >= paper.top - .5 && box.bottom <= paper.bottom + .5 };
        }));
      await info.attach(`stage-${stage}-row-placements`, { body: JSON.stringify(placements), contentType: 'application/json' });
      expect(placements.map(item => item.text).sort()).toEqual([...tokens].sort());
      expect(placements.every(item => item.physical)).toBe(true);
      expect(placements.find(item => item.text === 'LEFT-0')?.page).toBe(expectedFirst);
      if (mode === 'observed') expect(new Set(placements.map(item => item.page))).toEqual(new Set([2]));
      await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
      if (mode !== 'observed' || stage === 2) break;
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const saved = info.outputPath(`textbox-save-hwp-centered-row-${stage + 1}.docx`);
      await (await pending).saveAs(saved);
      bytes = await readFile(saved);
      const zip = await JSZip.loadAsync(bytes), xml = await zip.file('word/document.xml')!.async('string');
      const row = [...xml.matchAll(/<w:tr(?:\s[^>]*?)?>[\s\S]*?<\/w:tr>/g)].find(match => match[0].includes('LEFT-0'))?.[0];
      expect(row).toMatch(/wsr:intact="1"/);
      expect(row).toMatch(/<w:cantSplit w:val="(?:off|0|false)"\s*\/>/);
    }
  });
}
