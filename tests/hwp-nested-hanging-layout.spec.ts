import { expect, test } from '@playwright/test';
import { BorderStyle, Document, LineRuleType, Packer, Paragraph, Table, TableCell, TableLayoutType, TableRow, TextRun, WidthType } from 'docx';
import JSZip from 'jszip';
import { HwpCellParagraph } from '../src/docx-hwp-cell-line';
import { hwpWrapFingerprint, normalizeHwpWrapText } from '../src/hwp-text-advances';

const literal = 'Intro words   and the remaining text';
const ends = [10, normalizeHwpWrapText(literal).length];
const border = { style: BorderStyle.NONE } as const;
const margins = { top: 0, bottom: 0, left: 0, right: 0 };

const table = (children: (Paragraph | Table)[], width: number) => new Table({
  width: { size: width * 15, type: WidthType.DXA },
  columnWidths: [width * 15],
  layout: TableLayoutType.FIXED,
  borders: { top: border, right: border, bottom: border, left: border, insideHorizontal: border, insideVertical: border },
  rows: [new TableRow({ children: [new TableCell({ children, margins })] })],
});

async function fixture(mode: 'original' | 'text' | 'width' | 'scale' | 'top-level') {
  const witness = JSON.stringify({
    fingerprint: hwpWrapFingerprint(literal),
    ends, cellWidth: { outer: 190, text: 189 },
    nestedHanging: { hanging: 16, scale: 60, size: 16 },
  });
  const para = new HwpCellParagraph({
    children: [new TextRun({ text: mode === 'text' ? literal + '!' : literal,
      font: 'Courier New', size: 24, scale: mode === 'scale' ? 68 : 60 })],
    indent: { left: 240, hanging: 240 }, spacing: { before: 0, after: 0, line: 300, lineRule: LineRuleType.EXACT },
  }, undefined, undefined, undefined, undefined, undefined, false, witness);
  const inner = table([para, new Paragraph('cell tail')], mode === 'width' ? 194 : 190);
  const body = mode === 'top-level' ? inner : table([inner, new Paragraph('outer tail')], 200);
  return Packer.toBuffer(new Document({ sections: [{ children: [body] }] }));
}

for (const mode of ['original', 'text', 'width', 'scale', 'top-level'] as const) {
  test(`use verified nested source lines only in a matching print copy (${mode})`, async ({ page }) => {
    const bytes = await fixture(mode);
    await page.goto('/');
    await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
    const result = await page.evaluate(async ({ bytes, literal, stageCount }) => {
      const { importFile, exportDocx } = await import('/src/document-io.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { generateHTML } = await import('/node_modules/@tiptap/core/dist/index.js');
      const { wordLayoutOptions } = await import('/src/word-trailing-spaces.ts');
      const norm = (t: string) => t.replace(/\s/g, '');
      let current = new Blob([new Uint8Array(bytes)]);
      const stages = [], saves: number[][] = [];
      for (let stage = 0; stage < stageCount; stage++) {
        const loaded = await importFile(new File([current], 'nested.docx'));
        const doc = { ...loaded.document, content: parseDocumentHtml(loaded.html) };
        const html = generateHTML(doc.content, editorExtensions);
        const result = await createPrintLayout(html, doc.page, undefined, undefined, wordLayoutOptions(doc));
        try {
          const p = [...result.element.querySelectorAll<HTMLParagraphElement>('.print-content td > p')]
            .find(p => p.textContent?.startsWith(literal));
          if (!p) throw Error('Printed text missing');
          const cell = p.parentElement!.getBoundingClientRect(), range = document.createRange();
          range.selectNodeContents(p);
          const boxes = [...range.getClientRects()].filter(r => r.width > 0 && r.height > 0);
          stages.push({ literal: p.textContent, breaks: p.querySelectorAll('br').length,
            lines: [...p.children].filter(child => child.tagName !== 'BR').map(child => norm(child.textContent ?? '').length),
            transform: getComputedStyle(p).transform,
            contained: boxes.every(r => r.left >= cell.left && r.right <= cell.right) });
        } finally { result.dispose(); }
        if (stage + 1 < stageCount) {
          current = await exportDocx(doc);
          saves.push([...new Uint8Array(await current.arrayBuffer())]);
        }
      }
      return { stages, saves };
    }, { bytes: [...bytes], literal, stageCount: mode === 'original' ? 3 : 1 });
    for (const stage of result.stages) {
      expect(stage.literal).toBe(mode === 'text' ? literal + '!' : literal);
      expect(stage.breaks).toBe(mode === 'original' ? 1 : 0);
      if (mode === 'original') {
        expect(stage.transform).toContain('0.6');
        expect(stage.contained).toBe(true);
        expect(stage.lines).toEqual([ends[0], ends[1] - ends[0]]);
      } else {
        expect(stage.transform).toBe('none');
      }
    }
    for (const data of result.saves) {
      const zip = await JSZip.loadAsync(new Uint8Array(data));
      const xml = await zip.file('word/document.xml')!.async('string');
      expect(xml).toContain('wsa:wrap=');
      expect(xml).not.toMatch(/<w:br(?:\s[^>]*)?\s*\/>/);
      expect(xml).toContain('w:w w:val="60"');
    }
  });
}
