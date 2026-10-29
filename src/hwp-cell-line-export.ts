import { generateHTML } from '@tiptap/core';
import type { JSONContent } from '@tiptap/core';
import { editorExtensions } from './extensions';
import { createPrintLayout } from './print-layout';
import type { StudioDocument } from './types';
import { activeHwpCellLine, hwpCellAdvance, hwpCellFont } from './hwp-cell-line';
import { tableAnchor } from './table-anchor-data';

/** Other DOCX readers do not consume our optional terminal-line metadata.
 * Serialize compact spacing only for a currently single visual line; edited
 * or narrowed multi-line cells require their full native baseline advance.
 * Measure the same table/column widths and loaded fonts used by print. */
export async function measureHwpCellExportLines(doc: StudioDocument): Promise<WeakMap<JSONContent, number>> {
  const nodes: JSONContent[] = [], tables: JSONContent[] = [], heights = new WeakMap<JSONContent, number>();
  const walk = (node: JSONContent) => {
    if (activeHwpCellLine(node.attrs)) nodes.push(node);
    if (node.type === 'table' && tableAnchor(node.attrs?.tableAnchor)?.overlay) tables.push(node);
    node.content?.forEach(walk);
  };
  walk(doc.content);
  if (!nodes.length && !tables.length) return heights;
  const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
  const paragraphs = [...root.querySelectorAll<HTMLElement>('[data-hwp-cell-line]')];
  if (paragraphs.length !== nodes.length) return heights;
  const text = paragraphs.map(p => p.textContent);
  paragraphs.forEach((p, i) => p.dataset.hwpExportLine = String(i));
  const floating = [...root.querySelectorAll<HTMLElement>('table[data-table-anchor]')].filter(t => tableAnchor(t.dataset.tableAnchor)?.overlay);
  if (floating.length === tables.length) floating.forEach((t, i) => t.dataset.hwpExportTable = String(i));
  const layout = await createPrintLayout(root.innerHTML, doc.page);
  try {
    for (const table of layout.element.querySelectorAll<HTMLElement>('table.table-overlay-active[data-hwp-export-table]')) {
      const node = tables[Number(table.dataset.hwpExportTable)];
      if (node) heights.set(node, table.getBoundingClientRect().height);
    }
    const measured = new Map<number, HTMLElement[]>();
    for (const p of layout.element.querySelectorAll<HTMLElement>('[data-hwp-export-line]')) {
      const index = Number(p.dataset.hwpExportLine), values = measured.get(index) || [];
      values.push(p); measured.set(index, values);
    }
    for (const [index, fragments] of measured) {
      const node = nodes[index], line = node && activeHwpCellLine(node.attrs);
      if (!line) continue;
      const advance = hwpCellAdvance(line, hwpCellFont(node, line));
      // Repeated headers retain their complete paragraph text; split body
      // fragments do not. A paragraph split into one line per page still
      // needs full advance in the unfragmented DOCX paragraph.
      const fragmented = fragments.some(p => p.textContent !== text[index]);
      const height = fragmented ? advance : Math.max(...fragments.map(p => p.dataset.hwpCellTerminal ? Number(p.dataset.hwpCellAdvance) : parseFloat(getComputedStyle(p).lineHeight)));
      if (Number.isFinite(height) && height > 0) heights.set(node, height);
    }
  } finally { layout.dispose(); }
  return heights;
}
