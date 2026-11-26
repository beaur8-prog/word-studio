import { withoutEditorReflow } from './editor-layout-task';
import { generateHTML } from '@tiptap/core';
import type { JSONContent } from '@tiptap/core';
import { editorExtensions } from './extensions';
import { createPrintLayout } from './print-layout';
import type { StudioDocument } from './types';
import { activeHwpCellLine, hwpCellAdvance, hwpCellFont } from './hwp-cell-line';
import { cellTablePosition } from './cell-table-position';
import { tableAnchor } from './table-anchor-data';
import { hwpCellRasterFillAttribute, parseHwpCellRasterFill } from './hwp-cell-raster-fill';

/** Measure export geometry against print's current table/column widths and
 * loaded fonts. A wrapped cell uses its full native advance. A nested table
 * pair uses standard cell-relative anchors only while its rendered fit holds. */
export async function measureHwpExportLayout(doc: StudioDocument) {
  const nodes: JSONContent[] = [], tables: JSONContent[] = [], pairs: JSONContent[] = [], filledCells: JSONContent[] = [];
  const heights = new WeakMap<JSONContent, number>(), positionedPairs = new WeakSet<JSONContent>();
  const rasterCells = new WeakMap<JSONContent, { width: number; height: number; paragraphOffset: number }>();
  const result = { heights, positionedPairs, rasterCells };
  const walk = (node: JSONContent) => {
    if (activeHwpCellLine(node.attrs)) nodes.push(node);
    if (node.type === 'table' && tableAnchor(node.attrs?.tableAnchor)?.overlay) tables.push(node);
    if (node.type === 'table' && cellTablePosition(node.attrs?.cellTablePosition)) pairs.push(node);
    if (node.type === 'tableCell' && parseHwpCellRasterFill(node.attrs?.hwpCellRasterFill)) filledCells.push(node);
    node.content?.forEach(walk);
  };
  walk(doc.content);
  if (!nodes.length && !tables.length && !pairs.length && !filledCells.length) return result;
  const root = document.createElement('div'); root.innerHTML = generateHTML(doc.content, editorExtensions);
  const paragraphs = [...root.querySelectorAll<HTMLElement>('[data-hwp-cell-line]')];
  if (paragraphs.length !== nodes.length) return result;
  const text = paragraphs.map(p => p.textContent);
  paragraphs.forEach((p, i) => p.dataset.hwpExportLine = String(i));
  const floating = [...root.querySelectorAll<HTMLElement>('table[data-table-anchor]')].filter(t => tableAnchor(t.dataset.tableAnchor)?.overlay);
  if (floating.length === tables.length) floating.forEach((t, i) => t.dataset.hwpExportTable = String(i));
  const nested = [...root.querySelectorAll<HTMLElement>('table[data-cell-table-position]')];
  const pairedTables = new Map<Element, number[]>();
  const fills = [...root.querySelectorAll<HTMLElement>(`td[${hwpCellRasterFillAttribute}]`)];
  if (fills.length === filledCells.length) fills.forEach((cell, index) => { cell.dataset.hwpExportRasterCell = String(index); });
  if (nested.length === pairs.length) nested.forEach((table, index) => {
    const owner = table.parentElement?.closest('td');
    if (!owner) return;
    const indices = pairedTables.get(owner) ?? [];
    indices.push(index);
    pairedTables.set(owner, indices);
    table.dataset.hwpExportCellPair = String(index);
  });
  return withoutEditorReflow(async () => {
    const layout = await createPrintLayout(root.innerHTML, doc.page, undefined, undefined, { paintHwpCellLines: false });
    try {
      const fragments = new Map<number, HTMLElement[]>();
      for (const cell of layout.element.querySelectorAll<HTMLElement>('[data-hwp-export-raster-cell]')) {
        const index = Number(cell.dataset.hwpExportRasterCell), entries = fragments.get(index) ?? [];
        entries.push(cell);
        fragments.set(index, entries);
      }
      for (const [index, entries] of fragments) {
        const node = filledCells[index], paragraph = entries[0].firstElementChild;
        if (!node || entries.length !== 1 || !paragraph?.matches('p,h1,h2,h3,h4,h5,h6')) continue;
        const { width, height, top } = entries[0].getBoundingClientRect();
        const paragraphOffset = Math.max(0, paragraph.getBoundingClientRect().top - top);
        if (width >= 1 && height >= 1 && width < 10000 && height < 10000 && paragraphOffset <= height)
          rasterCells.set(node, { width, height, paragraphOffset });
      }
      for (const table of layout.element.querySelectorAll<HTMLElement>('table.table-overlay-active[data-hwp-export-table]')) {
        const node = tables[Number(table.dataset.hwpExportTable)];
        if (node) heights.set(node, table.getBoundingClientRect().height);
      }
      const pairFit = new Map<number, boolean>();
      for (const table of layout.element.querySelectorAll<HTMLElement>('table[data-hwp-export-cell-pair]')) {
        const index = Number(table.dataset.hwpExportCellPair);
        const active = table.parentElement?.classList.contains('cell-table-frame') ?? false;
        pairFit.set(index, active && pairFit.get(index) !== false);
      }
      // A rejected fragment of either table retires the pair together.
      // The standard DOCX must not float the other table on its own.
      for (const pair of pairedTables.values()) {
        if (pair.length === 2 && pair.every(index => pairFit.get(index)))
          for (const index of pair) positionedPairs.add(pairs[index]);
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
    return result;
  });
}
