import type { EditorView } from '@tiptap/pm/view';
import { Decoration } from '@tiptap/pm/view';

import { measureParagraphLines } from './editor-lines';
type Geometry = { origin: number; scale: number; top: number; bottom: number; height: number; extra: number; removed: { top: number; height: number }[]; repeatAt?: (page: number) => number };

/** Insert view-only gaps in each column of a continuing row. The row and its
 * cells remain single editable model nodes, so selection/history/save do not
 * need to merge artificial table fragments back into the user's document. */
export function paginateCellRow(view: EditorView, row: HTMLTableRowElement, rowPos: number, geometry: Geometry) {
  if (row.querySelector('[data-note-reference],img[data-image-floating]')) return null;
  const g = geometry, stride = g.height + 24;
  const naturalY = (y: number) => {
    const value = (y - g.origin) / g.scale;
    return value - g.removed.filter(r => r.top < value - .5).reduce((sum, r) => sum + r.height, 0);
  };
  const decorations: Decoration[] = [], signatures: string[] = [];
  const gapFor = (top: number, bottom: number, offset: number, before = 0, after = 0) => {
    const y = top + g.extra + offset, page = Math.max(0, Math.floor((y - g.top + .5) / stride));
    if (bottom + after + g.extra + offset <= page * stride + g.height - g.bottom + .5) return 0;
    return Math.max(0, Math.ceil(((page + 1) * stride + g.top + (g.repeatAt?.(page + 1) ?? 0) + before - y) * 64) / 64);
  };
  const addGap = (pos: number, gap: number, columns = 0) => {
    decorations.push(Decoration.widget(pos, () => {
      const node = document.createElement(columns ? 'tr' : 'span');
      node.className = `pagination-cell-gap${columns ? ' pagination-row-gap' : ''}`;
      node.contentEditable = 'false'; node.setAttribute('aria-hidden', 'true');
      node.style.cssText = `display:${columns ? 'table-row' : 'block'};height:${gap}px;line-height:0;pointer-events:none`;
      if (columns) { const cell = document.createElement('td'); cell.colSpan = columns; cell.style.height = `${gap}px`; node.append(cell); }
      return node;
    }, { side: -1, key: `cell-gap-${pos}-${gap}`, ignoreSelection: true }));
    signatures.push(`${pos}:${gap}`);
  };
  const flowRow = (current: HTMLTableRowElement, pos: number, offset: number, outerBefore = 0, outerAfter = 0): number | null => {
    let flowedContentBottom = 0, physicalGaps = 0;
    for (const cell of [...current.cells]) {
      let extra = 0, cellGaps = 0;
      const cellStyle = getComputedStyle(cell);
      const before = outerBefore + (parseFloat(cellStyle.paddingTop) || 0) + (parseFloat(cellStyle.borderTopWidth) || 0);
      const after = outerAfter + (parseFloat(cellStyle.paddingBottom) || 0) + (parseFloat(cellStyle.borderBottomWidth) || 0);
      const visit = (block: HTMLElement): boolean => {
        if (block.matches('table')) {
          // View-only page gutters are TR widgets too. Counting them as source
          // rows changes a one-row form into a multi-row table on the next pass,
          // alternately moving it whole and splitting it on every animation frame.
          const table = block as HTMLTableElement, rows = [...table.rows].filter(row => !row.classList.contains('ProseMirror-widget'));
          // The editor keeps tables inside a wrapper; printed fragments carry
          // these margins on the table itself. Both reserve the same enclosing
          // cell padding and table spacing at a page boundary.
          const wrapper = table.parentElement?.classList.contains('tableWrapper') ? table.parentElement : table;
          const tableStyle = getComputedStyle(wrapper);
          const nestedBefore = before + (parseFloat(tableStyle.marginTop) || 0);
          const nestedAfter = after + (parseFloat(tableStyle.marginBottom) || 0);
          let headerEnd = 0;
          while (headerEnd < rows.length && rows[headerEnd].cells.length && [...rows[headerEnd].cells].every(c => c.tagName === 'TH')) headerEnd++;
          for (let i = 0; i < rows.length;) {
            let end = i === 0 ? Math.max(1, Math.min(rows.length, headerEnd + 1)) : i + 1;
            for (let j = i; j < end; j++) for (const c of rows[j].cells) end = Math.min(rows.length, Math.max(end, j + c.rowSpan));
            const first = rows[i], top = naturalY(first.getBoundingClientRect().top), bottom = naturalY(rows[end - 1].getBoundingClientRect().bottom);
            const gap = gapFor(top, bottom, offset + extra, nestedBefore, nestedAfter);
            if (gap) {
              const wholeForm = rows.length === 1 && table.getBoundingClientRect().height / g.scale <= g.height - g.top - g.bottom + .5;
              // Print can continue a single-cell row inside a multi-row form
              // at its paragraph/line boundaries. Moving the whole nested row
              // here leaves an extra blank area and shifts all later rows.
              if (!wholeForm && end === i + 1 && first.dataset.rowSplit !== 'avoid' && first.dataset.wordRowRule !== 'exact' &&
                (first.dataset.rowSplit === 'allow' || first.cells.length === 1)) {
                const more = flowRow(first, view.posAtDOM(first, 0) - 1, offset + extra, nestedBefore, nestedAfter);
                if (more === null) return false;
                extra += more; cellGaps += more;
              } else if (bottom - top <= g.height - g.top - g.bottom) {
                const page = Math.max(0, Math.floor((top + g.extra + offset + extra - g.top + .5) / stride));
                // The unused body area before a form moves consumes an outer
                // row's minimum-height reserve. Only the physical paper gutter
                // and repeated header extend that reserve onto the next sheet.
                const gutter = g.top + g.bottom + 24 + (g.repeatAt?.(page + 1) ?? 0);
                addGap(view.posAtDOM(first, 0) - 1, gap, [...first.cells].reduce((n, c) => n + c.colSpan, 0)); extra += gap; cellGaps += Math.min(gap, gutter);
              } else return false;
            }
            i = end;
          }
          return true;
        }
        if (!block.matches('p,h1,h2,h3,h4,h5,h6,pre')) return [...block.children].every(child => visit(child as HTMLElement));
        for (const line of measureParagraphLines(view, block, naturalY)) {
          if (line.bottom - line.top > g.height - g.top - g.bottom) return false;
          const gap = gapFor(line.top, line.bottom, offset + extra, before + (line.before || 0), after + (line.after || 0));
          if (gap) { addGap(line.pos(), gap); extra += gap; cellGaps += Math.max(0, gap - before - after); }
        }
        return true;
      };
      if (![...cell.children].every(block => visit(block as HTMLElement))) return null;
      const style = getComputedStyle(cell), last = cell.lastElementChild;
      const bottom = naturalY((last || cell).getBoundingClientRect().bottom) + (parseFloat(style.paddingBottom) || 0) + (parseFloat(style.borderBottomWidth) || 0);
      flowedContentBottom = Math.max(flowedContentBottom, bottom + extra);
      physicalGaps = Math.max(physicalGaps, cellGaps);
    }
    const naturalBottom = naturalY(current.getBoundingClientRect().bottom);
    // Repeated cell padding consumes existing minimum-height reserve. Only
    // space between row fragments extends that reserve; actual content can
    // still require a taller row when there is no spare height.
    let extra = Math.max(physicalGaps, flowedContentBottom - naturalBottom, 0);
    // A form may reserve blank space below its last paragraph. That space
    // continues on the next sheet too: paper margins and repeated headings
    // must not silently consume the stored minimum row height.
    let cursor = flowedContentBottom + g.extra + offset;
    let page = Math.max(0, Math.floor((cursor - g.top + .5) / stride));
    while (naturalBottom + g.extra + offset + extra > page * stride + g.height - g.bottom + .5) {
      const boundary = page * stride + g.height - g.bottom;
      if (cursor > boundary + .5) { page++; continue; }
      const next = (page + 1) * stride + g.top + (g.repeatAt?.(page + 1) ?? 0);
      extra += next - boundary;
      cursor = next; page++;
    }
    // Page gutters must not consume a form's saved minimum row height. Keep
    // that reserve, and measure again against the original height next time.
    const height = current.getBoundingClientRect().height / g.scale + extra;
    decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, {
      class: 'pagination-split-row', style: `--pagination-row-height:${height}px;--pagination-source-row-height:${current.style.height || 'auto'}`,
    }));
    signatures.push(`row:${pos}:${height}`);
    return extra;
  };
  const extra = flowRow(row, rowPos, 0);
  return extra === null ? null : { decorations, signatures, extra };
}

/** Continue source-observed vertical merges at row boundaries without inserting
 * TRs (which would change how the browser counts rowspan). Reserve the page
 * gutter inside the next row's cell padding instead. */
export function paginateMergedRows(view: EditorView, rows: HTMLTableRowElement[], geometry: Geometry) {
  const g = geometry, stride = g.height + 24, usable = g.height - g.top - g.bottom;
  const y = (value: number) => {
    const n = (value - g.origin) / g.scale;
    return n - g.removed.filter(r => r.top < n - .5).reduce((s, r) => s + r.height, 0);
  };
  const merges = rows.flatMap((row, index) => [...row.cells].filter(c => c.rowSpan > 1).map(cell => ({ cell, index })));
  if (!merges.length || rows.some(r => r.dataset.wordRowRule === 'exact' || r.dataset.rowSplit === 'avoid' || r.getBoundingClientRect().height / g.scale > usable) ||
    merges.some(({ cell }) => cell.dataset.cellSplit !== 'allow' || cell.children.length !== 1 || !cell.firstElementChild?.matches('p') || cell.querySelector('img,table') || cell.firstElementChild.getBoundingClientRect().height / g.scale > usable)) return null;
  const decorations: Decoration[] = [], signatures: string[] = [], offsets = [0];
  let extra = 0;
  for (const row of rows) {
    const box = row.getBoundingClientRect(), top = y(box.top) + g.extra + extra, height = box.height / g.scale;
    const page = Math.max(0, Math.floor((top - g.top + .5) / stride)), boundary = page * stride + g.height - g.bottom;
    if (top + height > boundary + .5) {
      const gap = Math.max(0, Math.ceil(((page + 1) * stride + g.top + (g.repeatAt?.(page + 1) ?? 0) - top) * 64) / 64);
      for (const cell of [...row.cells]) {
        const pos = view.posAtDOM(cell, 0) - 1;
        const padding = parseFloat(getComputedStyle(cell).paddingTop) || 0;
        decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, { class: 'pagination-merge-padding',
          style: `--pagination-cell-padding:${padding + gap}px;--pagination-source-cell-padding:${padding}px` }));
      }
      const pos = view.posAtDOM(row, 0) - 1;
      decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, { class: 'pagination-merge-row',
        style: `--pagination-row-height:${height + gap}px;--pagination-source-row-height:${row.style.height || 'auto'}` }));
      extra += gap; signatures.push(`merge:${pos}:${gap}`);
    }
    offsets.push(extra);
  }
  for (const { cell, index } of merges) {
    const paragraph = cell.firstElementChild as HTMLElement, box = paragraph.getBoundingClientRect();
    const alignment = getComputedStyle(cell).verticalAlign;
    const before = offsets[index], inside = offsets[Math.min(rows.length, index + cell.rowSpan)] - before;
    const firstGap = offsets[index + 1] - before;
    const displacement = before + (alignment === 'middle' ? (inside + firstGap) / 2 : alignment === 'bottom' ? inside : firstGap);
    const top = y(box.top) + g.extra + displacement, bottom = y(box.bottom) + g.extra + displacement;
    const page = Math.max(0, Math.floor((top - g.top + .5) / stride));
    if (bottom > page * stride + g.height - g.bottom + .5) {
      const shift = Math.max(0, (page + 1) * stride + g.top - top), pos = view.posAtDOM(paragraph, 0) - 1;
      if (bottom + shift > y(cell.getBoundingClientRect().bottom) + g.extra + before + inside + .5) return null;
      decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, { class: 'pagination-merge-label', style: `--pagination-label-shift:${shift}px` }));
      signatures.push(`merge-label:${pos}:${shift}`);
    }
  }
  return { decorations, signatures, extra };
}
