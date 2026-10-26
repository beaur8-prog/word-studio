import type { EditorView } from '@tiptap/pm/view';
import { Decoration } from '@tiptap/pm/view';

type Line = { top: number; bottom: number; pos: () => number };
type Geometry = { origin: number; scale: number; top: number; bottom: number; height: number; extra: number; removed: { top: number; height: number }[] };

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
  const gapFor = (top: number, bottom: number, offset: number) => {
    const y = top + g.extra + offset, page = Math.max(0, Math.floor((y - g.top + .5) / stride));
    if (bottom + g.extra + offset <= page * stride + g.height - g.bottom + .5) return 0;
    return Math.max(0, Math.ceil(((page + 1) * stride + g.top - y) * 64) / 64);
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
  const paragraphLines = (block: HTMLElement) => {
      const firstPos = view.posAtDOM(block, 0) - 1;
      const blockLines: Line[] = [];
      const add = (top: number, bottom: number, pos: () => number) => {
        const previous = blockLines.at(-1);
        if (previous && Math.min(previous.bottom, bottom) - Math.max(previous.top, top) > Math.min(previous.bottom - previous.top, bottom - top) * .4) {
          previous.top = Math.min(previous.top, top); previous.bottom = Math.max(previous.bottom, bottom); return;
        }
        blockLines.push({ top, bottom, pos });
      };
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      while (walker.nextNode()) {
        const node = walker.currentNode;
        if (node instanceof HTMLElement) {
          if (node.tagName !== 'IMG') continue;
          const box = node.getBoundingClientRect();
          add(naturalY(box.top), naturalY(box.bottom), () => view.posAtDOM(node.parentNode!, [...node.parentNode!.childNodes].indexOf(node)));
        } else if (node instanceof Text && node.data && !node.parentElement?.closest('.pagination-gap,[data-flow-widget]')) {
          const range = document.createRange(); range.selectNodeContents(node);
          for (const box of range.getClientRects()) {
            if (box.width < .1 || box.height < 1) continue;
            const top = naturalY(box.top), bottom = naturalY(box.bottom), target = box.top;
            add(top, bottom, () => {
              // Find the first character on this visual line only if a page
              // gap is needed. Keep surrogate pairs and combining clusters.
              const offsets = [...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(node.data)].map(s => s.index);
              let low = 0, high = offsets.length - 1, first = 0;
              while (low <= high) {
                const mid = (low + high) >> 1, end = offsets[mid + 1] ?? node.length;
                range.setStart(node, offsets[mid]); range.setEnd(node, end);
                if (range.getBoundingClientRect().top < target - .5) { first = mid + 1; low = mid + 1; } else high = mid - 1;
              }
              return view.posAtDOM(node, offsets[Math.min(first, offsets.length - 1)]);
            });
          }
        }
      }
      if (!blockLines.length) { const box = block.getBoundingClientRect(); blockLines.push({ top: naturalY(box.top), bottom: naturalY(box.bottom), pos: () => firstPos }); }
      else blockLines[0].pos = () => firstPos;
      return blockLines;
  };
  const flowRow = (current: HTMLTableRowElement, pos: number, offset: number): number | null => {
    let naturalContentBottom = 0, flowedContentBottom = 0;
    for (const cell of [...current.cells]) {
      let extra = 0;
      const visit = (block: HTMLElement): boolean => {
        if (block.matches('table')) {
          const table = block as HTMLTableElement, rows = [...table.rows];
          let headerEnd = 0;
          while (headerEnd < rows.length && rows[headerEnd].cells.length && [...rows[headerEnd].cells].every(c => c.tagName === 'TH')) headerEnd++;
          for (let i = 0; i < rows.length;) {
            let end = i === 0 ? Math.max(1, Math.min(rows.length, headerEnd + 1)) : i + 1;
            for (let j = i; j < end; j++) for (const c of rows[j].cells) end = Math.min(rows.length, Math.max(end, j + c.rowSpan));
            const first = rows[i], top = naturalY(first.getBoundingClientRect().top), bottom = naturalY(rows[end - 1].getBoundingClientRect().bottom);
            const gap = gapFor(top, bottom, offset + extra);
            if (gap) {
              if (bottom - top <= g.height - g.top - g.bottom) {
                addGap(view.posAtDOM(first, 0) - 1, gap, [...first.cells].reduce((n, c) => n + c.colSpan, 0)); extra += gap;
              } else if (end === i + 1 && first.dataset.rowSplit === 'allow' && first.dataset.wordRowRule !== 'exact') {
                const more = flowRow(first, view.posAtDOM(first, 0) - 1, offset + extra);
                if (more === null) return false;
                extra += more;
              } else return false;
            }
            i = end;
          }
          return true;
        }
        if (!block.matches('p,h1,h2,h3,h4,h5,h6,pre')) return [...block.children].every(child => visit(child as HTMLElement));
        for (const line of paragraphLines(block)) {
          if (line.bottom - line.top > g.height - g.top - g.bottom) return false;
          const gap = gapFor(line.top, line.bottom, offset + extra);
          if (gap) { addGap(line.pos(), gap); extra += gap; }
        }
        return true;
      };
      if (![...cell.children].every(block => visit(block as HTMLElement))) return null;
      const style = getComputedStyle(cell), last = cell.lastElementChild;
      const bottom = naturalY((last || cell).getBoundingClientRect().bottom) + (parseFloat(style.paddingBottom) || 0) + (parseFloat(style.borderBottomWidth) || 0);
      naturalContentBottom = Math.max(naturalContentBottom, bottom);
      flowedContentBottom = Math.max(flowedContentBottom, bottom + extra);
    }
    const extra = Math.max(0, flowedContentBottom - naturalContentBottom);
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
      const gap = Math.max(0, Math.ceil(((page + 1) * stride + g.top - top) * 64) / 64);
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
