import type { EditorView } from '@tiptap/pm/view';
import { Decoration } from '@tiptap/pm/view';

import { measureParagraphLines } from './editor-lines';
type Reference = { id: string; page: number };
type Geometry = { origin: number; scale: number; top: number; bottom: number; height: number; extra: number; removed: { top: number; height: number }[]; repeatAt?: (page: number) => number; references?: { id: string; top: number }[]; reserveAt?: (page: number, own?: Reference[]) => number };

/** Insert view-only gaps in each column of a continuing row. The row and its
 * cells remain single editable model nodes, so selection/history/save do not
 * need to merge artificial table fragments back into the user's document. */
export function paginateCellRow(view: EditorView, row: HTMLTableRowElement, rowPos: number, geometry: Geometry) {
  if (row.querySelector('img[data-image-floating]')) return null;
  const g = geometry, stride = g.height + 24;
  const naturalY = (y: number) => {
    const value = (y - g.origin) / g.scale;
    return value - g.removed.filter(r => r.top < value - .5).reduce((sum, r) => sum + r.height, 0);
  };
  const attempt = (assumed: Reference[], minimum: Map<string, number>) => {
    const placed = new Map(assumed.map(ref => [ref.id, ref.page]));
    const references = () => [...placed].map(([id, page]) => ({ id, page }));
    const reserve = (page: number) => g.reserveAt?.(page, references()) ?? 0;
    const decorations: Decoration[] = [], signatures: string[] = [];
    const ownIn = (element: HTMLElement, top: number, bottom: number) => (g.references ?? []).filter(ref =>
      ref.top >= top - .5 && ref.top < bottom - .5 && [...element.querySelectorAll<HTMLElement>('[data-note-reference]')].some(el => el.dataset.noteReference === ref.id));
    const gapFor = (top: number, bottom: number, offset: number, before = 0, after = 0, own: { id: string; top: number }[] = []) => {
      const y = top + g.extra + offset;
      const page = Math.max(0, Math.floor((y - g.top + .5) / stride));
      let destination = Math.max(page, ...own.map(ref => minimum.get(ref.id) ?? page));
      for (const ref of own) placed.set(ref.id, destination);
      if (destination === page && bottom + after + g.extra + offset <= page * stride + g.height - g.bottom - reserve(page) + .5) return 0;
      destination = Math.max(page + 1, destination);
      for (const ref of own) placed.set(ref.id, destination);
      return Math.max(0, Math.ceil((destination * stride + g.top + (g.repeatAt?.(destination) ?? 0) + before - y) * 64) / 64);
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
              const gap = gapFor(top, bottom, offset + extra, nestedBefore, nestedAfter, ownIn(table, top, bottom));
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
                  const gutter = g.top + g.bottom + 24 + reserve(page) + (g.repeatAt?.(page + 1) ?? 0);
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
            const gap = gapFor(line.top, line.bottom, offset + extra, before + (line.before || 0), after + (line.after || 0), ownIn(block, line.top, line.bottom));
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
      while (naturalBottom + g.extra + offset + extra > page * stride + g.height - g.bottom - reserve(page) + .5) {
        const boundary = page * stride + g.height - g.bottom - reserve(page);
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
    return extra === null ? null : { decorations, signatures, extra, references: references() };
  };
  let assumed = (g.references ?? []).map(ref => ({ id: ref.id, page: Math.max(0, Math.floor((ref.top + g.extra - g.top + .5) / stride)) }));
  const minimum = new Map<string, number>();
  // A note in the right column also reserves space in the left column. Reflow
  // until every column uses the same final reference pages. Once a reference
  // moves, don't pull it back into the page whose note forced it forward.
  for (let pass = 0; pass < Math.min(64, Math.max(2, assumed.length * 2 + 4)); pass++) {
    const flow = attempt(assumed, minimum); if (!flow) return null;
    if (flow.references.every(ref => assumed.some(old => old.id === ref.id && old.page === ref.page))) return flow;
    assumed = flow.references;
    for (const ref of assumed) minimum.set(ref.id, Math.max(minimum.get(ref.id) ?? 0, ref.page));
  }
  return null;
}

/** Continue source-observed vertical merges at row boundaries without inserting
 * TRs (which would change how the browser counts rowspan). Reserve the page
 * gutter inside the next row's cell padding instead. */
export function paginateMergedRows(view: EditorView, rows: HTMLTableRowElement[], geometry: Geometry): ReturnType<typeof mergedRowFlow> {
  return mergedRowFlow(view, rows, geometry, [], 0);
}
function mergedRowFlow(view: EditorView, rows: HTMLTableRowElement[], geometry: Geometry, growth: number[], attempt: number): { decorations: Decoration[]; signatures: string[]; extra: number } | null {
  const g = geometry, stride = g.height + 24, usable = g.height - g.top - g.bottom;
  const y = (value: number) => {
    const n = (value - g.origin) / g.scale;
    return n - g.removed.filter(r => r.top < n - .5).reduce((s, r) => s + r.height, 0);
  };
  const merges = rows.flatMap((row, index) => [...row.cells].filter(c => c.rowSpan > 1).map(cell => ({ cell, index })));
  const atomicTable = (block: Element) => block.matches('table,.tableWrapper');
  const blockPosition = (block: HTMLElement) => atomicTable(block)
    ? view.posAtDOM(block.parentNode!, [...block.parentNode!.childNodes].indexOf(block)) : view.posAtDOM(block, 0) - 1;
  if (!merges.length || rows.some(r => r.dataset.wordRowRule === 'exact' || r.dataset.rowSplit === 'avoid' || r.getBoundingClientRect().height / g.scale > usable) ||
    merges.some(({ cell }) => cell.dataset.cellSplit !== 'allow' || [...cell.children].some(block => !block.matches('p,h1,h2,h3,h4,h5,h6') && !atomicTable(block)) || cell.querySelector('img[data-image-floating],[data-note-reference]'))) return null;
  const decorations: Decoration[] = [], signatures: string[] = [], offsets = [0], topOffsets: number[] = [], pageGaps: number[] = [];
  const collapsed = rows[0].closest('table') && getComputedStyle(rows[0].closest('table')!).borderCollapse === 'collapse';
  const borderFactor = collapsed ? .5 : 1;
  let extra = 0;
  for (const [index, row] of rows.entries()) {
    const added = growth[index] || 0;
    const box = row.getBoundingClientRect(), top = y(box.top) + g.extra + extra, height = box.height / g.scale + added;
    if (height > usable + .5) return null;
    const page = Math.max(0, Math.floor((top - g.top + .5) / stride)), boundary = page * stride + g.height - g.bottom - (g.reserveAt?.(page) ?? 0);
    let gap = 0;
    if (top + height > boundary + .5) {
      const edge = collapsed ? Math.max(0, ...[...row.cells].map(cell => (parseFloat(getComputedStyle(cell).borderTopWidth) || 0) / 2)) : 0;
      gap = Math.max(0, Math.ceil(((page + 1) * stride + g.top + (g.repeatAt?.(page + 1) ?? 0) + edge - top) * 64) / 64);
      for (const cell of [...row.cells]) {
        const pos = view.posAtDOM(cell, 0) - 1;
        const padding = parseFloat(getComputedStyle(cell).paddingTop) || 0;
        decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, { class: 'pagination-merge-padding',
          style: `--pagination-cell-padding:${padding + gap}px;--pagination-source-cell-padding:${padding}px` }));
      }
      extra += gap;
    }
    if (gap || added) {
      const pos = view.posAtDOM(row, 0) - 1;
      decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, { class: 'pagination-merge-row',
        style: `--pagination-row-height:${height + gap}px;--pagination-source-row-height:${row.style.height || 'auto'}` }));
      signatures.push(`merge:${pos}:${gap}:${added}`);
    }
    topOffsets.push(extra); pageGaps.push(gap); extra += added;
    offsets.push(extra);
  }
  for (const { cell, index } of merges) {
    const end = Math.min(rows.length, index + cell.rowSpan), style = getComputedStyle(cell);
    const factor = style.verticalAlign === 'middle' ? .5 : style.verticalAlign === 'bottom' ? 1 : 0;
    const paddingTop = (parseFloat(style.paddingTop) || 0) + (parseFloat(style.borderTopWidth) || 0) * borderFactor;
    const paddingBottom = (parseFloat(style.paddingBottom) || 0) + (parseFloat(style.borderBottomWidth) || 0) * borderFactor;
    const blocks = [...cell.children] as HTMLElement[];
    const lines = blocks.flatMap(block => {
      const atomic = atomicTable(block), box = block.getBoundingClientRect();
      // Nested tables wholly inside one row fragment remain editable atomic
      // blocks. Their internal cell traversal is not a sequence of text lines.
      const measured = atomic ? [{ top: y(box.top), bottom: y(box.bottom), before: 0, after: 0, pos: () => blockPosition(block) }]
        : measureParagraphLines(view, block, y), css = getComputedStyle(block);
      return measured.map((line, n) => ({ ...line, block, atomic,
        marginTop: n === 0 ? parseFloat(css.marginTop) || 0 : 0,
        marginBottom: n === measured.length - 1 ? parseFloat(css.marginBottom) || 0 : 0,
      }));
    });
    if (!lines.length) continue;
    const firstTop = y(blocks[0].getBoundingClientRect().top) - lines[0].marginTop;
    const lastBottom = y(blocks.at(-1)!.getBoundingClientRect().bottom) + lines.at(-1)!.marginBottom;
    const cuts = [index, ...rows.map((_, i) => i).filter(i => i > index && i < end && pageGaps[i] > .5), end];
    const addedHeight = (from: number, to: number) => growth.slice(from, to).reduce((sum, value) => sum + (value || 0), 0);
    const firstHeight = y(rows[cuts[1] - 1].getBoundingClientRect().bottom) - y(rows[index].getBoundingClientRect().top) + addedHeight(index, cuts[1]);
    const rich = blocks.length > 1 || lastBottom - firstTop > firstHeight - paddingTop - paddingBottom;
    const targets: number[] = [];
    let start = 0, finalNeeded = 0, finalAvailable = 0;
    for (let part = 0; part < cuts.length - 1 && start < lines.length; part++) {
      const from = cuts[part], to = cuts[part + 1];
      const top = y(rows[from].getBoundingClientRect().top) + g.extra + topOffsets[from];
      const bottom = y(rows[to - 1].getBoundingClientRect().bottom) + g.extra + offsets[to];
      const naturalHeight = y(rows[end - 1].getBoundingClientRect().bottom) - y(rows[from].getBoundingClientRect().top) + addedHeight(from, end);
      const sourceTop = start ? lines[start].top - (lines[start].before || 0) - lines[start].marginTop : firstTop;
      const remainingHeight = lastBottom - sourceTop;
      if (part === cuts.length - 2) { finalNeeded = remainingHeight + paddingTop + paddingBottom; finalAvailable = bottom - top; }
      const beforeSplit = top + paddingTop + (rich ? 0 : factor) * Math.max(0, naturalHeight - paddingTop - paddingBottom - remainingHeight);
      // Font range boxes may extend beyond an exact final line's flow box.
      // Use its real paragraph end when budgeting the last line's advance.
      const lineEnd = (n: number) => n === lines.length - 1 ? lastBottom : lines[n].bottom + (lines[n].after || 0) + lines[n].marginBottom;
      let stop = start;
      while (stop < lines.length && beforeSplit + lineEnd(stop) - sourceTop <= bottom - paddingBottom + .5) stop++;
      // Splitting the nested table itself requires its own row pagination.
      // Retain the conservative fallback when that would be necessary.
      if (stop < lines.length && lines[stop].atomic) return null;
      if (stop === start) continue;
      const keptHeight = lineEnd(stop - 1) - sourceTop;
      const placedTop = top + paddingTop + factor * Math.max(0, bottom - top - paddingTop - paddingBottom - keptHeight);
      for (let n = start; n < stop; n++) targets[n] = placedTop + lines[n].top - sourceTop;
      start = stop;
    }
    if (start < lines.length) {
      // A continuation repeats cell padding and cannot reuse the unused tail
      // of the preceding fragment. Print grows its final row to fit. Plan the
      // same view-only growth here, then recalculate all crossing merges and
      // page cuts together; saved minimum heights and undo stay unchanged.
      if (attempt >= 8 || finalNeeded <= finalAvailable + .5) return null;
      const next = [...growth];
      next[end - 1] = (next[end - 1] || 0) + Math.ceil((finalNeeded - finalAvailable) * 64) / 64;
      return mergedRowFlow(view, rows, geometry, next, attempt + 1);
    }
    const cellPos = view.posAtDOM(cell, 0) - 1;
    decorations.push(Decoration.node(cellPos, cellPos + view.state.doc.nodeAt(cellPos)!.nodeSize, {
      class: 'pagination-merge-flow', style: `--pagination-source-valign:${style.verticalAlign}`,
    }));
    let inserted = 0;
    const base = y(cell.getBoundingClientRect().top) + g.extra + topOffsets[index] + paddingTop - firstTop;
    for (let n = 0; n < lines.length;) {
      const first = n, block = lines[n].block;
      const shift = targets[n] - (lines[n].top + base + inserted);
      while (++n < lines.length && lines[n].block === block) {
        const gap = targets[n] - (lines[n].top + base + inserted + shift);
        if (gap > .5) {
          const pos = lines[n].pos();
          decorations.push(Decoration.widget(pos, () => {
            const node = document.createElement('span'); node.className = 'pagination-cell-gap';
            node.contentEditable = 'false'; node.setAttribute('aria-hidden', 'true');
            node.style.cssText = `display:block;height:${gap}px;line-height:0;pointer-events:none`;
            return node;
          }, { side: -1, key: `merge-line-${pos}-${gap}`, ignoreSelection: true }));
          inserted += gap; signatures.push(`merge-line:${pos}:${gap}`);
        }
      }
      const pos = blockPosition(block);
      decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, {
        class: 'pagination-merge-label', style: `--pagination-label-shift:${shift}px`,
      }));
      signatures.push(`merge-label:${pos}:${shift}:${first}`);
    }
  }
  return { decorations, signatures, extra };
}
