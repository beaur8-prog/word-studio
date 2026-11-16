import type { EditorView } from '@tiptap/pm/view';
import { Decoration } from '@tiptap/pm/view';
import { measureParagraphLines } from './editor-lines';
import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { layoutWordLines } from './word-line-metrics';
import { centeredContinuationParagraph, hasCenteredOrdinaryCell, hasHwpCellLines, measureHwpTerminalParagraph, precedingSourceRow, sameHwpTerminalParagraph, tableFragmenter } from './table-fragments';
import { paragraphLineOffsets, prepareParagraphEnds } from './paragraph-fragments';

function measureOrdinaryParagraph(paragraph: HTMLParagraphElement) {
  const rect = paragraph.getBoundingClientRect(), range = document.createRange();
  range.selectNodeContents(paragraph);
  const boxes = [...range.getClientRects()].filter(ink => ink.width > .1 && ink.height > .1).map(ink =>
    [ink.left - rect.left, ink.top - rect.top, ink.width, ink.height]);
  return { rect, boxes, offsets: paragraphLineOffsets(paragraph) };
}

type Geometry = { origin: number; scale: number; top: number; bottom: number; height: number; extra: number;
  removed: { top: number; height: number }[]; repeatAt?: (page: number) => number; measureRepeatAt?: (page: number) => number; reserveAt?: (page: number) => number; continuationHeaders?: HTMLTableRowElement[]; longProtectedGroup?: boolean };

/** Plain continuing merges can use the actual print-fragment geometry. More
 * complex objects keep the existing general flow until they can be mapped. */
export function canMeasureMergedFragments(rows: HTMLTableRowElement[]) {
  const merges = rows.flatMap(row => [...row.cells].filter(cell => cell.rowSpan > 1));
  return merges.length > 0 && merges.every(cell => cell.dataset.cellSplit === 'allow' && cell.children.length > 0 &&
    [...cell.children].every(p => p.matches('p,h1,h2,h3,h4,h5,h6') &&
      !p.querySelector('img:not(.ProseMirror-separator),svg,table,[data-page-number],[data-note-reference],[data-docx-tab],[data-docx-bookmark],[data-word-overlap]')));
}

/** A browser redistributes minimum row heights when a vertical merge is split.
 * Measure those same print fragments instead of budgeting every continuation
 * against the unsplit table's original row distribution. Only the view changes. */
export function measuredMergedFragments(view: EditorView, rows: HTMLTableRowElement[], g: Geometry, protectedRows = false) {
  const merges = rows.flatMap(row => [...row.cells].filter(cell => cell.rowSpan > 1));
  if (!protectedRows && !canMeasureMergedFragments(rows)) return null;
  const table = rows[0].closest('table')!;
  const naturalY = (value: number) => {
    const y = (value - g.origin) / g.scale;
    return y - g.removed.filter(gap => gap.top < y - .5).reduce((sum, gap) => sum + gap.height, 0);
  };
  const blocks = merges.flatMap(cell => [...cell.children].map(element => {
    const p = element as HTMLElement, pos = view.posAtDOM(p, 0) - 1;
    const lines = measureParagraphLines(view, p, naturalY);
    const terminal = protectedRows && p.dataset.hwpCellTerminal === 'true' ? measureHwpTerminalParagraph(p) : null;
    return { p, cell, pos, lines, terminal, offsets: lines.map((line, i) => i ? line.pos() - pos - 1 : 0), targets: new Array<number>(lines.length).fill(NaN) };
  }));
  if (protectedRows && blocks.some(block => block.p.dataset.hwpCellTerminal === 'true' && !block.terminal)) return null;
  const originalTop = naturalY(rows[0].getBoundingClientRect().top) + g.extra;
  const originalBottom = naturalY(rows.at(-1)!.getBoundingClientRect().bottom) + g.extra;
  const host = document.createElement('div'); host.className = 'tiptap'; host.setAttribute('aria-hidden', 'true');
  host.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;padding:0;min-height:0;width:${table.getBoundingClientRect().width / g.scale}px`;
  const probe = table.cloneNode(false) as HTMLTableElement;
  probe.style.margin = '0'; probe.style.width = `${table.getBoundingClientRect().width / g.scale}px`;
  const columns = table.querySelector(':scope > colgroup'); if (columns) probe.append(columns.cloneNode(true));
  const body = document.createElement('tbody'); probe.append(body); host.append(probe);
  const copyRow = (row: HTMLTableRowElement) => {
    const copy = row.cloneNode(true) as HTMLTableRowElement;
    for (const widget of copy.querySelectorAll('.ProseMirror-widget,.pagination-gap,.ProseMirror-trailingBreak,.ProseMirror-separator')) widget.remove();
    for (const node of [copy, ...copy.querySelectorAll<HTMLElement>('*')])
      node.classList.remove('pagination-merge-row', 'pagination-merge-padding', 'pagination-merge-flow', 'pagination-merge-label', 'pagination-split-row');
    return copy;
  };
  const copies = rows.map((row, i) => {
    const copy = copyRow(row); copy.dataset.fragmentRow = String(i);
    for (let n = 0; n < row.cells.length; n++) {
      const cell = row.cells[n], target = copy.cells[n];
      if (!merges.includes(cell)) {
        if (protectedRows && cell.tagName === 'TD') target.dataset.fragmentCell = String(n);
        continue;
      }
      for (let b = 0; b < cell.children.length; b++) {
        const id = blocks.findIndex(block => block.p === cell.children[b]);
        (target.children[b] as HTMLElement).dataset.fragmentBlock = String(id);
        (target.children[b] as HTMLElement).dataset.fragmentOffset = '0';
      }
    }
    return copy;
  });
  body.append(...copies); document.body.append(host);
  const ownerInsets = new Map<HTMLTableCellElement, number>();
  const sourceInsets = new Map<HTMLTableCellElement, number>();
  const seamOffsets = new Map<HTMLParagraphElement, number>();
  const hasCenteredOwner = protectedRows && merges.some(cell => getComputedStyle(cell).verticalAlign === 'middle');
  const hasOrdinaryCenters = hasCenteredOwner && hasCenteredOrdinaryCell(rows);
  const ordinaryTargets = new Map<HTMLParagraphElement, ReturnType<typeof measureOrdinaryParagraph> & {
    targetTop: number; terminal: ReturnType<typeof measureHwpTerminalParagraph>;
  }>();
  const owners = hasCenteredOwner ? rows.flatMap((row, rowIndex) => [...row.cells].flatMap((cell, columnIndex) => {
    if (!merges.includes(cell)) return [];
    const copy = copies[rowIndex].cells[columnIndex];
    return [{ cell, copy, row: copies[rowIndex], alignment: copy.style.verticalAlign }];
  })) : [];
  const paragraphInset = (cell: HTMLTableCellElement, row: HTMLTableRowElement) => {
    const paragraph = centeredContinuationParagraph(cell);
    return paragraph ? paragraph.getBoundingClientRect().top - row.getBoundingClientRect().top : undefined;
  };
  let leadingEdge: HTMLTableRowElement | null = null;
  let measurementHeaders: HTMLTableRowElement[] = [];
  const rowTargets: { top: number; height: number }[] = [], usedPages = new Set<number>();
  const repeatAt = g.measureRepeatAt ?? g.repeatAt;
  const stride = g.height + 24, usable = g.height - g.top - g.bottom;
  const canPlanLongGroup = g.longProtectedGroup && hasCenteredOwner && !hasOrdinaryCenters && !hasHwpCellLines([...rows, ...(g.continuationHeaders ?? [])]) && originalBottom - originalTop > usable + .5;
  let currentTop = originalTop, page = Math.max(0, Math.floor((currentTop - g.top + .5) / stride));
  let rowAnchor = true;
  let pending = copies, finalBottom = originalBottom;
  const fragments = tableFragmenter({ bodyHeight: () => usable - (g.reserveAt?.(page) ?? 0), hasDeferredNote: () => false, deferredNotes: () => new Set() });
  try {
    if (owners.length) {
      // An adjacent source row can determine the used collapsed border.
      const previous = precedingSourceRow(rows[0]);
      if (previous) {
        if (!columns || [...previous.cells].reduce((width, cell) => width + cell.colSpan, 0) !== columns.children.length) return null;
        leadingEdge = copyRow(previous);
        body.prepend(leadingEdge);
      }
      for (const entry of owners) entry.copy.style.verticalAlign = 'top';
      prepareParagraphEnds(probe); layoutWordLines(probe); layoutHwpCellLines(probe);
      for (const { cell, copy, row } of owners) {
        const first = copy.firstElementChild;
        if (!(first instanceof HTMLParagraphElement)) return null;
        const inset = first.getBoundingClientRect().top - row.getBoundingClientRect().top;
        const margin = parseFloat(getComputedStyle(first).marginTop) || 0;
        ownerInsets.set(cell, inset - margin);
      }
      for (const { copy, alignment } of owners) copy.style.verticalAlign = alignment;
      if (g.continuationHeaders?.length) {
        copies.forEach((row, i) => {
          for (const cell of row.cells) {
            if (cell.dataset.fragmentCell === undefined) continue;
            const inset = paragraphInset(cell, row);
            if (inset !== undefined) sourceInsets.set(rows[i].cells[Number(cell.dataset.fragmentCell)], inset);
          }
        });
        measurementHeaders = g.continuationHeaders.map(copyRow);
      }
    }

    for (let pass = 0; pending.length && pass < 512; pass++) {
      const first = !usedPages.size;
      const repeated = !first && measurementHeaders.length > 0;
      const prefix = repeated ? measurementHeaders : first && leadingEdge ? [leadingEdge] : [];
      body.replaceChildren(...prefix, ...pending); prepareParagraphEnds(probe); layoutWordLines(probe); layoutHwpCellLines(probe);
      const origin = probe.getBoundingClientRect().top;
      // Source and repeated-header positions address the first row; a new
      // bare table starts at its outer border. Do not count that half-border
      // twice when mapping a collapsed grid into a measured fragment.
      const inset = pending[0].getBoundingClientRect().top - origin;
      const tableTop = currentTop - (rowAnchor ? inset : 0);
      const boundary = page * stride + g.height - g.bottom - (g.reserveAt?.(page) ?? 0);
      let tail: HTMLTableRowElement[] | null = null;
      const fitsHere = () => probe.getBoundingClientRect().height <= boundary - tableTop + .5;
      if (!fitsHere()) {
        const fragmentHeight = probe.getBoundingClientRect().height - (prefix.length ? inset : 0);
        if (protectedRows && !canPlanLongGroup && (repeated || fragmentHeight + (repeatAt?.(page + 1) ?? 0) > usable + .5)) return null;
        const verifyProtectedFit = () => {
          layoutWordLines(probe);
          layoutHwpCellLines(probe);
          return fitsHere();
        };
        tail = fragments.splitMergedGroup(pending, origin + boundary - tableTop, protectedRows ? verifyProtectedFit : undefined);
        if (!tail) {
          // Like print, a group with no fitting first row starts on a fresh
          // sheet. An oversized unsplittable group retains the old fallback.
          if (protectedRows && (!usedPages.size || canPlanLongGroup)) return null;
          if (currentTop <= page * stride + g.top + (repeatAt?.(page) ?? 0) + .5) return null;
          page++; const header = repeatAt?.(page) ?? 0;
          currentTop = page * stride + g.top + header; rowAnchor = header > 0; continue;
        }
        layoutWordLines(probe); layoutHwpCellLines(probe);
        if (probe.getBoundingClientRect().height > boundary - tableTop + .5) return null;
      }
      const localY = (value: number) => tableTop + value - origin;
      usedPages.add(page);
      if (repeated) {
        const row = pending[0], source = rows[Number(row.dataset.fragmentRow)];
        for (const cell of row.cells) {
          if (cell.dataset.fragmentCell === undefined || hasOrdinaryCenters && getComputedStyle(cell).verticalAlign === 'middle') continue;
          const original = source.cells[Number(cell.dataset.fragmentCell)];
          const before = sourceInsets.get(original), after = paragraphInset(cell, row);
          if (before === undefined || after === undefined) return null;
          const offset = after - before;
          if (Math.abs(offset) >= 1 / 64) seamOffsets.set(original.firstElementChild as HTMLParagraphElement, offset);
        }
      }
      for (const row of pending) {
        const sourceIndex = Number(row.dataset.fragmentRow), rowBox = row.getBoundingClientRect();
        rowTargets[sourceIndex] = { top: localY(rowBox.top), height: rowBox.height };
        if (hasOrdinaryCenters) for (const cell of row.cells) {
          if (cell.dataset.fragmentCell === undefined || getComputedStyle(cell).verticalAlign !== 'middle') continue;
          const paragraph = centeredContinuationParagraph(cell);
          if (!paragraph) return null;
          const original = rows[sourceIndex].cells[Number(cell.dataset.fragmentCell)].firstElementChild as HTMLParagraphElement;
          const terminal = measureHwpTerminalParagraph(paragraph);
          if (original.dataset.hwpCellTerminal === 'true' ?
            !terminal || !sameHwpTerminalParagraph(terminal, measureHwpTerminalParagraph(original)) :
            paragraph.dataset.hwpCellTerminal === 'true') return null;
          const record = measureOrdinaryParagraph(paragraph);
          ordinaryTargets.set(original, { ...record, terminal, targetTop: localY(record.rect.top) });
        }
      }
      for (const p of probe.querySelectorAll<HTMLElement>('[data-fragment-block]')) {
        const block = blocks[Number(p.dataset.fragmentBlock)], offset = Number(p.dataset.fragmentOffset);
        if (block.terminal && !sameHwpTerminalParagraph(block.terminal, measureHwpTerminalParagraph(p))) return null;
        // DOM offsets include real BRs, matching ProseMirror's hard-break unit.
        const position = (node: Node, at: number) => {
          const range = document.createRange(); range.selectNodeContents(p); range.setEnd(node, at);
          return range.toString().length + range.cloneContents().querySelectorAll('br').length;
        };
        const lines = measureParagraphLines({ posAtDOM: position } as EditorView, p, localY);
        for (let i = 0; i < lines.length; i++) {
          const logical = offset + (i ? lines[i].pos() : 0), n = block.offsets.indexOf(logical);
          if (n < 0 || Number.isFinite(block.targets[n])) return null;
          block.targets[n] = lines[i].top;
        }
      }
      finalBottom = localY(pending.at(-1)!.getBoundingClientRect().bottom);
      pending = tail || [];
      // Cutting a merge can make its retained rows shorter. Print may then
      // fit the next source row on the same sheet; advance only when that
      // next group has no fitting row at the remaining physical boundary.
      if (pending.length && protectedRows) {
        page++;
        const header = measurementHeaders.length ? 0 : repeatAt?.(page) ?? 0;
        currentTop = page * stride + g.top + header; rowAnchor = header > 0;
      } else if (pending.length) { currentTop = finalBottom; rowAnchor = true; }
    }
    if (pending.length || rows.some((_, i) => !rowTargets[i]) || blocks.some(block => block.targets.some(y => !Number.isFinite(y)))) return null;
    const gaps = rowTargets.map((target, i) => Math.max(0, target.top -
      (i ? rowTargets[i - 1].top + rowTargets[i - 1].height : originalTop)));
    if (ordinaryTargets.size) {
      // Re-create the continuous editable grid, with the same height and
      // padding the view decorations will apply. A middle-aligned cell must
      // keep its authored alignment while its physical paragraph moves.
      const liveRows = rows.map((row, i) => {
        const copy = copyRow(row), gap = gaps[i];
        copy.style.height = `${rowTargets[i].height + gap}px`;
        [...row.cells].forEach((cell, n) => {
          if (gap > .5) copy.cells[n].style.paddingTop = `${(parseFloat(getComputedStyle(cell).paddingTop) || 0) + gap}px`;
          if (cell.rowSpan > 1) copy.cells[n].style.verticalAlign = 'top';
        });
        return copy;
      });
      body.replaceChildren(...(leadingEdge ? [leadingEdge] : []), ...liveRows);
      prepareParagraphEnds(probe);
      layoutWordLines(probe);
      layoutHwpCellLines(probe);
      const liveOrigin = liveRows[0].getBoundingClientRect().top - originalTop;
      for (const [i, row] of liveRows.entries()) {
        const live = row.getBoundingClientRect(), target = rowTargets[i];
        if (Math.abs(live.top - liveOrigin + gaps[i] - target.top) > .5 ||
          Math.abs(live.bottom - liveOrigin - target.top - target.height) > .5) return null;
        for (let n = 0; n < row.cells.length; n++) {
          const original = rows[i].cells[n].firstElementChild as HTMLParagraphElement;
          const targetP = ordinaryTargets.get(original);
          if (!targetP) continue;
          const projected = centeredContinuationParagraph(row.cells[n]);
          if (!projected) return null;
          const current = measureOrdinaryParagraph(projected);
          const geometryChanged = targetP.terminal ?
            !sameHwpTerminalParagraph(targetP.terminal, measureHwpTerminalParagraph(projected)) :
            projected.dataset.hwpCellTerminal === 'true' || Math.abs(current.rect.width - targetP.rect.width) > .5 ||
              current.offsets.length !== targetP.offsets.length || current.offsets.some((n, j) => n !== targetP.offsets[j]) ||
              current.boxes.length !== targetP.boxes.length ||
              current.boxes.some((box, j) => box.some((value, k) => Math.abs(value - targetP.boxes[j][k]) > .5));
          if (geometryChanged) return null;
          const shift = targetP.targetTop - (current.rect.top - liveOrigin);
          if (Math.abs(shift) >= 1 / 64) seamOffsets.set(original, shift);
        }
      }
    }
    const decorations: Decoration[] = [], signatures: string[] = [];
    rows.forEach((row, i) => {
      const target = rowTargets[i], gap = gaps[i], pos = view.posAtDOM(row, 0) - 1;
      decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, { class: 'pagination-merge-row',
        style: `--pagination-row-height:${target.height + gap}px;--pagination-source-row-height:${row.style.height || 'auto'}` }));
      signatures.push(`measured-merge-row:${pos}:${target.height}:${gap}`);
      if (gap > .5) for (const cell of [...row.cells]) {
        const p = view.posAtDOM(cell, 0) - 1, padding = parseFloat(getComputedStyle(cell).paddingTop) || 0;
        decorations.push(Decoration.node(p, p + view.state.doc.nodeAt(p)!.nodeSize, { class: 'pagination-merge-padding',
          style: `--pagination-cell-padding:${padding + gap}px;--pagination-source-cell-padding:${padding}px` }));
      }
    });
    for (const [paragraph, shift] of seamOffsets) {
      const pos = view.posAtDOM(paragraph, 0) - 1;
      decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize,
        { class: 'pagination-merge-label', style: `--pagination-label-shift:${shift}px` }));
      signatures.push(`measured-merge-seam:${pos}:${shift}`);
    }
    for (const cell of merges) {
      const own = blocks.filter(block => block.cell === cell), cellPos = view.posAtDOM(cell, 0) - 1;
      const style = getComputedStyle(cell), border = getComputedStyle(table).borderCollapse === 'collapse' ? .5 : 1;
      const padding = ownerInsets.get(cell) ?? ((parseFloat(style.paddingTop) || 0) + (parseFloat(style.borderTopWidth) || 0) * border);
      const firstTop = naturalY(own[0].p.getBoundingClientRect().top) - (parseFloat(getComputedStyle(own[0].p).marginTop) || 0);
      const base = rowTargets[rows.indexOf(cell.parentElement as HTMLTableRowElement)].top + padding - firstTop;
      decorations.push(Decoration.node(cellPos, cellPos + view.state.doc.nodeAt(cellPos)!.nodeSize,
        { class: 'pagination-merge-flow', style: `--pagination-source-valign:${style.verticalAlign}` }));
      let inserted = 0;
      for (const block of own) {
        const shift = block.targets[0] - (block.lines[0].top + base + inserted);
        for (let i = 1; i < block.lines.length; i++) {
          const gap = block.targets[i] - (block.lines[i].top + base + inserted + shift);
          if (gap < -.5 || protectedRows && gap > .5) return null;
          if (gap > .5) {
            const pos = block.lines[i].pos();
            decorations.push(Decoration.widget(pos, () => {
              const node = document.createElement('span'); node.className = 'pagination-cell-gap';
              node.contentEditable = 'false'; node.setAttribute('aria-hidden', 'true');
              node.style.cssText = `display:block;height:${gap}px;line-height:0;pointer-events:none`; return node;
            }, { side: -1, key: `measured-merge-line-${pos}-${gap}`, ignoreSelection: true }));
            inserted += gap; signatures.push(`measured-merge-line:${pos}:${gap}`);
          }
        }
        decorations.push(Decoration.node(block.pos, block.pos + view.state.doc.nodeAt(block.pos)!.nodeSize,
          { class: 'pagination-merge-label', style: `--pagination-label-shift:${shift}px` }));
        signatures.push(`measured-merge-label:${block.pos}:${shift}`);
      }
    }
    usedPages.forEach(page => g.repeatAt?.(page));
    return { decorations, signatures, extra: Math.max(0, finalBottom - originalBottom) };
  } finally { host.remove(); }
}
