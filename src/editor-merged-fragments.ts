import type { EditorView } from '@tiptap/pm/view';
import { Decoration } from '@tiptap/pm/view';
import { measureParagraphLines } from './editor-lines';
import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { layoutWordLines } from './word-line-metrics';
import { tableFragmenter } from './table-fragments';
import { prepareParagraphEnds } from './paragraph-fragments';

type Geometry = { origin: number; scale: number; top: number; bottom: number; height: number; extra: number;
  removed: { top: number; height: number }[]; repeatAt?: (page: number) => number; measureRepeatAt?: (page: number) => number; reserveAt?: (page: number) => number };

/** A browser redistributes minimum row heights when a vertical merge is split.
 * Measure those same print fragments instead of budgeting every continuation
 * against the unsplit table's original row distribution. Only the view changes. */
export function measuredMergedFragments(view: EditorView, rows: HTMLTableRowElement[], g: Geometry) {
  const merges = rows.flatMap(row => [...row.cells].filter(cell => cell.rowSpan > 1));
  if (!merges.some(cell => cell.querySelector('[data-hwp-cell-line]')) ||
    merges.some(cell => [...cell.children].some(p => !p.matches('p,h1,h2,h3,h4,h5,h6') ||
      p.querySelector('img:not(.ProseMirror-separator),svg,table,[data-page-number],[data-note-reference],[data-docx-tab],[data-docx-bookmark],[data-word-overlap]')))) return null;
  const table = rows[0].closest('table')!;
  const naturalY = (value: number) => {
    const y = (value - g.origin) / g.scale;
    return y - g.removed.filter(gap => gap.top < y - .5).reduce((sum, gap) => sum + gap.height, 0);
  };
  const blocks = merges.flatMap(cell => [...cell.children].map(element => {
    const p = element as HTMLElement, pos = view.posAtDOM(p, 0) - 1;
    const lines = measureParagraphLines(view, p, naturalY);
    return { p, cell, pos, lines, offsets: lines.map((line, i) => i ? line.pos() - pos - 1 : 0), targets: new Array<number>(lines.length).fill(NaN) };
  }));
  const originalTop = naturalY(rows[0].getBoundingClientRect().top) + g.extra;
  const originalBottom = naturalY(rows.at(-1)!.getBoundingClientRect().bottom) + g.extra;
  const host = document.createElement('div'); host.className = 'tiptap'; host.setAttribute('aria-hidden', 'true');
  host.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;padding:0;min-height:0;width:${table.getBoundingClientRect().width / g.scale}px`;
  const probe = table.cloneNode(false) as HTMLTableElement;
  probe.style.margin = '0'; probe.style.width = `${table.getBoundingClientRect().width / g.scale}px`;
  const columns = table.querySelector(':scope > colgroup'); if (columns) probe.append(columns.cloneNode(true));
  const body = document.createElement('tbody'); probe.append(body); host.append(probe);
  const copies = rows.map((row, i) => {
    const copy = row.cloneNode(true) as HTMLTableRowElement; copy.dataset.fragmentRow = String(i);
    for (const widget of copy.querySelectorAll('.ProseMirror-widget,.pagination-gap,.ProseMirror-trailingBreak,.ProseMirror-separator')) widget.remove();
    for (const node of [copy, ...copy.querySelectorAll<HTMLElement>('*')])
      node.classList.remove('pagination-merge-row', 'pagination-merge-padding', 'pagination-merge-flow', 'pagination-merge-label', 'pagination-split-row');
    for (let n = 0; n < row.cells.length; n++) {
      const cell = row.cells[n], target = copy.cells[n];
      if (!merges.includes(cell)) continue;
      for (let b = 0; b < cell.children.length; b++) {
        const id = blocks.findIndex(block => block.p === cell.children[b]);
        (target.children[b] as HTMLElement).dataset.fragmentBlock = String(id);
        (target.children[b] as HTMLElement).dataset.fragmentOffset = '0';
      }
    }
    return copy;
  });
  body.append(...copies); document.body.append(host);
  const rowTargets: { top: number; height: number }[] = [], usedPages = new Set<number>();
  const repeatAt = g.measureRepeatAt ?? g.repeatAt;
  const stride = g.height + 24, usable = g.height - g.top - g.bottom;
  let currentTop = originalTop, page = Math.max(0, Math.floor((currentTop - g.top + .5) / stride));
  let rowAnchor = true;
  let pending = copies, finalBottom = originalBottom;
  const fragments = tableFragmenter({ bodyHeight: () => usable - (g.reserveAt?.(page) ?? 0), hasDeferredNote: () => false, deferredNotes: () => new Set() });
  try {
    for (let pass = 0; pending.length && pass < 512; pass++) {
      body.replaceChildren(...pending); prepareParagraphEnds(probe); layoutWordLines(probe); layoutHwpCellLines(probe);
      const origin = probe.getBoundingClientRect().top;
      // Source and repeated-header positions address the first row; a new
      // bare table starts at its outer border. Do not count that half-border
      // twice when mapping a collapsed grid into a measured fragment.
      const inset = pending[0].getBoundingClientRect().top - origin;
      const tableTop = currentTop - (rowAnchor ? inset : 0);
      const boundary = page * stride + g.height - g.bottom - (g.reserveAt?.(page) ?? 0);
      let tail: HTMLTableRowElement[] | null = null;
      if (probe.getBoundingClientRect().height > boundary - tableTop + .5) {
        tail = fragments.splitMergedGroup(pending, origin + boundary - tableTop);
        if (!tail) {
          // Like print, a group with no fitting first row starts on a fresh
          // sheet. An oversized unsplittable group retains the old fallback.
          if (currentTop <= page * stride + g.top + (repeatAt?.(page) ?? 0) + .5) return null;
          page++; const header = repeatAt?.(page) ?? 0;
          currentTop = page * stride + g.top + header; rowAnchor = header > 0; continue;
        }
        layoutWordLines(probe); layoutHwpCellLines(probe);
        if (probe.getBoundingClientRect().height > boundary - tableTop + .5) return null;
      }
      const localY = (value: number) => tableTop + value - origin;
      usedPages.add(page);
      for (const row of pending) rowTargets[Number(row.dataset.fragmentRow)] = {
        top: localY(row.getBoundingClientRect().top), height: row.getBoundingClientRect().height,
      };
      for (const p of probe.querySelectorAll<HTMLElement>('[data-fragment-block]')) {
        const block = blocks[Number(p.dataset.fragmentBlock)], offset = Number(p.dataset.fragmentOffset);
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
      if (pending.length) { currentTop = finalBottom; rowAnchor = true; }
    }
    if (pending.length || rows.some((_, i) => !rowTargets[i]) || blocks.some(block => block.targets.some(y => !Number.isFinite(y)))) return null;
    const decorations: Decoration[] = [], signatures: string[] = [];
    let previousBottom = originalTop;
    rows.forEach((row, i) => {
      const target = rowTargets[i], gap = Math.max(0, target.top - previousBottom), pos = view.posAtDOM(row, 0) - 1;
      decorations.push(Decoration.node(pos, pos + view.state.doc.nodeAt(pos)!.nodeSize, { class: 'pagination-merge-row',
        style: `--pagination-row-height:${target.height + gap}px;--pagination-source-row-height:${row.style.height || 'auto'}` }));
      signatures.push(`measured-merge-row:${pos}:${target.height}:${gap}`);
      if (gap > .5) for (const cell of [...row.cells]) {
        const p = view.posAtDOM(cell, 0) - 1, padding = parseFloat(getComputedStyle(cell).paddingTop) || 0;
        decorations.push(Decoration.node(p, p + view.state.doc.nodeAt(p)!.nodeSize, { class: 'pagination-merge-padding',
          style: `--pagination-cell-padding:${padding + gap}px;--pagination-source-cell-padding:${padding}px` }));
      }
      previousBottom = target.top + target.height;
    });
    for (const cell of merges) {
      const own = blocks.filter(block => block.cell === cell), cellPos = view.posAtDOM(cell, 0) - 1;
      const style = getComputedStyle(cell), border = getComputedStyle(table).borderCollapse === 'collapse' ? .5 : 1;
      const padding = (parseFloat(style.paddingTop) || 0) + (parseFloat(style.borderTopWidth) || 0) * border;
      const firstTop = naturalY(own[0].p.getBoundingClientRect().top) - (parseFloat(getComputedStyle(own[0].p).marginTop) || 0);
      const base = rowTargets[rows.indexOf(cell.parentElement as HTMLTableRowElement)].top + padding - firstTop;
      decorations.push(Decoration.node(cellPos, cellPos + view.state.doc.nodeAt(cellPos)!.nodeSize,
        { class: 'pagination-merge-flow', style: `--pagination-source-valign:${style.verticalAlign}` }));
      let inserted = 0;
      for (const block of own) {
        const shift = block.targets[0] - (block.lines[0].top + base + inserted);
        for (let i = 1; i < block.lines.length; i++) {
          const gap = block.targets[i] - (block.lines[i].top + base + inserted + shift);
          if (gap < -.5) return null;
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
