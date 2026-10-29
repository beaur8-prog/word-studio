import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { paragraphLineOffsets, paragraphTextWalker, splitText, textUnitLength } from './paragraph-fragments';

interface FragmentContext { bodyHeight(): number; hasDeferredNote(element: HTMLElement): boolean; deferredNotes(): ReadonlySet<string> }

/** Shared measured table fragments for print and editable continuation plans. */
export function tableFragmenter(context: FragmentContext) {
  // Split editable form rows at paragraph or nested-table row boundaries. A
  // large single cell must not shrink all its text just to fit one sheet.
  function splitElement(element: HTMLElement, boundary: number, depth = 0): HTMLElement | null {
    if (depth > 30) return null;
    if (element.matches('p,h1,h2,h3,h4,h5,h6') && ((element.textContent?.length ?? 0) > 1 || element.querySelector('br'))) {
      let offsets = paragraphLineOffsets(element);
      if (context.hasDeferredNote(element)) {
        const walker = paragraphTextWalker(element); let offset = 0;
        while (walker.nextNode()) {
          const text = walker.currentNode, ref = text.parentElement?.closest<HTMLElement>('[data-note-reference]');
          if (ref && context.deferredNotes().has(ref.dataset.noteReference!)) break;
          offset += textUnitLength(text);
        }
        const limit = Math.max(0, ...offsets.filter(n => n <= offset));
        offsets = offsets.filter(n => n <= limit);
      }
      let low = 0, high = offsets.length - 1, best = 0;
      const marker = document.createComment('split'); element.replaceWith(marker);
      while (low <= high) {
        const mid = (low + high) >> 1, [head] = splitText(element, offsets[mid]); marker.before(head);
        layoutHwpCellLines(head);
        if (head.getBoundingClientRect().bottom <= boundary + .5) { best = offsets[mid]; low = mid + 1; } else high = mid - 1;
        head.remove();
      }
      marker.replaceWith(element);
      if (!best) return null;
      const [head, tail] = splitText(element, best); element.replaceChildren(...head.childNodes); element.style.marginBottom = '0';
      element.dataset.printSoftEnd = 'true'; delete element.dataset.printTrailingBreak;
      if (head.dataset.hwpCellContinuation) element.dataset.hwpCellContinuation = 'true';
      if (head.dataset.hwpCellFullLine) element.dataset.hwpCellFullLine = 'true';
      layoutHwpCellLines(element); return tail;
    }
    if (element.matches('tr')) {
      if (element.dataset.rowSplit === 'avoid' || element.dataset.wordRowRule === 'exact') return null;
      const row = element as HTMLTableRowElement;
      if (row.cells.length) {
        if (row.cells.length > 1 && row.dataset.rowSplit !== 'allow' || [...row.cells].some(c => c.rowSpan > 1)) return null;
        const original = row.cloneNode(true) as HTMLTableRowElement, height = row.getBoundingClientRect().height;
        const available = boundary - row.getBoundingClientRect().top;
        if (available <= .5) return null;
        const tail = row.cloneNode(false) as HTMLTableRowElement;
        row.style.removeProperty('height');
        let moved = false, reservedSpace = false;
        // Split every column against the same physical boundary. Even a
        // finished column needs its empty continuation cell to keep the grid.
        for (const cell of [...row.cells]) {
          const next = cell.cloneNode(false) as HTMLTableCellElement; tail.append(next);
          // Reflow a long cell from its top while finding the cut. A stored
          // row minimum and middle/bottom alignment otherwise move each trial
          // paragraph as it gets shorter, changing the available line budget.
          const align = cell.style.verticalAlign;
          const style = getComputedStyle(cell);
          const innerHeight = [...cell.children].reduce((sum, child) => {
            const css = getComputedStyle(child);
            return sum + child.getBoundingClientRect().height + (parseFloat(css.marginTop) || 0) + (parseFloat(css.marginBottom) || 0);
          }, 0);
          if (innerHeight > available - (parseFloat(style.paddingTop) || 0) - (parseFloat(style.paddingBottom) || 0) + .5) cell.style.verticalAlign = 'top';
          const rest = splitElement(cell, boundary, depth + 1);
          if (rest) { next.append(...rest.childNodes); moved = true; }
          else if (context.hasDeferredNote(cell) || [...cell.children].some(child => child.getBoundingClientRect().bottom > boundary - (parseFloat(getComputedStyle(cell).paddingBottom) || 0) + .5)) {
            next.append(...cell.childNodes); moved = true;
          }
          cell.style.verticalAlign = align;
        }
        // A stored minimum height can reserve space after the text. Divide
        // that reserve too; repeating the full height on the continuation
        // would enlarge the row, while dropping it would shorten the form.
        if (!moved && height > available + .5) {
          row.style.height = `${available}px`; moved = reservedSpace = true;
        }
        const progress = reservedSpace || [...row.cells].some(c => c.textContent?.trim() || c.querySelector('img,table'));
        if (!moved || !progress || row.getBoundingClientRect().bottom > boundary + .5) {
          row.replaceChildren(...original.childNodes); row.style.cssText = original.style.cssText; return null;
        }
        tail.style.height = `${Math.max(0, height - row.getBoundingClientRect().height)}px`;
        return tail;
      }
    }
    if (element.matches('table')) {
      const table = element as HTMLTableElement;
      // A nested form table that fits a sheet moves as a unit. Splitting its
      // cell paragraphs merely to fill the preceding sheet changes the form.
      if (table.rows.length === 1 && table.getBoundingClientRect().height <= context.bodyHeight() + .5) return null;
      if ([...table.rows].some(row => [...row.cells].some(cell => cell.rowSpan > 1))) return null;
    }
    const children = [...element.children] as HTMLElement[];
    const style = getComputedStyle(element);
    const innerBoundary = boundary - (parseFloat(style.paddingBottom) || 0) - (parseFloat(style.borderBottomWidth) || 0);
    const index = children.findIndex(child => !child.matches('colgroup') && (context.hasDeferredNote(child) || child.getBoundingClientRect().bottom + (parseFloat(getComputedStyle(child).marginBottom) || 0) > innerBoundary + .5));
    if (index < 0) return null;
    const child = children[index], tail = element.cloneNode(false) as HTMLElement;
    const childTail = splitElement(child, innerBoundary - (parseFloat(getComputedStyle(child).marginBottom) || 0), depth + 1);
    if (childTail) tail.append(childTail);
    else if (index === 0 || children.slice(0, index).every(c => c.matches('colgroup'))) return null;
    else tail.append(child);
    children.slice(index + 1).forEach(c => tail.append(c));
    if (element.matches('table')) element.querySelectorAll(':scope > colgroup').forEach(c => tail.prepend(c.cloneNode(true)));
    if (element.matches('tr')) { element.style.removeProperty('height'); tail.style.removeProperty('height'); }
    return tail;
  }
  // Only explicitly permitted continuing merges may cross a row boundary. Keep
  // the grid and every cell's content exactly once; this is print DOM only.
  function splitMergedGroup(group: HTMLTableRowElement[], boundary: number): HTMLTableRowElement[] | null {
    if (group.length < 2 || group.some(r => r.dataset.rowSplit === 'avoid' || r.dataset.wordRowRule === 'exact')) return null;
    const entries: { cell: HTMLTableCellElement; row: number; col: number }[] = [];
    const occupied: number[] = [];
    for (let row = 0; row < group.length; row++) {
      let col = 0;
      for (const cell of [...group[row].cells]) {
        while ((occupied[col] || 0) > row) col++;
        entries.push({ cell, row, col });
        for (let j = col; j < col + cell.colSpan; j++) occupied[j] = row + cell.rowSpan;
        col += cell.colSpan;
      }
    }
    for (let cut = group.length - 1; cut > 0; cut--) {
      if (group[cut - 1].getBoundingClientRect().bottom > boundary + .5) continue;
      const crossings = entries.filter(e => e.row < cut && e.row + e.cell.rowSpan > cut);
      if (!crossings.length || crossings.some(e => e.cell.dataset.cellSplit !== 'allow')) continue;
      const tails = group.slice(cut);
      // The continuation starts at a source row boundary, not at the bottom
      // of the sheet. Filling the unused area below that row stretches the
      // retained rows and makes editor and print disagree about the grid.
      const rowBoundary = group[cut - 1].getBoundingClientRect().bottom;
      for (const { cell, row, col } of crossings) {
        const next = cell.cloneNode(false) as HTMLTableCellElement;
        next.dataset.printMergeContinuation = 'true';
        // Partition rich content using the available row height. Applying a
        // whole-cell middle/bottom offset before partitioning can leave most
        // of the first fragment empty and overflow the last one. Each printed
        // fragment then applies the original vertical alignment independently.
        const align = cell.style.verticalAlign;
        const cellStyle = getComputedStyle(cell), child = cell.firstElementChild as HTMLElement | null;
        const childStyle = child && getComputedStyle(child);
        const rich = cell.dataset.printMergeFlow === 'rich' || cell.children.length > 1 ||
          (child?.getBoundingClientRect().height ?? 0) + (parseFloat(childStyle?.marginTop || '') || 0) + (parseFloat(childStyle?.marginBottom || '') || 0) >
          rowBoundary - cell.getBoundingClientRect().top - (parseFloat(cellStyle.paddingTop) || 0) - (parseFloat(cellStyle.paddingBottom) || 0);
        if (rich) { cell.style.verticalAlign = 'top'; next.dataset.printMergeFlow = 'rich'; }
        const remainder = splitElement(cell, rowBoundary);
        if (remainder) next.append(...remainder.childNodes);
        else if ([...cell.children].some(c => c.getBoundingClientRect().bottom > rowBoundary + .5)) next.append(...cell.childNodes);
        cell.style.verticalAlign = align;
        next.rowSpan = row + cell.rowSpan - cut; cell.rowSpan = cut - row;
        cell.style.removeProperty('height'); next.style.removeProperty('height');
        const following = entries.find(e => e.row === cut && e.col > col)?.cell;
        tails[0].insertBefore(next, following || null);
      }
      tails.forEach(r => r.remove()); group.splice(cut);
      return tails;
    }
    return null;
  }
  return { splitElement, splitMergedGroup };
}
