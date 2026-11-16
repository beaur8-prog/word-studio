import { rowCanContinue, rowSplitPolicy } from './row-split-policy';
import { layoutHwpCellLines } from './hwp-cell-line-metrics';
import { paragraphLineOffsets, paragraphTextWalker, splitText, textUnitLength } from './paragraph-fragments';

function headerRow(row: HTMLTableRowElement) {
  return row.cells.length > 0 && [...row.cells].every(cell => cell.tagName === 'TH');
}

export function canPlaceProtectedTable(root: HTMLElement) {
  return !root.querySelector('[data-note-reference],img[data-image-floating],table[data-table-anchor],[data-document-textbox]');
}

export function centeredContinuationParagraph(cell: HTMLTableCellElement) {
  const paragraph = cell.firstElementChild;
  return paragraph instanceof HTMLParagraphElement && cell.children.length === 1 &&
    getComputedStyle(cell).verticalAlign === 'top' ? paragraph : undefined;
}

export function precedingSourceRow(row: HTMLTableRowElement) {
  let previous = row.previousElementSibling;
  while (previous?.classList.contains('ProseMirror-widget')) previous = previous.previousElementSibling;
  return previous instanceof HTMLTableRowElement ? previous : null;
}

export function canSplitProtectedMergedRows(rows: HTMLTableRowElement[]) {
  const table = rows[0]?.closest('table');
  if (!table || table.dataset.wordTableLayout !== 'fixed' || table.dataset.tableAnchor ||
    table.parentElement?.closest('td,th,[data-column-count]:not([data-column-count="1"]),.document-textbox,.document-furniture')) return false;
  const firstBody = rows.findIndex(row => !headerRow(row));
  if (firstBody < 0 || rows.length - firstBody < 2) return false;
  const supportedRows = rows.every((row, index) => (index < firstBody ? [...row.cells].every(cell => cell.rowSpan === 1) :
      row.dataset.rowSplit === 'avoid' && row.dataset.wordRowRule !== 'exact' && Number.parseFloat(row.style.height) > 0 &&
      [...row.cells].every(cell => cell.tagName === 'TD') && [...row.cells].some(cell => cell.rowSpan === 1)) &&
    [...row.cells].every(cell => cell.rowSpan >= 1 && !cell.style.height && !cell.style.minHeight) &&
    !row.querySelector('table,img:not(.ProseMirror-separator),svg,[data-note-reference],[data-page-number],[data-word-keep-next="true"],[data-hwp-cell-line]') &&
    row.getBoundingClientRect().height > 0);
  if (!supportedRows) return false;
  const body = rows.slice(firstBody);
  if (!body.some(row => [...row.cells].some(cell => cell.rowSpan > 1 && getComputedStyle(cell).verticalAlign === 'middle'))) return true;

  // The editor must reproduce the next physical paper's used border without
  // changing row height or moving other paragraphs. Limit this path to cells
  // whose whole paragraph can be mapped back to the continuous editable row.
  const columns = table.querySelector(':scope > colgroup')?.children.length ?? 0;
  const previous = precedingSourceRow(body[0]);
  return columns > 0 &&
    (!previous || !previous.querySelector(':scope > [data-print-merge-continuation]') &&
      [...previous.cells].reduce((width, cell) => width + cell.colSpan, 0) === columns) &&
    body.every(row => [...row.cells].every(cell => {
      if (cell.rowSpan === 1) return centeredContinuationParagraph(cell);
      const paragraph = cell.firstElementChild, alignment = getComputedStyle(cell).verticalAlign;
      return paragraph instanceof HTMLParagraphElement && cell.childElementCount === 1 &&
        (alignment === 'top' || alignment === 'middle');
    }));
}

function mergeContentFitsOwnerRow(cell: HTMLTableCellElement, row: HTMLTableRowElement, borderFactor: number) {
  if (cell.dataset.printMergeContinuation && !cell.hasChildNodes()) return true;
  const p = cell.firstElementChild;
  const style = getComputedStyle(cell);
  if (style.verticalAlign !== 'top' || cell.children.length !== 1 || !p?.matches('p') ||
    p.querySelector('[data-docx-tab],[data-docx-bookmark],[data-word-overlap]')) return false;
  const padding = Number.parseFloat(style.paddingBottom) || 0;
  const border = (Number.parseFloat(style.borderBottomWidth) || 0) * borderFactor;
  const margin = Math.max(0, Number.parseFloat(getComputedStyle(p).marginBottom) || 0);
  const range = document.createRange();
  range.selectNodeContents(p);
  const contentBottom = Math.max(p.getBoundingClientRect().bottom, range.getBoundingClientRect().bottom);
  return contentBottom + margin + padding + border <= row.getBoundingClientRect().bottom + .5;
}

function measureCenteredParagraph(cell: HTMLTableCellElement) {
  const paragraph = cell.firstElementChild;
  if (getComputedStyle(cell).verticalAlign !== 'middle' || cell.children.length !== 1 || !(paragraph instanceof HTMLParagraphElement) ||
    !paragraph.textContent?.trim() || paragraph.querySelector('[data-docx-tab],[data-docx-bookmark],[data-word-overlap],[data-symbol-font]')) return null;
  const walker = paragraphTextWalker(paragraph), range = document.createRange(), ink: DOMRect[] = [];
  while (walker.nextNode()) {
    if (!(walker.currentNode instanceof Text)) continue;
    range.selectNodeContents(walker.currentNode);
    ink.push(...[...range.getClientRects()].filter(rect => rect.width > .1 && rect.height > .1));
  }
  if (!ink.length) return null;
  const box = paragraph.getBoundingClientRect();
  const css = getComputedStyle(paragraph);
  return { box, ink, lines: paragraphLineOffsets(paragraph),
    top: Math.min(box.top - Math.max(0, parseFloat(css.marginTop) || 0), ...ink.map(rect => rect.top)),
    bottom: Math.max(box.bottom + Math.max(0, parseFloat(css.marginBottom) || 0), ...ink.map(rect => rect.bottom)) };
}

function centeredParagraphFits(cell: HTMLTableCellElement, before: NonNullable<ReturnType<typeof measureCenteredParagraph>>, borderFactor: number) {
  const after = measureCenteredParagraph(cell);
  if (!after || after.lines.length !== before.lines.length || after.lines.some((line, index) => line !== before.lines[index]) ||
    after.ink.length !== before.ink.length || Math.abs(after.box.width - before.box.width) > .5 ||
    Math.abs(after.box.height - before.box.height) > .5) return false;
  const bounds = cell.getBoundingClientRect(), css = getComputedStyle(cell);
  const top = bounds.top + (parseFloat(css.paddingTop) || 0) + (parseFloat(css.borderTopWidth) || 0) * borderFactor;
  const bottom = bounds.bottom - (parseFloat(css.paddingBottom) || 0) - (parseFloat(css.borderBottomWidth) || 0) * borderFactor;
  const left = bounds.left + (parseFloat(css.paddingLeft) || 0) + (parseFloat(css.borderLeftWidth) || 0) * borderFactor;
  const right = bounds.right - (parseFloat(css.paddingRight) || 0) - (parseFloat(css.borderRightWidth) || 0) * borderFactor;
  if (after.top < top - .5 || after.bottom > bottom + .5) return false;
  return after.ink.every((ink, index) => {
    const source = before.ink[index];
    return ink.left >= left - .5 && ink.right <= right + .5 &&
      Math.abs(ink.x - after.box.x - (source.x - before.box.x)) <= .5 &&
      Math.abs(ink.y - after.box.y - (source.y - before.box.y)) <= .5 &&
      Math.abs(ink.width - source.width) <= .5 && Math.abs(ink.height - source.height) <= .5;
  });
}

interface FragmentContext { bodyHeight(): number; hasDeferredNote(element: HTMLElement): boolean; deferredNotes(): ReadonlySet<string> }

/** Shared measured table fragments for print and editable continuation plans. */
export function tableFragmenter(context: FragmentContext) {
  // Split editable form rows at paragraph or nested-table row boundaries. A
  // large single cell must not shrink all its text just to fit one sheet.
  function hasEarlierBodyContent(row: HTMLTableRowElement) {
    const body = row.closest<HTMLElement>('.print-content');
    if (!body) return false;
    let start = body.getBoundingClientRect().top;
    for (let node: HTMLElement | null = row; node && node !== body; node = node.parentElement) {
      const css = getComputedStyle(node);
      if (node.matches('td,th')) {
        const collapsed = getComputedStyle(node.closest('table')!).borderCollapse === 'collapse';
        start += (parseFloat(css.paddingTop) || 0) + (parseFloat(css.borderTopWidth) || 0) / (collapsed ? 2 : 1);
      }
      if (node instanceof HTMLTableElement && node.rows.length) {
        start += (parseFloat(css.marginTop) || 0) + Math.max(0, node.rows[0].getBoundingClientRect().top - node.getBoundingClientRect().top);
        let headerEnd = 0;
        while (headerEnd < node.rows.length && node.rows[headerEnd].cells.length && [...node.rows[headerEnd].cells].every(cell => cell.tagName === 'TH')) headerEnd++;
        // Repeated headers and enclosing cell padding belong at the top of
        // every fragment; they are not preceding body text to move again.
        if (headerEnd && ![...node.rows].slice(0, headerEnd).some(header => header.contains(row)))
          start += node.rows[headerEnd - 1].getBoundingClientRect().bottom - node.rows[0].getBoundingClientRect().top;
      }
    }
    return row.getBoundingClientRect().top > start + .5;
  }
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
      const canContinue = rowCanContinue(element.dataset, element.getBoundingClientRect().height, context.bodyHeight());
      if (element.dataset.wordRowRule === 'exact' || rowSplitPolicy(element.dataset) === 'avoid' && !canContinue) return null;
      const row = element as HTMLTableRowElement;
      // A nested oversized cantSplit row must also begin on a fresh body.
      // Returning it intact lets the enclosing cell move it into its tail.
      if (rowSplitPolicy(row.dataset) === 'avoid' && canContinue && hasEarlierBodyContent(row)) return null;
      if (row.cells.length) {
        if (row.cells.length > 1 && !canContinue || [...row.cells].some(c => c.rowSpan > 1)) return null;
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
      // An atomic nested form moves as a unit. An explicit cell-break
      // permission must still work when preceding edits leave less space.
      if (table.rows.length === 1 && rowSplitPolicy(table.rows[0].dataset) !== 'allow' &&
        table.getBoundingClientRect().height <= context.bodyHeight() + .5) return null;
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
  function splitMergedGroup(group: HTMLTableRowElement[], boundary: number, protectedFit?: () => boolean): HTMLTableRowElement[] | null {
    if (group.length < 2) return null;
    if (protectedFit) {
      if (group.at(-1)!.getBoundingClientRect().bottom <= boundary + .5) return null;
    } else if (group.some(r => rowSplitPolicy(r.dataset) === 'avoid' || r.dataset.wordRowRule === 'exact')) return null;
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
    const firstBody = protectedFit ? group.findIndex(row => !headerRow(row)) : 0;
    if (firstBody < 0) return null;
    const borderFactor = protectedFit && getComputedStyle(group[0].closest('table')!).borderCollapse === 'collapse' ? .5 : 1;
    for (let cut = group.length - 1; cut > firstBody; cut--) {
      if (group[cut - 1].getBoundingClientRect().bottom > boundary + .5) continue;
      const crossings = entries.filter(e => e.row < cut && e.row + e.cell.rowSpan > cut);
      if (!crossings.length) continue;
      const tails = group.slice(cut);
      if (protectedFit) {
        // Keep the owner paragraph intact. A centered paragraph is measured
        // again in the shortened cell because its unsplit center cannot be
        // used to decide whether its ink fits in the retained fragment.
        const centered = new Map<HTMLTableCellElement, NonNullable<ReturnType<typeof measureCenteredParagraph>>>();
        for (const { cell } of crossings) {
          const paragraph = measureCenteredParagraph(cell);
          if (paragraph) centered.set(cell, paragraph);
        }
        if (crossings.some(({ cell, row }) => !centered.has(cell) && !mergeContentFitsOwnerRow(cell, group[row], borderFactor))) continue;
        const headRows = group.slice(0, cut);
        const rowHeights = centered.size ? headRows.map(row => row.getBoundingClientRect().height) : [];
        const parent = tails[0].parentNode!;
        const followingRow = tails.at(-1)!.nextSibling;
        const spans = crossings.sort((a, b) => a.col - b.col).map(({ cell, row, col }) => {
          const span = cell.rowSpan;
          const next = cell.cloneNode(false) as HTMLTableCellElement;
          next.dataset.printMergeContinuation = 'true';
          next.rowSpan = row + span - cut;
          cell.rowSpan = cut - row;
          const following = entries.find(entry => entry.row === cut && entry.col > col)?.cell;
          tails[0].insertBefore(next, following || null);
          return { cell, span, next, row };
        });
        tails.forEach(row => row.remove());
        const fitsTop = () => group[cut - 1].getBoundingClientRect().bottom <= boundary + .5 &&
          spans.every(({ cell, row }) => mergeContentFitsOwnerRow(cell, group[row], borderFactor)) && protectedFit();
        const fitsCentered = () => protectedFit() && group[cut - 1].getBoundingClientRect().bottom <= boundary + .5 &&
          headRows.every((row, index) => Math.abs(row.getBoundingClientRect().height - rowHeights[index]) <= .5) &&
          spans.every(({ cell, row }) => {
            const prior = centered.get(cell);
            return prior ? centeredParagraphFits(cell, prior, borderFactor) : mergeContentFitsOwnerRow(cell, group[row], borderFactor);
          });
        if (centered.size ? fitsCentered() : fitsTop()) {
          group.splice(cut);
          return tails;
        }
        for (const { cell, span, next } of spans) { cell.rowSpan = span; next.remove(); }
        for (const tail of tails) parent.insertBefore(tail, followingRow);
        continue;
      }
      if (crossings.some(e => e.cell.dataset.cellSplit !== 'allow')) continue;
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
