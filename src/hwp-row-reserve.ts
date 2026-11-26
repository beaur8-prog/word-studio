import { cssPixels } from './hwp-cell-line';
import { cssPoints } from './css-length';
import { paragraphLineOffsets } from './paragraph-fragments';
import { rowSplitPolicy } from './row-split-policy';
import { hwpSourceWrapAttribute, hwpWrapFingerprint } from './hwp-text-advances';

type HwpRowReserve = { kind: 1 | 2 | 3 | 4; authoredHeight: number; observedHeight: number; signature: string; headHeight?: number };

export function hwpRowReserve(value: unknown): HwpRowReserve | null {
  if (typeof value !== 'string' || value.length > 64) return null;
  const match = /^([1234]) (\d+(?:\.\d+)?) (\d+(?:\.\d+)?)(?: (\d+(?:\.\d+)?))? ([a-f0-9]{8}|[a-f0-9]{16})$/.exec(value);
  if (!match || match[5].length !== (match[1] === '1' ? 8 : 16)) return null;
  const authoredHeight = Number(match[2]), observedHeight = Number(match[3]), headHeight = Number(match[4]);
  const kind = Number(match[1]) as HwpRowReserve['kind'];
  if (kind === 4 ? !(headHeight > 0 && headHeight < observedHeight) : match[4] !== undefined) return null;
  return observedHeight > 0 && observedHeight <= 2000 && authoredHeight > 0 && authoredHeight <= 2000 &&
    (kind === 3 || observedHeight < authoredHeight) ? { kind, authoredHeight, observedHeight, signature: match[5],
      ...(kind === 4 && { headHeight }) } : null;
}

function terminalCells(row: HTMLTableRowElement): HTMLTableCellElement[] | null {
  const previous = row.previousElementSibling;
  if (!(previous instanceof HTMLTableRowElement) || previous.cells.length !== 2 || row.cells.length !== 1) return null;
  const cells = [...previous.cells];
  const owner = cells.find(cell => cell.rowSpan === 2);
  if (!owner || cells.some(cell => cell.tagName !== 'TD' || cell.colSpan !== 1 || cell !== owner && cell.rowSpan !== 1)) return null;
  const ordinary = row.cells[0];
  if (ordinary.tagName !== 'TD' || ordinary.colSpan !== 1 || ordinary.rowSpan !== 1) return null;
  return [owner, ordinary];
}

// Only typography and box properties affect this fit budget. An ordinary
// spelling edit still reflows, while a changed line style retires the budget.
function signature(cells: HTMLTableCellElement[]): string | null {
  const values: (string | number)[] = [];
  const length = (raw: string) => {
    const px = cssPixels(raw);
    return px === null ? raw : Math.round(px * 15);
  };
  const spacing = (raw: string) => !raw || /^0(?:[.]0+)?(?:px|pt)$/.test(raw) ? 0 : length(raw);
  const family = (name: string) => name.replace(/['"\s]/g, '').toLowerCase();
  for (const cell of cells) {
    const column = cell.getAttribute('colwidth');
    if (!column || !/^\d{1,4}$/.test(column) || cell.style.verticalAlign !== 'top' || !cell.children.length) return null;
    const style = cell.style;
    values.push(column, cell.children.length, spacing(style.paddingTop), spacing(style.paddingBottom),
      spacing(style.paddingLeft), spacing(style.paddingRight), style.whiteSpace);
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const border = style.getPropertyValue(`border-${side}-style`);
      values.push(border && border !== 'none' ? spacing(style.getPropertyValue(`border-${side}-width`)) : 0);
    }
    for (const node of cell.children) {
      if (!(node instanceof HTMLParagraphElement) || !node.textContent?.trim() || node.childElementCount > 1) return null;
      const p = node.style, run = node.firstElementChild;
      if (run && (!(run instanceof HTMLSpanElement) || run.childElementCount !== 0)) return null;
      values.push(length(p.lineHeight), length(p.fontSize), spacing(p.marginTop), spacing(p.marginBottom),
        spacing(p.marginLeft), spacing(p.marginRight), spacing(p.textIndent), family(p.fontFamily),
        p.whiteSpace, p.wordBreak, node.dataset.hwpCellLine || '');
      if (run instanceof HTMLSpanElement) values.push(length(run.style.fontSize), family(run.style.fontFamily),
        spacing(run.style.letterSpacing), run.style.fontWeight, run.style.fontStyle);
    }
  }
  let hash = 2166136261;
  for (const char of JSON.stringify(values)) hash = Math.imul(hash ^ char.charCodeAt(0), 16777619);
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function containerSignature(row: HTMLTableRowElement, inlineStroke = false) {
  const table = row.closest('table'), cell = row.cells[0];
  if (!table || table.parentElement?.closest('table') || table.rows.length !== 1 || row.cells.length !== 1 ||
      !cell || cell.tagName !== 'TD' || cell.colSpan !== 1 || cell.rowSpan !== 1 ||
      (inlineStroke ? rowSplitPolicy(row.dataset) === 'allow'
        : row.dataset.rowSplit !== 'allow' || cell.style.verticalAlign !== 'top') ||
      !cell.querySelector(':scope > table') || cell.querySelector('img,svg,[data-note-reference],[data-document-textbox]')) return;
  const blocks = [...cell.children];
  if (!blocks.every(block => block.matches('p,table'))) return;
  // Literal text and the order of authored objects and empty paragraphs
  // must still match. Preserve the standard row height for DOCX readers.
  const structure = [...cell.querySelectorAll('p,table,tr,td,th,br,[data-docx-tab]')].map(block => [
    block.tagName, block.getAttribute('colwidth'), block.getAttribute('colspan'), block.getAttribute('rowspan'),
    block.getAttribute(hwpSourceWrapAttribute),
  ]);
  const inkStyle = inlineStroke ? [cell.style.verticalAlign, cell.style.borderTopStyle, cell.style.borderBottomStyle,
    ...[cell.style.borderTopWidth, cell.style.borderBottomWidth].map(value => Math.round((cssPoints(value) ?? 0) * 8)),
    ...[cell.style.paddingTop, cell.style.paddingBottom, cell.style.paddingLeft, cell.style.paddingRight]
      .map(value => Math.round((cssPixels(value) ?? 0) * 15))] : undefined;
  return hwpWrapFingerprint(JSON.stringify([cell.getAttribute('colwidth'), blocks.map(block => block.textContent), structure,
    ...(inkStyle ? [inkStyle] : [])]));
}

export function printedHwpContainerHeight(row: HTMLTableRowElement) {
  const reserve = hwpRowReserve(row.dataset.hwpRowReserve), height = cssPixels(row.style.height);
  if (reserve?.kind !== 2 || height === null || Math.abs(height - reserve.authoredHeight) > .1 ||
      row.dataset.wordRowRule === 'exact' || containerSignature(row) !== reserve.signature) return;
  return reserve.observedHeight;
}

/** Reserve the measured logical cell while drawing both full borders. */
export function reserveHwpInlineStrokes(table: HTMLTableElement) {
  const row = table.rows[0], cell = row?.cells[0];
  if (!row || !cell || table.dataset.tableAnchor || table.style.zoom ||
      table.parentElement?.closest('td,th,[data-column-count]:not([data-column-count="1"])') ||
      getComputedStyle(table).borderCollapse !== 'collapse') return;
  const reserve = hwpRowReserve(row.dataset.hwpRowReserve), minimum = cssPixels(row.style.height);
  if (reserve?.kind !== 3 || minimum === null || Math.abs(reserve.authoredHeight - minimum) > .1 ||
      row.dataset.wordRowRule === 'exact' || containerSignature(row, true) !== reserve.signature) return;
  const blank = (element: Element | null): element is HTMLParagraphElement => {
    if (!(element instanceof HTMLParagraphElement) || element.textContent?.trim() ||
        element.matches('[data-list-glyph],[data-word-list-id]') ||
        element.querySelector('img,svg,table,[data-note-reference],[data-document-field],[data-page-number]')) return false;
    const style = getComputedStyle(element);
    return style.backgroundColor === 'rgba(0, 0, 0, 0)' && style.boxShadow === 'none' && style.outlineStyle === 'none' &&
      [style.borderTopWidth, style.borderBottomWidth, style.borderLeftWidth, style.borderRightWidth].every(value => parseFloat(value) === 0);
  };
  const previous = table.previousElementSibling;
  if (!blank(previous)) return;
  const box = table.getBoundingClientRect(), observedCell = cell.getBoundingClientRect(), css = getComputedStyle(table);
  const top = observedCell.top - box.top, bottom = box.bottom - observedCell.bottom;
  const marginTop = parseFloat(css.marginTop) || 0, marginBottom = parseFloat(css.marginBottom) || 0;
  // A source blank can reserve its line with margins instead of a tall
  // content box. Check the previous painted paragraph and the blank itself.
  let before = previous.previousElementSibling;
  while (blank(before)) before = before.previousElementSibling;
  if (!(before instanceof HTMLParagraphElement) ||
      before.querySelector('img,svg,[data-note-reference],[data-document-textbox],[data-word-overlap]')) return;
  const paint = document.createRange(); paint.selectNodeContents(before);
  const previousBottom = Math.max(before.getBoundingClientRect().bottom,
    ...[...paint.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0).map(rect => rect.bottom));
  if (Math.abs(observedCell.height - reserve.observedHeight) > .12 || !(top > 0 && bottom > 0) ||
      bottom > marginBottom || box.top - top < previousBottom) return;
  table.style.marginTop = `${marginTop - top}px`;
  table.style.marginBottom = `${marginBottom - bottom}px`;
}

const continuedTextProperties = ['font-family', 'font-size', 'font-weight', 'font-style', 'letter-spacing',
  '--word-character-scale'] as const;

function signedTwips(raw: string, hanging = false) {
  if (!/^-?(?:\d+(?:\.\d+)?|\.\d+)(?:px|pt)$/.test(raw)) return raw === 'normal' ? '' : raw;
  const value = parseFloat(raw) * (raw.endsWith('pt') ? 20 : 15);
  return hanging && value < 0 ? -Math.round(-value) : Math.round(value);
}

function continuedTextRuns(paragraph: HTMLElement) {
  const spans = [...paragraph.querySelectorAll('span')];
  if (spans.some(span => span.getAttributeNames().some(name => name !== 'style') ||
    [...span.style].some(prop => prop !== 'color' && !continuedTextProperties.some(name => name === prop)))) return;
  const runs: { text: string; geometry: string }[] = [];
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const text = walker.currentNode.textContent!;
    if (!text) continue;
    const styles = continuedTextProperties.map(property => {
      for (let node = walker.currentNode.parentElement; node; node = node.parentElement) {
        const value = node.style.getPropertyValue(property);
        if (value || node === paragraph) return value;
      }
      return '';
    });
    const [family, size, weight, style, spacing, scale] = styles;
    const geometry = JSON.stringify([family.replace(/['"]/g, '').split(',').map(font => font.trim().toLowerCase()),
      signedTwips(size), weight, style, signedTwips(spacing), scale.trim() || '100']);
    const last = runs.at(-1);
    if (last?.geometry === geometry) last.text += text;
    else runs.push({ text, geometry });
  }
  return runs;
}

function continuedRowSignature(row: HTMLTableRowElement) {
  const cells = [...row.cells];
  if (cells.filter(cell => cell.rowSpan === 1).length < 3 || cells.some(cell =>
    cell.tagName !== 'TD' || cell.colSpan !== 1 || !cell.hasAttribute('colwidth') || !cell.children.length ||
    [...cell.children].some(child => child.tagName !== 'P') ||
    cell.querySelector('p :not(span),[id],[data-docx-bookmark],[data-docx-tab]'))) return;
  const twips = (value: string) => Math.round((cssPixels(value) ?? 0) * 15);
  const geometry = [];
  for (const cell of cells) {
    const { style } = cell;
    const padding = [style.paddingTop, style.paddingBottom, style.paddingLeft, style.paddingRight].map(twips);
    const paragraphs = [];
    for (const paragraph of cell.children) {
      const p = paragraph as HTMLElement, css = p.style;
      const runs = continuedTextRuns(p);
      if (!runs) return;
      // DOCX flattens nested text style spans and stores tracking and hanging
      // indents in twips. Fingerprint the effective runs in those same units.
      paragraphs.push([runs, css.fontFamily.replace(/['"]/g, '').toLowerCase(), twips(css.fontSize), twips(css.lineHeight),
        css.fontWeight, css.fontStyle, css.verticalAlign, css.whiteSpace, css.wordBreak, css.textAlign,
        [css.marginTop, css.marginBottom, css.marginLeft, css.marginRight].map(twips), signedTwips(css.textIndent, true)]);
    }
    geometry.push([cell.rowSpan, cell.getAttribute('colwidth'), padding, style.verticalAlign, paragraphs]);
  }
  return hwpWrapFingerprint(JSON.stringify(geometry));
}

/** Attach the plain cell geometry once on the genuine source-import path. */
export function sealHwpRowReserves(root: HTMLElement) {
  for (const row of root.querySelectorAll<HTMLTableRowElement>('tr[data-hwp-row-reserve]')) {
    const draft = row.dataset.hwpRowReserve;
    const inlineStroke = draft?.startsWith('3 ') === true;
    const continued = draft?.startsWith('4 ');
    const isContainer = inlineStroke || draft?.startsWith('2 '), cells = !isContainer && !continued && terminalCells(row);
    const proof = continued ? continuedRowSignature(row)
      : isContainer ? containerSignature(row, inlineStroke) : cells && signature(cells);
    if (proof && hwpRowReserve(`${draft} ${proof}`)) {
      row.dataset.hwpRowReserve = `${draft} ${proof}`;
      if (!isContainer) row.dataset.hwpRowContinued = 'true';
    } else if (!hwpRowReserve(draft)) delete row.dataset.hwpRowReserve;
  }
}

export function fittedHwpRowReserve(row: HTMLTableRowElement, available?: number): number | null {
  const reserve = hwpRowReserve(row.dataset.hwpRowReserve), minimum = cssPixels(row.style.height);
  if (!reserve || minimum === null || Math.abs(minimum - reserve.authoredHeight) > .1 || row.dataset.wordRowRule === 'exact') return null;
  if (reserve.kind === 4) {
    const table = row.closest('table'), cells = [...row.cells].filter(cell => cell.rowSpan === 1);
    const outerPadding = Math.max(0, ...cells.map(cell => cssPixels(cell.style.paddingBottom) ?? 0));
    return table?.dataset.hwpCenteredBordersActive === 'true' && table.dataset.hwpCenteredBorders === 'continuation' &&
      row.dataset.hwpRowContinued === 'true' && available !== undefined && reserve.headHeight !== undefined &&
      Math.abs(available - reserve.headHeight) <= outerPadding + .2 &&
      continuedRowSignature(row) === reserve.signature ? reserve.observedHeight : null;
  }
  if (reserve.kind !== 1) return null;
  const cells = terminalCells(row);
  if (!cells || signature(cells) !== reserve.signature || cells.some(cell =>
    [...cell.children].some(paragraph => paragraphLineOffsets(paragraph as HTMLElement).length > 0))) return null;
  return reserve.observedHeight;
}
