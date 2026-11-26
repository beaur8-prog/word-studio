import {
  hwpRelatedParagraphToken, hwpRelatedTableLayoutAttribute, hwpRelatedTableToken, parseHwpRelatedTableLayout,
  type HwpRelatedTableLayout,
} from './hwp-related-table-layout';
import { pageMargins, paperDimensions, type PageSettings } from './types';

const close = (a: number, b: number, tolerance = .75) => Math.abs(a - b) <= tolerance;
const gridWidths = (table: HTMLTableElement) =>
  Array.from(table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col'), col => parseFloat(col.style.width));

function pixels(value: string) {
  return /^(?:\d+\.?\d*|\.\d+)(?:px|pt)$/.test(value) ? parseFloat(value) * (value.endsWith('pt') ? 4 / 3 : 1) : NaN;
}

function isEmptyCarrier(element: Element | null): element is HTMLParagraphElement {
  return element instanceof HTMLParagraphElement && !element.textContent?.trim() &&
    !element.querySelector('br,img,svg,table,[data-page-number],[data-note-reference],[data-docx-tab],[data-docx-bookmark]');
}

function localPage(table: HTMLTableElement, initialPage: PageSettings) {
  for (let previous = table.previousElementSibling; previous; previous = previous.previousElementSibling) {
    if (previous.hasAttribute('data-section-page')) {
      try { return JSON.parse(previous.getAttribute('data-section-page')!) as PageSettings; }
      catch { return; }
    }
  }
  return initialPage;
}

function stillMatchesPair(table: HTMLTableElement, member: HwpRelatedTableLayout['pair'][number]) {
  const [width, other] = gridWidths(table);
  if (table.rows.length !== member.rows || other !== undefined || !close(width, member.width) ||
    !close(pixels(table.style.marginLeft), member.left) ||
    table.matches('[data-table-anchor],[data-hwp-kept-table-page],[data-word-table-width-auto]') ||
    table.querySelector('table')) return false;
  const paragraphs: string[] = [];
  for (const [cell] of Array.from(table.rows, row => [...row.cells])) {
    if (!cell || cell.rowSpan !== 1 || cell.colSpan !== 1) return false;
    paragraphs.push(...Array.from(cell.children).filter(child => child.matches('p,h1,h2,h3,h4,h5,h6')).map(p => p.textContent ?? ''));
  }
  return [...table.rows].every(row => row.cells.length === 1) && hwpRelatedTableToken(paragraphs) === member.text;
}

function rowMargins(table: HTMLTableElement, proof: HwpRelatedTableLayout) {
  const widths = gridWidths(table);
  const row = table.rows[proof.row];
  if (table.rows.length !== proof.rows || widths.length !== 2 || widths.some((width, col) => !close(width, proof.rowWidths[col])) ||
    row?.cells.length !== 2) return;
  const changes = new Map<HTMLParagraphElement, number>();
  for (const [column, cell] of Array.from(row.cells).entries()) {
    const before = proof.before[column];
    if (cell.rowSpan !== 1 || cell.colSpan !== 1 || cell.children.length !== before.length) return;
    for (const [index, node] of Array.from(cell.children).entries()) {
      const { text, gap, font, pitch, breaks } = before[index];
      if (!(node instanceof HTMLParagraphElement) || hwpRelatedParagraphToken(node.textContent ?? '') !== text ||
        node.querySelectorAll('br').length !== breaks ||
        node.querySelector('img,svg,table,[data-page-number],[data-note-reference],[data-docx-tab],[data-docx-bookmark]') ||
        (parseFloat(node.style.marginTop) || 0) !== 0 || (parseFloat(node.style.marginBottom) || 0) !== 0 ||
        (!close(pixels(node.style.lineHeight), pitch, .12) &&
          (index !== before.length - 1 || !close(pixels(node.style.lineHeight), font, .12)))) return;
      changes.set(node, gap);
    }
  }
  return changes;
}

/** Organize only the derived print tree. The two editable tables, their
 * widths and their blank spacing carriers remain ordinary document nodes. */
export function layoutHwpRelatedTables(source: HTMLElement, initialPage: PageSettings) {
  for (const left of source.querySelectorAll<HTMLTableElement>(`table[${hwpRelatedTableLayoutAttribute}]`)) {
    if (left.closest('td,th,[data-document-note],[data-document-header],[data-document-footer],[data-column-count]')) continue;
    const proof = parseHwpRelatedTableLayout(left.getAttribute(hwpRelatedTableLayoutAttribute));
    if (!proof) continue;
    const settings = localPage(left, initialPage);
    if (!settings || (settings.columns?.count ?? 1) !== 1) continue;
    const [width, height] = paperDimensions(settings), margin = pageMargins(settings);
    const bodyWidth = (width - margin.left - margin.right) * 96 / 25.4;
    const bodyHeight = (height - margin.top - margin.bottom) * 96 / 25.4;
    if (!close(bodyWidth, proof.bodyWidth, .2) || !close(bodyHeight, proof.bodyHeight, .2)) continue;
    const carrier = left.nextElementSibling, previousCarrier = left.previousElementSibling;
    const right = carrier?.nextElementSibling;
    const preceding = isEmptyCarrier(previousCarrier) ? previousCarrier.previousElementSibling : previousCarrier;
    if (!(right instanceof HTMLTableElement) || !(preceding instanceof HTMLTableElement) ||
      !isEmptyCarrier(carrier) || !carrier.hasAttribute('data-exact-line-height') ||
      !stillMatchesPair(left, proof.pair[0]) || !stillMatchesPair(right, proof.pair[1])) continue;
    const margins = rowMargins(preceding, proof);
    if (!margins) continue;

    // Preserve the validated logical input's rounded table grid and indent.
    // Even the fraction of a twip rounded during the actual saves can move a
    // final glyph to a different line in the derived table.
    const leftInset = pixels(left.style.marginLeft), rightInset = pixels(right.style.marginLeft);
    const [leftWidth] = gridWidths(left), [rightWidth] = gridWidths(right);
    const widths = [leftInset, leftWidth, rightInset - leftInset - leftWidth, rightWidth, bodyWidth - rightInset - rightWidth];
    if (widths.some(value => !(value > 0))) continue;
    const group = document.createElement('table');
    group.dataset.hwpRelatedPrintGroup = 'true';
    group.style.cssText = `width:${bodyWidth}px;table-layout:fixed;border:0;margin:0;vertical-align:top;border-collapse:collapse`;
    const columns = group.appendChild(document.createElement('colgroup'));
    const row = group.createTBody().insertRow();
    row.dataset.rowSplit = 'false';
    left.before(group);
    for (const [index, width] of widths.entries()) {
      columns.appendChild(document.createElement('col')).style.width = width + 'px';
      const cell = row.insertCell();
      cell.style.cssText = 'padding:0;border:0;vertical-align:top';
      const table = index === 1 ? left : index === 3 ? right : undefined;
      if (table) {
        table.dataset.hwpRelatedInlineMember = 'true';
        table.style.margin = '0';
        cell.append(table);
      }
    }
    preceding.dataset.hwpRelatedRowFlow = 'true';
    for (const [paragraph, gap] of margins) paragraph.style.marginTop = gap + 'px';
  }
}
