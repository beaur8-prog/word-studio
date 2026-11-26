import { hwpWrapFingerprint, normalizeHwpWrapText } from './hwp-text-advances';

export const hwpCenteredCellLabelsAttribute = 'data-hwp-centered-cell-labels';
export const hwpCenteredCellLabelsNamespace = 'urn:wordstudio:hwp-centered-cell-labels:1';

type CompactFragment = { height: number; peerWidth: number; peerPrefix: string };
type CenteredLabels = { width: number; fingerprint: string; firstFragment?: true; compact?: CompactFragment };

export const hwpNativeCompactRowAttribute = 'data-print-hwp-compact-row';

export function parseHwpCenteredCellLabels(value: unknown): CenteredLabels | undefined {
  if (typeof value !== 'string' || value.length > 78) return;
  const match = /^(\d{1,4}(?:\.\d{1,2})?) ([a-f0-9]{16})( first(?: (\d{1,4}(?:\.\d{1,2})?) (\d{1,4}(?:\.\d{1,2})?) ([a-f0-9]{16}))?)?$/.exec(value);
  const width = Number(match?.[1]);
  if (!match || !(width > 0)) return;
  const height = Number(match[4]), peerWidth = Number(match[5]);
  if (match[6] && (!(height > 0 && height <= 2000) || !(peerWidth > 0))) return;
  return { width, fingerprint: match[2], ...(match[3] && { firstFragment: true as const }),
    ...(match[6] && { compact: { height, peerWidth, peerPrefix: match[6] } }) };
}

export function matchesHwpCenteredCellLabels(cell: HTMLTableCellElement) {
  const proof = parseHwpCenteredCellLabels(cell.getAttribute(hwpCenteredCellLabelsAttribute));
  return proof && cell.tagName === 'TD' && cell.colSpan === 1 && cell.rowSpan === 1 &&
    cell.style.verticalAlign === 'middle' && cell.children.length >= 1 && cell.children.length <= 2 &&
    (!proof.firstFragment || cell.children.length === (proof.compact ? 2 : 1)) &&
    !cell.closest('table')?.style.zoom && !cell.querySelector('img,svg,table,br,[data-note-reference],[data-docx-tab],[data-document-field]') &&
    [...cell.children].every(child => child.tagName === 'P') &&
    Math.abs(cell.getBoundingClientRect().width - proof.width) <= .75 &&
    hwpWrapFingerprint(JSON.stringify([...cell.children].map(child => child.textContent ?? ''))) === proof.fingerprint;
}

export function matchesHwpCenteredMergeHeadLabel(cell: HTMLTableCellElement, firstOnly = false) {
  return !!matchesHwpCenteredCellLabels(cell) && (!firstOnly && cell.children.length === 2 ||
    parseHwpCenteredCellLabels(cell.getAttribute(hwpCenteredCellLabelsAttribute))?.firstFragment === true);
}

export function hwpCenteredLabelInkFits(cell: HTMLTableCellElement) {
  const bounds = cell.getBoundingClientRect(), range = document.createRange();
  return [...cell.children].every(child => {
    range.selectNodeContents(child);
    const ink = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    // Separate lines cannot share a vertical band. A later wrapping edit
    // should take the ordinary fragmentation path even if the cell is tall.
    return ink.length > 0 && Math.max(...ink.map(rect => rect.top)) < Math.min(...ink.map(rect => rect.bottom)) &&
      ink.every(rect => rect.top >= bounds.top - .5 && rect.bottom <= bounds.bottom + .5 &&
        rect.left >= bounds.left - .5 && rect.right <= bounds.right + .5);
  });
}

/** Preserve a witnessed short row only on the finished first print fragment.
 * The tentative native result is checked again after patching the cell grid. */
export function markHwpNativeCompactRow(table: HTMLTableElement, body: Element) {
  const row = table.rows[table.rows.length - 1];
  if (!row || row.dataset.hwpRowContinued !== 'true' || row.cells.length !== 2 ||
      table.parentElement !== body || table === body.firstElementChild || table.style.zoom ||
      !table.hasAttribute('data-table-anchor') ||
      ![...table.rows].slice(0, -1).some((prior, index) =>
        [...prior.cells].some(cell => cell.rowSpan === table.rows.length - index && cell.colSpan === 1))) return;
  const [label, peer] = [...row.cells];
  if (peer.rowSpan !== 1 || peer.colSpan !== 1 || peer.querySelector('img,svg,table,[data-note-reference],[data-docx-tab],[data-document-field]')) return;
  const proof = parseHwpCenteredCellLabels(label.getAttribute(hwpCenteredCellLabelsAttribute))?.compact;
  if (!proof || !matchesHwpCenteredCellLabels(label) ||
      Math.abs(peer.getBoundingClientRect().width - proof.peerWidth) > .75 ||
      hwpWrapFingerprint(normalizeHwpWrapText(peer.textContent ?? '')) !== proof.peerPrefix) return;
  const bounds = row.getBoundingClientRect(), frame = body.getBoundingClientRect();
  if (bounds.top < frame.top || bounds.bottom > frame.bottom + .5 ||
      !(proof.height < bounds.height - .5)) return;
  row.setAttribute(hwpNativeCompactRowAttribute, String(proof.height));
}
