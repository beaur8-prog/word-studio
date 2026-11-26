import { splitText } from './paragraph-fragments';
import { hwpWrapFingerprint, normalizeHwpWrapText } from './hwp-text-advances';

export const hwpCenteredCellLabelsAttribute = 'data-hwp-centered-cell-labels';
export const hwpCenteredCellLabelsNamespace = 'urn:wordstudio:hwp-centered-cell-labels:1';

type PeerCut = { fingerprint: string; paragraph: number; offset: number };
type CompactFragment = { height: number; peerWidth: number; peerPrefix: string; cut?: PeerCut };
type CenteredLabels = { width: number; fingerprint: string; firstFragment?: true; compact?: CompactFragment };

export const hwpNativeCompactRowAttribute = 'data-print-hwp-compact-row';

export function parseHwpCenteredCellLabels(value: unknown): CenteredLabels | undefined {
  if (typeof value !== 'string' || value.length > 110) return;
  const values = value.split(' ');
  const [width, fingerprint, first, height, peerWidth, peerPrefix, peerText, paragraph, offset] = values;
  const positive = (text: string) => /^\d{1,4}(?:\.\d{1,2})?$/.test(text) && Number(text) > 0;
  const digest = (text: string) => /^[a-f0-9]{16}$/.test(text);
  if (!positive(width) || !digest(fingerprint)) return;
  const base = { width: Number(width), fingerprint };
  if (values.length === 2) return base;
  if (first !== 'first') return;
  const named = { ...base, firstFragment: true as const };
  if (values.length === 3) return named;
  if (![6, 9].includes(values.length) || !positive(height) || Number(height) > 2000 ||
      !positive(peerWidth) || !digest(peerPrefix)) return;
  const compact = { height: Number(height), peerWidth: Number(peerWidth), peerPrefix };
  if (values.length === 6) return { ...named, compact };
  if (!digest(peerText) || !/^\d{1,2}$/.test(paragraph) || Number(paragraph) >= 24 ||
      !/^\d{1,4}$/.test(offset) || !(Number(offset) > 0)) return;
  return { ...named, compact: { ...compact, cut: { fingerprint: peerText, paragraph: Number(paragraph), offset: Number(offset) } } };
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

/** A different font renderer may fit more characters on the retained cell
 * line. Respect its authored boundary only within the same split paragraph. */
export function restoreHwpCompactPeerSplit(label: HTMLTableCellElement, original: HTMLTableCellElement,
  head: HTMLTableCellElement, tail: HTMLTableCellElement) {
  const proof = parseHwpCenteredCellLabels(label.getAttribute(hwpCenteredCellLabelsAttribute))?.compact;
  const cut = proof?.cut;
  if (!cut || Math.abs(head.getBoundingClientRect().width - proof.peerWidth) > .75 ||
      head.children.length !== cut.paragraph + 1 || !tail.firstElementChild ||
      head.children.length + tail.children.length !== original.children.length + 1 ||
      original.querySelector('br,img,svg,table,[data-note-reference],[data-docx-tab],[data-document-field],[data-docx-bookmark],[data-symbol-font]')) return;
  const literal = [...original.children].map(p => p.textContent ?? '');
  const whole = hwpWrapFingerprint(JSON.stringify(literal.map(normalizeHwpWrapText)));
  const para = original.children[cut.paragraph];
  const retained = head.lastElementChild, following = tail.firstElementChild;
  if (!(para instanceof HTMLParagraphElement) || !(retained instanceof HTMLParagraphElement) ||
      !(following instanceof HTMLParagraphElement) || whole !== cut.fingerprint) return;
  const prefix = retained.textContent ?? '', suffix = following.textContent ?? '', full = literal[cut.paragraph];
  if (prefix + suffix !== full || !(cut.offset < prefix.length) || cut.offset <= 0 ||
      /[\uD800-\uDFFF]/u.test(full) || hwpWrapFingerprint(normalizeHwpWrapText(
        literal.slice(0, cut.paragraph).join('') + full.slice(0, cut.offset))) !== proof.peerPrefix) return;
  const [sourceHead, sourceTail] = splitText(para, cut.offset);
  retained.replaceChildren(...sourceHead.childNodes);
  following.replaceChildren(...sourceTail.childNodes);
}
