import { cssPixels } from './hwp-cell-line';
import { hwpWrapFingerprint } from './hwp-text-advances';

export const hwpIntactRowAttribute = 'data-hwp-intact-top-row';

export type HwpIntactRow = {
  height: number;
  widths: number[];
  fingerprint: string;
};

export function parseHwpIntactRow(value: unknown): HwpIntactRow | undefined {
  if (typeof value !== 'string' || value.length > 220) return;
  try {
    const { height, widths, fingerprint } = JSON.parse(value);
    if (Number.isFinite(height) && height > 0 && height <= 2000 && Array.isArray(widths) &&
        widths.length >= 2 && widths.length <= 6 &&
        widths.every(width => Number.isFinite(width) && width > 0 && width <= 10000) &&
        typeof fingerprint === 'string' && /^[0-9a-f]{16}$/.test(fingerprint)) return { height, widths, fingerprint };
  } catch { /* The row can still use its ordinary split permission. */ }
}

/** The original row was intact. Only use that placement while the editable
 * row still has the same text and fits the measured source geometry. */
export function matchesHwpIntactTopRow(row: HTMLTableRowElement, height: number, scale: number) {
  const proof = parseHwpIntactRow(row.getAttribute(hwpIntactRowAttribute));
  const minimum = cssPixels(row.style.height);
  if (!proof || row.dataset.rowSplit !== 'allow' || row.dataset.wordRowRule === 'exact' || minimum === null ||
      Math.abs(height - proof.height) > .2 || Math.abs(minimum - proof.height) > .16 ||
      row.cells.length !== proof.widths.length ||
      row.querySelector('img,table,svg,br:not(.ProseMirror-trailingBreak),[data-note-reference],[data-document-field],[data-document-textbox]')) return false;
  const text: string[][] = [];
  for (const [index, cell] of [...row.cells].entries()) {
    if (cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
        Math.abs(cell.getBoundingClientRect().width / scale - proof.widths[index]) > .75 ||
        cell.style.verticalAlign !== 'top') return false;
    const paragraphs = [...cell.children].filter(block => !block.classList.contains('ProseMirror-widget'));
    if (!paragraphs.length || paragraphs.some(block => block.tagName !== 'P')) return false;
    text.push(paragraphs.map(block => block.textContent ?? ''));
  }
  return hwpWrapFingerprint(JSON.stringify(text)) === proof.fingerprint;
}
