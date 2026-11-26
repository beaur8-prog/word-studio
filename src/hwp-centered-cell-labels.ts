import { hwpWrapFingerprint } from './hwp-text-advances';

export const hwpCenteredCellLabelsAttribute = 'data-hwp-centered-cell-labels';
export const hwpCenteredCellLabelsNamespace = 'urn:wordstudio:hwp-centered-cell-labels:1';

export function parseHwpCenteredCellLabels(value: unknown): { width: number; fingerprint: string; firstFragment?: true } | undefined {
  if (typeof value !== 'string' || value.length > 34) return;
  const match = /^(\d{1,4}(?:\.\d{1,2})?) ([a-f0-9]{16})( first)?$/.exec(value);
  const width = Number(match?.[1]);
  if (match && width > 0) return { width, fingerprint: match[2], ...(match[3] ? { firstFragment: true as const } : {}) };
}

export function matchesHwpCenteredCellLabels(cell: HTMLTableCellElement) {
  const proof = parseHwpCenteredCellLabels(cell.getAttribute(hwpCenteredCellLabelsAttribute));
  return proof && cell.tagName === 'TD' && cell.colSpan === 1 && cell.rowSpan === 1 &&
    cell.style.verticalAlign === 'middle' && cell.children.length >= 1 && cell.children.length <= 2 && (!proof.firstFragment || cell.children.length === 1) &&
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
