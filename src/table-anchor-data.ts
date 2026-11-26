export type TableAnchor = { x: number; y: number; bottom: number; reference?: 'paragraph'; horizontal?: 'margin'; horizontalAlign?: 'center'; overlay?: 'page-bottom' | 'paper-bottom'; wrap?: 'right'; right?: number };
export const paragraphTableNamespace = 'urn:wordstudio:paragraph-table:1';

export function hwpTableTopGap(value: unknown): number | undefined {
  if (typeof value !== 'string' && typeof value !== 'number') return;
  const gap = Number(value);
  return Number.isFinite(gap) && gap > 0 && gap <= 1000 ? gap : undefined;
}

/** Page-bottom overlays do not reserve flow height. Paragraph labels can
 * wrap text on the right; other anchors retain above/below flow semantics. */
export function tableAnchor(value: unknown): TableAnchor | undefined {
  try {
    const a = typeof value === 'string' ? JSON.parse(value) : value;
    if (a && [a.x, a.y, a.bottom].every(Number.isFinite) && a.x >= 0 && a.x <= 5000 && a.y >= 0 && a.y <= 5000 && a.bottom >= 0 && a.bottom <= 1000 &&
      (a.reference === undefined || a.reference === 'paragraph') && (a.horizontal === undefined || a.horizontal === 'margin') &&
      (a.horizontalAlign === undefined || a.horizontalAlign === 'center' && a.horizontal === 'margin' && a.x === 0 && !a.reference && !a.overlay && !a.wrap) &&
      (a.overlay === undefined || ['page-bottom', 'paper-bottom'].includes(a.overlay) && !a.reference && !a.horizontal) &&
      (a.wrap === undefined || a.wrap === 'right' && a.reference === 'paragraph' && a.horizontal === 'margin' && !a.overlay && a.y === 0 && Number.isFinite(a.right) && a.right >= 0 && a.right <= 1000))
      return { x: a.x, y: a.y, bottom: a.bottom, ...(a.reference ? { reference: a.reference } : {}), ...(a.horizontal ? { horizontal: a.horizontal } : {}),
        ...(a.horizontalAlign ? { horizontalAlign: a.horizontalAlign } : {}), ...(a.overlay ? { overlay: a.overlay } : {}), ...(a.wrap ? { wrap: a.wrap, right: a.right } : {}) };
  } catch { /* Unsupported floating settings retain ordinary flow. */ }
}

/** Reserve a usable text strip consistently in editor, print and DOCX. */
export function tableWrapFits(anchor: TableAnchor | undefined, width: number, bodyWidth: number): boolean {
  return !!anchor?.wrap && [width, bodyWidth].every(Number.isFinite) && width > 0 &&
    anchor.x + width + anchor.right! <= bodyWidth - 32;
}


/** Physical page of a witnessed, intact HWP inline table. */
export function hwpKeptTablePage(value: unknown) {
  const page = typeof value === 'number' ? value : typeof value === 'string' && /^\d{1,5}$/.test(value) ? Number(value) : NaN;
  return Number.isSafeInteger(page) && page > 1 && page <= 10000 ? page : undefined;
}

/** HWP inline object margins, in native HWPUNIT. Zero is explicit. */
export const hwpInlineTableBottomAttribute = 'data-hwp-inline-table-bottom';
export function hwpInlineTableBottom(value: unknown) {
  const margin = hwpInlineTableTop(value);
  return margin !== undefined && margin < 283 ? margin : undefined;
}

export const hwpInlineTableTopAttribute = 'data-hwp-inline-table-top';
export function hwpInlineTableTop(value: unknown): number | undefined {
  const top = typeof value === 'number' ? value
    : typeof value === 'string' && /^\d{1,5}$/.test(value) ? Number(value) : NaN;
  return Number.isSafeInteger(top) && top >= 0 && top <= 32767 ? top : undefined;
}
