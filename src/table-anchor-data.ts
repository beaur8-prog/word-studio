export type TableAnchor = { x: number; y: number; bottom: number; reference?: 'paragraph'; horizontal?: 'margin'; overlay?: 'page-bottom' | 'paper-bottom'; wrap?: 'right'; right?: number };
export const paragraphTableNamespace = 'urn:wordstudio:paragraph-table:1';

/** Page-bottom overlays do not reserve flow height. Paragraph labels can
 * wrap text on the right; other anchors retain above/below flow semantics. */
export function tableAnchor(value: unknown): TableAnchor | undefined {
  try {
    const a = typeof value === 'string' ? JSON.parse(value) : value;
    if (a && [a.x, a.y, a.bottom].every(Number.isFinite) && a.x >= 0 && a.x <= 5000 && a.y >= 0 && a.y <= 5000 && a.bottom >= 0 && a.bottom <= 1000 &&
      (a.reference === undefined || a.reference === 'paragraph') && (a.horizontal === undefined || a.horizontal === 'margin') &&
      (a.overlay === undefined || ['page-bottom', 'paper-bottom'].includes(a.overlay) && !a.reference && !a.horizontal) &&
      (a.wrap === undefined || a.wrap === 'right' && a.reference === 'paragraph' && a.horizontal === 'margin' && !a.overlay && a.y === 0 && Number.isFinite(a.right) && a.right >= 0 && a.right <= 1000))
      return { x: a.x, y: a.y, bottom: a.bottom, ...(a.reference ? { reference: a.reference } : {}), ...(a.horizontal ? { horizontal: a.horizontal } : {}), ...(a.overlay ? { overlay: a.overlay } : {}), ...(a.wrap ? { wrap: a.wrap, right: a.right } : {}) };
  } catch { /* Unsupported floating settings retain ordinary flow. */ }
}

/** Reserve a usable text strip consistently in editor, print and DOCX. */
export function tableWrapFits(anchor: TableAnchor | undefined, width: number, bodyWidth: number): boolean {
  return !!anchor?.wrap && [width, bodyWidth].every(Number.isFinite) && width > 0 &&
    anchor.x + width + anchor.right! <= bodyWidth - 32;
}
