export type TableAnchor = { x: number; y: number; bottom: number; reference?: 'paragraph'; horizontal?: 'margin'; overlay?: 'page-bottom' | 'paper-bottom' };
export const paragraphTableNamespace = 'urn:wordstudio:paragraph-table:1';

/** Page-bottom overlays never reserve flow height. Other supported anchors
 * retain their existing above/below flow semantics. */
export function tableAnchor(value: unknown): TableAnchor | undefined {
  try {
    const a = typeof value === 'string' ? JSON.parse(value) : value;
    if (a && [a.x, a.y, a.bottom].every(Number.isFinite) && a.x >= 0 && a.x <= 5000 && a.y >= 0 && a.y <= 5000 && a.bottom >= 0 && a.bottom <= 1000 &&
      (a.reference === undefined || a.reference === 'paragraph') && (a.horizontal === undefined || a.horizontal === 'margin') &&
      (a.overlay === undefined || ['page-bottom', 'paper-bottom'].includes(a.overlay) && !a.reference && !a.horizontal))
      return { x: a.x, y: a.y, bottom: a.bottom, ...(a.reference ? { reference: a.reference } : {}), ...(a.horizontal ? { horizontal: a.horizontal } : {}), ...(a.overlay ? { overlay: a.overlay } : {}) };
  } catch { /* Unsupported floating settings retain ordinary flow. */ }
}
