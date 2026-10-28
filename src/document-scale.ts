/** The editor uses uniform CSS transforms for document zoom. offsetWidth is
 * rounded to whole pixels, so rect.width / offsetWidth invents a small zoom
 * even at 100% on fractional paper widths; that error grows down long files. */
export function documentScale(element: HTMLElement): number {
  let scale = 1;
  for (let current: HTMLElement | null = element; current; current = current.parentElement) {
    const style = getComputedStyle(current);
    const zoom = Number.parseFloat(style.zoom);
    if (Number.isFinite(zoom) && zoom > 0) scale *= zoom;
    if (style.transform !== 'none') {
      const matrix = new DOMMatrixReadOnly(style.transform);
      scale *= Math.hypot(matrix.a, matrix.b);
    }
  }
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}
