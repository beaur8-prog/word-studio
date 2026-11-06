/** Quantize shared column edges once. Rounding each cell independently loses
 * up to one pixel per column and disagrees with cells spanning those columns. */
export function tablePixelGrid(widths: readonly number[], requested: number, available: number): number[] {
  const total = widths.reduce((sum, width) => sum + width, 0);
  if (!widths.length || !(total > 0) || !widths.every(width => Number.isFinite(width) && width > 0) ||
    !Number.isFinite(requested) || !Number.isFinite(available)) return [];
  // The editor requires at least one pixel per column. An impossible grid
  // retains that minimum, rather than producing zero or negative columns.
  const target = Math.max(widths.length, Math.min(Math.round(requested), Math.floor(available + 1e-6)));
  let sourceEdge = 0, edge = 0;
  return widths.map((width, index) => {
    sourceEdge += width;
    const next = Math.max(edge + 1, Math.min(target - (widths.length - index - 1), Math.round(sourceEdge * target / total)));
    const result = next - edge; edge = next; return result;
  });
}
