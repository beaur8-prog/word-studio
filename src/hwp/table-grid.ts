/** Resolve cumulative source edges before dividing merged cells. Treating each
 * colspan as equal columns destroys the narrow label grids used by HWP forms. */
export function hwpTableGrid(cells: Record<string, number>[], axis: 'col' | 'row', length: number): number[] {
  const edges: (number | undefined)[] = Array(length + 1); edges[0] = 0;
  for (let pass = 0; pass < 4; pass++) for (const cell of cells) {
    const start = cell[axis], end = start + cell[axis + 'Span'], size = cell[axis === 'col' ? 'width' : 'height'];
    if (edges[start] !== undefined && edges[end] === undefined) edges[end] = edges[start]! + size;
    else if (edges[end] !== undefined && edges[start] === undefined) edges[start] = Math.max(0, edges[end]! - size);
  }
  for (let i = 1; i <= length; i++) if (edges[i] === undefined) {
    let end = i + 1; while (end <= length && edges[end] === undefined) end++;
    const step = end <= length ? (edges[end]! - edges[i - 1]!) / (end - i + 1) : axis === 'col' ? 6000 : 1000;
    for (let j = i; j < end && j <= length; j++) edges[j] = edges[i - 1]! + step * (j - i + 1);
    i = end - 1;
  }
  return Array.from({ length }, (_, i) => Math.max(1, Math.round((edges[i + 1]! - edges[i]!) / 75 * 100) / 100));
}
