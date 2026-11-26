export const hwpNativeGridAttribute = 'data-hwp-native-grid';

type CellWidth = [column: number, span: number, width: number];
type NativeGrid = [width: number, columns: number[], widths: CellWidth[]];

/** The standard DOCX grid is rounded to CSS pixels in the editor. Retain
 * native cell spans only when their exact widths agree with that same grid. */
export function parseHwpNativeGrid(value: unknown): NativeGrid | undefined {
  if (typeof value !== 'string' || value.length > 4096) return;
  try {
    const stored = JSON.parse(value);
    if (!Array.isArray(stored) || stored.length !== 3) return;
    const [width, columns, widths] = stored as NativeGrid;
    const validWidth = (size: number) => Number.isSafeInteger(size) && size > 0 && size <= 300000;
    if (!validWidth(width) || !Array.isArray(columns) || columns.length < 2 || columns.length > 64 ||
        columns.some(col => !Number.isSafeInteger(col) || col <= 0 || col > 4000) ||
        Math.abs(columns.reduce((sum, col) => sum + col, 0) * 75 - width) > columns.length * 38 ||
        !Array.isArray(widths) || !widths.length || widths.length > 128) return;
    const seen = new Set<string>();
    for (const cell of widths) {
      if (!Array.isArray(cell) || cell.length !== 3) return;
      const [column, span, size] = cell;
      if (!Number.isSafeInteger(column) || !Number.isSafeInteger(span) || column < 0 || span < 1 ||
          column + span > columns.length || !validWidth(size)) return;
      const key = column + ':' + span;
      if (seen.has(key) || Math.abs(columns.slice(column, column + span).reduce((sum, col) => sum + col, 0) * 75 - size) > span * 38) return;
      seen.add(key);
    }
    return [width, columns, widths];
  } catch { /* Keep ordinary grid widths for unsupported properties. */ }
}

export function hwpNativeGrid(width: number, columns: number[],
  cells: readonly { col: number; colSpan: number; rowSpan: number; width: number }[]): string | undefined {
  if (columns.length > 64 || !cells.some(cell => cell.rowSpan > 1) || !cells.some(cell => cell.colSpan > 1)) return;
  const spans = new Map<string, CellWidth>();
  const explicit = new Set<number>();
  for (const cell of cells) {
    const key = cell.col + ':' + cell.colSpan;
    const previous = spans.get(key);
    if (previous && previous[2] !== cell.width) return;
    spans.set(key, [cell.col, cell.colSpan, cell.width]);
    if (cell.colSpan === 1) explicit.add(cell.col);
    if (spans.size > 128) return;
  }
  if (explicit.size === columns.length) return;
  const value = JSON.stringify([width, columns.map(Math.round), [...spans.values()]]);
  return parseHwpNativeGrid(value) && value;
}
