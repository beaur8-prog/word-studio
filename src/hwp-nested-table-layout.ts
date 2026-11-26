import { hwpWrapFingerprint, normalizeHwpWrapText } from './hwp-text-advances';

export const hwpNestedTableLayoutAttribute = 'data-hwp-nested-table-layout';
export const hwpNestedCellFaceAttribute = 'data-hwp-nested-cell-face';

export type HwpNestedTableGap = [block: number, after: number, before: number, tableTrailing: number | null];
export type HwpNestedTableMerge = [row: number, column: number, span: number];
export type HwpNestedTableCut = {
  block: number;
  rows: number;
  columns: number;
  boundary: number;
  fingerprint: string;
  merges: HwpNestedTableMerge[];
};
export type HwpNestedCellFace = {
  block: number;
  row: number;
  column: number;
  width: number;
  height: number;
  size: number;
  pitch: number;
  scale: number;
  fingerprint: string;
  ends: number[];
};
export type HwpNestedTableLayout = {
  blocks: number;
  fingerprint: string;
  gaps: HwpNestedTableGap[];
  cuts: HwpNestedTableCut[];
  faces?: HwpNestedCellFace[];
};

const integer = (n: unknown, lower: number, upper: number): n is number =>
  typeof n === 'number' && Number.isInteger(n) && n >= lower && n <= upper;
const decimal = (n: unknown, lower: number, upper: number): n is number =>
  typeof n === 'number' && Number.isFinite(n) && n >= lower && n <= upper;
const gapValue = (n: unknown): n is number =>
  typeof n === 'number' && Number.isFinite(n) && n >= 0 && n <= 200;

export function parseHwpNestedTableLayout(value: unknown): HwpNestedTableLayout | undefined {
  if (typeof value !== 'string' || value.length > 9000) return;
  try {
    const layout = JSON.parse(value) as HwpNestedTableLayout;
    if (!layout || !integer(layout.blocks, 4, 2048) || typeof layout.fingerprint !== 'string' || !/^[0-9a-f]{16}$/.test(layout.fingerprint) ||
      !Array.isArray(layout.gaps) || !layout.gaps.length || layout.gaps.length > layout.blocks - 1 ||
      !Array.isArray(layout.cuts) || !layout.cuts.length || layout.cuts.length > 64) return;
    if (!layout.gaps.every((g, i) => Array.isArray(g) && g.length === 4 &&
      integer(g[0], i ? layout.gaps[i - 1][0] + 1 : 0, layout.blocks - 2) &&
      gapValue(g[1]) && gapValue(g[2]) && g[1] + g[2] > 0 && (g[3] === null || gapValue(g[3])))) return;
    if (layout.faces !== undefined && (!Array.isArray(layout.faces) || layout.faces.length > 64 || !layout.faces.every(face =>
      parseHwpNestedCellFace(face, layout.blocks)))) return;
    if (!layout.cuts.every((cut, i) => cut && integer(cut.block, i ? layout.cuts[i - 1].block + 1 : 0, layout.blocks - 1) &&
      integer(cut.rows, 2, 1024) && integer(cut.columns, 1, 256) && integer(cut.boundary, 1, cut.rows - 1) &&
      typeof cut.fingerprint === 'string' && /^[0-9a-f]{16}$/.test(cut.fingerprint) &&
      Array.isArray(cut.merges) && cut.merges.length <= cut.columns &&
      cut.merges.every(merge => Array.isArray(merge) && merge.length === 3 && integer(merge[0], 0, cut.boundary - 1) &&
        integer(merge[1], 0, cut.columns - 1) && integer(merge[2], cut.boundary - merge[0] + 1, cut.rows - merge[0])))) return;
    return layout;
  } catch { return; }
}

export function parseHwpNestedCellFace(value: unknown, blocks = 2048): HwpNestedCellFace | undefined {
  if (!value || typeof value !== 'object') return;
  const face = value as HwpNestedCellFace;
  if (!integer(face.block, 0, blocks - 1) || !integer(face.row, 0, 1023) || !integer(face.column, 0, 255) ||
    !decimal(face.width, 12, 600) || !decimal(face.height, 6, 150) || !decimal(face.size, 6, 50) ||
    !decimal(face.pitch, face.size * .85, face.size * 1.02) || !integer(face.scale, 80, 99) ||
    typeof face.fingerprint !== 'string' || !/^[0-9a-f]{16}$/.test(face.fingerprint) ||
    !Array.isArray(face.ends) || face.ends.length < 1 || face.ends.length > 2 ||
    !face.ends.every((end, i) => integer(end, i ? face.ends[i - 1] + 1 : 5, 40))) return;
  return face;
}

export function hwpNestedParagraphToken(text: string) {
  const literal = normalizeHwpWrapText(text).replace(/\ufeff/g, '');
  return `p:${literal.length}:${hwpWrapFingerprint(literal)}`;
}

export function hwpNestedTableToken(rows: number, columns: number) {
  return `t:${rows}:${columns}`;
}

export function hwpNestedLayoutFingerprint(tokens: readonly string[]) {
  return hwpWrapFingerprint(tokens.join('|'));
}

export function hwpNestedCellToken(row: number, column: number, rowSpan: number, colSpan: number, text: string) {
  return `${row}:${column}:${rowSpan}:${colSpan}:${hwpNestedParagraphToken(text)}`;
}
