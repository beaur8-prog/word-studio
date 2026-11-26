import { hwpWrapFingerprint, normalizeHwpWrapText } from './hwp-text-advances';

export const hwpRelatedTableLayoutAttribute = 'data-hwp-related-table-layout';

type ParagraphBefore = { gap: number; text: string; font: number; pitch: number; breaks: number };
type RelatedInlineTable = { left: number; width: number; rows: number; text: string };

export type HwpRelatedTableLayout = {
  bodyWidth: number;
  bodyHeight: number;
  rows: number;
  row: number;
  rowWidths: [number, number];
  before: [ParagraphBefore[], ParagraphBefore[]];
  pair: [RelatedInlineTable, RelatedInlineTable];
};

const inRange = (value: unknown, min: number, max: number): value is number =>
  typeof value === 'number' && Number.isFinite(value) && value >= min && value <= max;
const integer = (value: unknown, min: number, max: number): value is number =>
  inRange(value, min, max) && Number.isInteger(value);
const token = (value: unknown) => typeof value === 'string' && /^[a-f0-9]{16}$/.test(value);

export function parseHwpRelatedTableLayout(raw: unknown): HwpRelatedTableLayout | undefined {
  if (typeof raw !== 'string' || raw.length > 6500) return;
  try {
    const proof = JSON.parse(raw) as HwpRelatedTableLayout;
    if (!proof || !inRange(proof.bodyWidth, 100, 2500) || !inRange(proof.bodyHeight, 100, 2500) ||
      !integer(proof.rows, 3, 2000) || !integer(proof.row, 1, proof.rows - 2) ||
      !Array.isArray(proof.rowWidths) || proof.rowWidths.length !== 2 ||
      !proof.rowWidths.every(width => inRange(width, 12, proof.bodyWidth)) ||
      !Array.isArray(proof.before) || proof.before.length !== 2 ||
      !proof.before.every(cell => Array.isArray(cell) && integer(cell.length, 2, 48) && cell.every(p =>
        p && inRange(p.gap, .2, 80) && token(p.text) && inRange(p.font, 6, 72) &&
        inRange(p.pitch, p.font, p.font * 4) && integer(p.breaks, 0, 15))) ||
      !Array.isArray(proof.pair) || proof.pair.length !== 2 || !proof.pair.every(table => table &&
        integer(table.rows, 1, 48) && inRange(table.left, .5, proof.bodyWidth) &&
        inRange(table.width, 12, proof.bodyWidth - table.left - .5) && token(table.text))) return;
    const [left, right] = proof.pair;
    if (right.left - left.left - left.width < .5) return;
    return proof;
  } catch { return; }
}

export function hwpRelatedParagraphToken(text: string) {
  return hwpWrapFingerprint(normalizeHwpWrapText(text).replace(/\ufeff/g, ''));
}

export function hwpRelatedTableToken(paragraphs: readonly string[]) {
  return hwpWrapFingerprint(paragraphs.filter(text => normalizeHwpWrapText(text).replace(/\ufeff/g, '')).map(hwpRelatedParagraphToken).join('|'));
}
