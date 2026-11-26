import type { HwpDocument } from '@rhwp/core';
import { hwpShortCellLabel, hwpWrapFingerprint, normalizeHwpWrapText } from '../hwp-text-advances';
import type { NativeCellInkRun } from './source-caption-cell-line';
import { sourceCellBoundsTolerance, type HwpPaintedLine } from './paragraph-layout';

export function sourceCenteredCellLabels(
  doc: HwpDocument, section: number, paragraph: number, controlIndex: number, index: number,
  first: { pageIndex: number; y: number; h: number; w: number },
  lines: Map<string, HwpPaintedLine[]>, merged = false,
): string | undefined {
  try {
    const count = doc.getCellParagraphCount(section, paragraph, controlIndex, index);
    if (count !== 2 && (count !== 1 || merged)) return;
    const text: string[] = [], ink: HwpPaintedLine[] = [];
    for (let p = 0; p < count; p++) {
      const path = JSON.stringify([{ controlIndex, cellIndex: index, cellParaIndex: p }]);
      const literal = doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
      // Spaces belong to the fingerprint even though they do not classify the label.
      const label = count === 1 ? literal.replace(/^\d{1,3} /, '') : literal;
      if (!hwpShortCellLabel(label.replace(/ /g, '')) || literal.length > 24) return;
      const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
      const painted = lines.get(`${section}:${paragraph}:${controlIndex}:${index}:${p}`);
      if (!metric.stored || metric.controlCount || metric.lines?.length !== 1 || painted?.length !== 1 ||
          painted[0].page !== first.pageIndex || !(painted[0].h > 0) || painted[0].y < first.y ||
          painted[0].y + painted[0].h > first.y + first.h) return;
      text.push(literal); ink.push(painted[0]);
    }
    if (count === 2 && ink[0].y + ink[0].h > ink[1].y) return;
    const center = first.y + first.h / 2;
    if (merged) {
      // The native cut can leave a partially filled row inside a merge.
      // The complete short label still straddles that fragment's center.
      if (center < ink[0].y + ink[0].h / 2 || center > ink[1].y + ink[1].h / 2) return;
    } else if (Math.abs(ink[0].y + ink[count - 1].y + ink[count - 1].h - 2 * center) >
      2 * sourceCellBoundsTolerance) return;
    return `${Math.round(first.w * 100) / 100} ${hwpWrapFingerprint(JSON.stringify(text))}`;
  } catch { /* Optional placement data does not prevent conversion. */ }
}

type PhysicalCell = { cellIdx: number; pageIndex: number; y: number; h: number; w: number };

/** Record a compact first fragment only when the other ordinary cell has
 * a physical text prefix on that sheet and the remaining literal on the next. */
export function sourceCompactContinuedCell(doc: HwpDocument, section: number, paragraph: number, controlIndex: number,
  first: PhysicalCell, second: PhysicalCell): string | undefined {
  try {
    const close = sourceCellBoundsTolerance;
    if (second.pageIndex !== first.pageIndex + 1 || first.cellIdx !== second.cellIdx ||
        !(first.h > 0 && first.h <= 2000) || !(second.h > 0) ||
        !(first.w > 0 && first.w < 10000) || Math.abs(first.w - second.w) > close) return;
    const count = doc.getCellParagraphCount(section, paragraph, controlIndex, first.cellIdx);
    if (count < 2 || count > 24) return;
    const literal = Array.from({length:count}, (_, index) =>
      doc.getTextInCell(section, paragraph, controlIndex, first.cellIdx, index, 0, 0xffffffff)).join('');
    if (!literal || literal.length > 10000 || /[\x00-\x1f\ufffc]/u.test(literal)) return;
    const owned = ({pageIndex,y,h}:PhysicalCell) => {
      const { runs } = JSON.parse(doc.getPageTextLayout(pageIndex)) as { runs: (NativeCellInkRun & { headerFooter?: unknown })[] };
      const parts = runs.filter(run => !run.headerFooter && run.secIdx === section &&
        run.parentParaIdx === paragraph && run.cellPath?.length === 1 &&
        run.cellPath[0].controlIndex === controlIndex && run.cellPath[0].cellIndex === first.cellIdx && run.text?.trim());
      if (!parts.length || !parts.every(run => run.h > 0 && run.y >= y - close && run.y + run.h <= y + h + close)) return;
      return normalizeHwpWrapText(parts.map(run => run.text).join(''));
    };
    const prefix = owned(first), suffix = owned(second);
    if (!prefix || !suffix || prefix + suffix !== normalizeHwpWrapText(literal)) return;
    const px = (value: number) => Math.round(value * 100) / 100;
    return `${px(first.h)} ${px(first.w)} ${hwpWrapFingerprint(prefix)}`;
  } catch { /* Retain ordinary table flow when source evidence is incomplete. */ }
}
