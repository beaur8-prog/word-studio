import type { HwpDocument } from '@rhwp/core';
import { type HwpNestedCellFace } from '../hwp-nested-table-layout';
import { hwpSingleBodyFace, hwpWrapFingerprint, normalizeHwpWrapText } from '../hwp-text-advances';

type Box = { cellIdx: number; row: number; col: number; rowSpan: number; colSpan: number; x: number; y: number; w: number; h: number };
type Ink = { x: number; y: number; w: number; h: number; text: string; displayText?: string };

const near = (a: number, b: number) => Math.abs(a - b) <= .16;

/** Record a natural source line in a tight, explicitly condensed cell. The
 * editable font and scale remain authored. Print can verify a local substitute
 * only when the original face is absent. */
export function sourceNestedCellFaces(doc: HwpDocument, section: number, paragraph: number, block: number,
  address: (index: number) => string, boxes: readonly Box[], painted: readonly Ink[]): HwpNestedCellFace[] {
  const faces: HwpNestedCellFace[] = [];
  for (const box of boxes) {
    if (box.colSpan !== 1 || box.rowSpan !== 1 || box.h > 80 || box.w < 30 || box.w > 200) continue;
    const path = address(box.cellIdx);
    if (doc.getCellParagraphCountByPath(section, paragraph, path) !== 1) continue;
    const text = doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
    if (!/^[가-힣 ]{8,40}$/u.test(text)) continue;
    const own = JSON.parse(doc.getCellPropertiesByPath(section, paragraph, path, box.cellIdx));
    const style = JSON.parse(doc.getCellParaPropertiesAtByPath(section, paragraph, path));
    if (own.isHeader || own.verticalAlign !== 1 || own.textDirection || style.alignment !== 'center' ||
        style.indent || style.marginLeft || style.marginRight || style.spacingBefore || style.spacingAfter ||
        style.lineSpacingType !== 'Percent' || style.fontLineHeight || style.lineSpacing < 85 || style.lineSpacing > 100) continue;
    const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
    const lines = metric.lines as { textStart: number; width: number; height: number; spacing: number }[] | undefined;
    if (!metric.stored || metric.controlCount || !lines?.length || lines.length > 2 || lines[0].textStart !== 0) continue;
    const { height: size, width, spacing } = lines[0], pitch = size + spacing;
    if (!(width > 0 && width <= box.w && size > 0 && near(pitch, size * style.lineSpacing / 100)) ||
        lines.some(line => !near(line.width, width) || !near(line.height, size) || !near(line.spacing, spacing))) continue;
    const runs = JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, path, 0, text.length));
    if (runs.length !== 1) continue;
    const chars = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, path, 0));
    const scale = Number(chars.ratios?.[0]);
    if (chars.bold || chars.italic || !near(chars.fontSize / 75, size) || !Number.isInteger(scale) || scale < 80 || scale >= 100 ||
        chars.ratios?.length !== 7 || !chars.ratios.every((value: number) => value === scale) ||
        chars.relativeSizes?.length !== 7 || !chars.relativeSizes.every((value: number) => value === 100) ||
        chars.fontFamilies?.length !== 7 || !chars.fontFamilies.every((family: string) => family === hwpSingleBodyFace)) continue;
    const bands = new Map<number, Ink[]>();
    for (const run of painted) {
      if (!normalizeHwpWrapText(run.displayText ?? run.text) || run.x < box.x - .16 || run.y < box.y - .16 ||
          run.x + run.w > box.x + box.w + .16 || run.y + run.h > box.y + box.h + .16) continue;
      const band = bands.get(run.y) ?? [];
      band.push(run);
      bands.set(run.y, band);
    }
    if (bands.size !== lines.length) continue;
    const whole = normalizeHwpWrapText(text);
    const expected = [...lines.slice(1).map(line => normalizeHwpWrapText(text.slice(0, line.textStart)).length), whole.length];
    let count = 0, maximumInk = 0;
    const parts = [...bands].sort(([a], [b]) => a - b).map(([, band]) => {
      const ordered = band.sort((a, b) => a.x - b.x);
      maximumInk = Math.max(maximumInk, ordered.at(-1)!.x + ordered.at(-1)!.w - ordered[0].x);
      return normalizeHwpWrapText(ordered.map(run => run.displayText ?? run.text).join(''));
    });
    if (parts.join('') !== whole || parts.some((part, i) => (count += part.length) !== expected[i]) ||
        maximumInk < width * .9 || maximumInk > width + .16) continue;
    faces.push({ block, row: box.row, column: box.col, width: box.w, height: box.h,
      size: +size.toFixed(4), pitch: +pitch.toFixed(4), scale,
      fingerprint: hwpWrapFingerprint(text), ends: expected });
  }
  return faces;
}
