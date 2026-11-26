import type { HwpDocument } from '@rhwp/core';
import { hwpWrapFingerprint, normalizeHwpWrapText, type HwpSourceWrap } from '../hwp-text-advances';
import type { CellPath } from './cell-content';
import { cellParagraphKey, sourceCellBoundsTolerance } from './paragraph-layout';

type NestedCellBox = { cellIdx: number; pageIndex: number; y: number; h: number; w: number };
type StoredLine = { textStart: number; verticalPosition: number; height: number; width: number };

/** The nested cell is already drawn at its recorded width. Only an ordinary
 * fully hanging paragraph with unchanged painted line bands can supply a
 * source witness for its saved uniform glyph scale. */
export function sourceNestedHangingScale(
  doc: HwpDocument, section: number, paragraph: number, path: CellPath,
  physicalBoxes: readonly NestedCellBox[], pitches: Map<string, number>, hanging: number,
): HwpSourceWrap | undefined {
  const leaf = path.at(-1)!;
  const key = cellParagraphKey(section, paragraph, path);
  const bodies = physicalBoxes.filter(box => box.cellIdx === leaf.cellIndex);
  if (bodies.length !== 1) return;
  const [box] = bodies;
  const paintedTop = pitches.get(`top:${key}`), paintedPage = pitches.get(`page:${key}`);
  if (paintedTop === undefined || paintedPage !== box.pageIndex || !(box.w > hanging && box.h > 0)) return;

  try {
    const address = JSON.stringify(path);
    const literal = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
    // Stored start offsets use Unicode scalars. Restrict this proof to text
    // for which those coordinates also index the editable UTF-16 string.
    if (!literal.trim() || /[\uD800-\uDFFF\p{Bidi_Control}\x00-\x1f\ufffc]/u.test(literal)) return;
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address)) as
      { stored: boolean; controlCount: number; lines: StoredLine[] };
    if (!stored || controlCount || lines.length < 2 || lines.length > 80 ||
        pitches.get(`lines:${key}`) !== lines.length || lines[0].textStart !== 0) return;
    const [first] = lines, tolerance = sourceCellBoundsTolerance;
    if (!(first.width > hanging && first.width <= box.w + tolerance)) return;
    if (lines.some((line, index) => !(line.height > 0 &&
      Math.abs(line.width - first.width) <= tolerance && line.textStart < literal.length &&
      (!index || line.textStart > lines[index - 1].textStart && line.verticalPosition > lines[index - 1].verticalPosition) &&
      paintedTop + line.verticalPosition - first.verticalPosition >= box.y - tolerance &&
      paintedTop + line.verticalPosition - first.verticalPosition + line.height <= box.y + box.h + tolerance))) return;

    const runs = JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, address, 0, literal.length)) as
      { startOffset: number; endOffset: number }[];
    let offset = 0, scale = 0, size = 0;
    for (const run of runs) {
      if (run.startOffset !== offset || run.endOffset <= offset) return;
      const char = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, address, offset));
      const width = char.ratios?.[0], em = Number(char.fontSize) / 75;
      if (!Number.isInteger(width) || width < 50 || width >= 100 || !(em >= 4 && em <= 120) ||
          !Array.isArray(char.ratios) || char.ratios.length !== 7 || !char.ratios.every((value: number) => value === width) ||
          !Array.isArray(char.relativeSizes) || char.relativeSizes.some((value: number) => value !== 100) ||
          (scale && scale !== width) || (size && size !== em)) return;
      offset = run.endOffset; scale = width; size = em;
    }
    if (!scale || offset !== literal.length) return;
    const ends = [...lines.slice(1).map(line => normalizeHwpWrapText(literal.slice(0, line.textStart)).length),
      normalizeHwpWrapText(literal).length];
    if (ends.some((end, i) => end <= (ends[i - 1] ?? 0))) return;
    return { fingerprint: hwpWrapFingerprint(literal), ends, cellWidth: { outer: box.w, text: Math.min(box.w, first.width) },
      nestedHanging: { scale, hanging, size } };
  } catch { /* Missing optional source geometry retains the natural layout. */ }
}
