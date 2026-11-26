import type { HwpDocument } from '@rhwp/core';
import { noteTerminal, type HwpTerminalNote } from '../hwp-note-terminal';
import { hwpWrapFingerprint, normalizeHwpWrapText, type HwpSourceWrap } from '../hwp-text-advances';

type Line = { height: number; spacing: number; baseline: number; textHeight: number; textStart: number; width: number; verticalPosition: number };
type Metrics = { stored: boolean; controlCount: number; lines?: Line[] };
type Paper = { marginTop: number; marginHeader: number; footerArea?: { y: number }; columns?: { x: number; width: number }[] };
type Footnote = { ok?: boolean; sectionIdx?: number; paraIdx?: number; controlIdx?: number; sourceType?: string };
type NativeRun = { secIdx: number; paraIdx: number; charStart: number; text: string; x: number; y: number; w: number; h: number };
const near = (a: number, b: number, allowance = .16) => Math.abs(a - b) <= allowance;

/** The two co-located note references and the second line are part of the
 * same source body paragraph. Preserve the paint origin only when its
 * original page also ends here and the following paragraph starts afresh. */
export function sourceFootnoteTerminal(doc: HwpDocument, section: number, paragraph: number,
  pitches: ReadonlyMap<string, number>, positions: readonly number[],
  notes: [HwpTerminalNote, HwpTerminalNote]): HwpSourceWrap | undefined {
  const key = `${section}:${paragraph}`, page = pitches.get('page:' + key);
  if (paragraph < 2 || page === undefined || positions.length !== 2 ||
      !positions.every(Number.isSafeInteger) || !(positions[0] > 0 && positions[0] < positions[1]) ||
      pitches.has('cross-page:' + key) || pitches.get(`page:${section}:${paragraph - 2}`) !== page ||
      pitches.get(`page:${section}:${paragraph + 1}`) !== page + 1) return;
  try {
    if (!doc.pageHasFootnoteFootholds(page)) return;
    const source = (index: number): Footnote | undefined => {
      try {
        const note = JSON.parse(doc.getPageFootnoteInfo(page, index)) as Footnote;
        return note.ok ? note : undefined;
      } catch { return; }
    };
    if (source(2) || ![source(0), source(1)].every((note, index) => note?.sourceType === 'body' &&
        note.sectionIdx === section && note.paraIdx === paragraph && note.controlIdx === index)) return;
    const metrics = (index: number) => JSON.parse(doc.getParagraphLineMetrics(section, index, '[]')) as Metrics;
    const text = (index: number) => doc.getTextRange(section, index, 0, 0xffffffff);
    const current = metrics(paragraph), blank = metrics(paragraph - 1), prior = metrics(paragraph - 2), next = metrics(paragraph + 1);
    if (![current, blank, prior, next].every(v => v.stored) || current.controlCount !== 2 || current.lines?.length !== 2 ||
        blank.controlCount || blank.lines?.length !== 1 || prior.controlCount || prior.lines?.length !== 1 ||
        next.controlCount || !next.lines?.length || text(paragraph - 1) !== '') return;
    const [first, second] = current.lines, [empty] = blank.lines, [previous] = prior.lines;
    const paper = JSON.parse(doc.getPageInfo(page)) as Paper, lane = paper.columns?.[0];
    if (!lane || paper.columns?.length !== 1 || paper.footerArea?.y === undefined) return;
    const bodyTop = paper.marginTop + paper.marginHeader, pitch = first.height + first.spacing;
    if (first.textStart || second.textStart <= positions[1] || next.lines[0].verticalPosition < 0 ||
        next.lines[0].verticalPosition > pitch / 3 || !near(first.width, lane.width) ||
        !near(first.width, second.width, .05) || !near(first.height, second.height, .05) ||
        !near(first.height, first.textHeight, .05) || !near(second.height, second.textHeight, .05) ||
        !near(first.baseline, second.baseline, .05) || !near(first.spacing, second.spacing, .05) ||
        !near(first.verticalPosition + pitch, second.verticalPosition, .08) ||
        !near(empty.verticalPosition + empty.height + empty.spacing, first.verticalPosition, .08) ||
        !near(previous.verticalPosition + previous.height + previous.spacing, empty.verticalPosition, .08)) return;
    const properties = JSON.parse(doc.getParaPropertiesAt(section, paragraph)), hanging = -Number(properties.indent);
    if (properties.alignment !== 'justify' || properties.spacingBefore || properties.spacingAfter ||
        properties.marginLeft || properties.marginRight || properties.keepLines || properties.keepWithNext ||
        properties.fontLineHeight || properties.lineSpacingType !== 'Percent' ||
        !near(pitch, first.height * Number(properties.lineSpacing) / 100, .08)) return;
    const literal = text(paragraph), earlier = text(paragraph - 2);
    if (!literal.trim() || literal.length > 160 || !earlier.trim() || earlier.length > 160 ||
        /[\u0000-\u001f\u0590-\u08ff\u202a-\u202e\u2066-\u2069\uD800-\uDFFF\ufffc]/u.test(literal + earlier)) return;
    const runs = (JSON.parse(doc.getPageTextLayout(page)) as { runs: NativeRun[] }).runs
      .filter(run => run.secIdx === section && run.paraIdx === paragraph);
    const sourceTail = runs.filter(run => near(run.y, bodyTop + second.verticalPosition) && run.text?.trim());
    if (sourceTail.length !== 1 || !runs.some(run => near(run.y, bodyTop + first.verticalPosition) && near(run.x, lane.x)) ||
        runs.some(run => !near(run.y, bodyTop + first.verticalPosition) && !near(run.y, bodyTop + second.verticalPosition))) return;
    const [tail] = sourceTail;
    if (!Number.isSafeInteger(tail.charStart) || tail.charStart <= positions[1] ||
        !near(tail.h, second.height) || !near(tail.x, lane.x + hanging) ||
        tail.x + tail.w > lane.x + lane.width + .16 || literal.slice(tail.charStart) !== tail.text) return;
    const normalized = normalizeHwpWrapText(literal), end = normalizeHwpWrapText(literal.slice(0, tail.charStart)).length;
    const terminal = noteTerminal({ width: first.width, fullBodyHeight: paper.footerArea.y - bodyTop,
      lineTop: first.verticalPosition, secondBaseline: second.verticalPosition + second.baseline, pitch, font: first.height,
      hanging, prior: hwpWrapFingerprint(normalizeHwpWrapText(earlier)), notes });
    if (terminal && end > 0 && end < normalized.length)
      return { fingerprint: hwpWrapFingerprint(normalized), ends: [end, normalized.length], noteTerminal: terminal };
  } catch { /* The ordinary page fit remains available without complete source evidence. */ }
}
