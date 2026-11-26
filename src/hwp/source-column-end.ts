import type { HwpDocument } from '@rhwp/core';
import type { HwpColumnEndFit, HwpTerminalLineFit } from '../hwp-body-line';
import { hwpWrapFingerprint, normalizeHwpWrapText, type HwpSourceWrap } from '../hwp-text-advances';

type StoredLine = { textStart: number; verticalPosition: number; height: number; textHeight: number;
  spacing: number; baseline: number; width: number };
type Lines = { stored: boolean; controlCount: number; lines?: StoredLine[] };
type TextPosition = { page: number; x: number; y: number };
type Paper = { marginTop: number; marginHeader: number; footerArea?: { y: number }; columns?: { x: number; width: number }[] };
type SourceRun = { type: string; text?: string; rotation?: number; isVertical?: boolean; projectionKind?: string;
  source?: { stableSourceKey?: string }; bbox?: { x: number; y: number };
  placement?: { baselineY?: number; runToPage?: { a: number; b: number; c: number; d: number; e: number; f: number } } };
type Layer = { ops?: SourceRun[]; children?: Layer[]; child?: Layer };

const near = (a: number, b: number, tolerance = .16) => Math.abs(a - b) <= tolerance;
const unmeasurable = /[\u0000-\u001f\u0590-\u08ff\u202a-\u202e\u2066-\u2069\uD800-\uDFFF\ufffc]/u;
const bodyTop = (paper: Paper) => paper.marginTop + paper.marginHeader;
const laneAt = (paper: Paper, x: number) => paper.columns?.length === 2
  ? paper.columns.findIndex(column => x >= column.x - .2 && x < column.x + column.width) : -1;

function exactSourceBaseline(layer: Layer | undefined, section: number, paragraph: number,
  text: string, firstX: number, y: number, maxRuns = 5): number | null {
  const prefix = `section:${section}/para:${paragraph}/char:`;
  const found: { offset: number; text: string; x: number; baseline: number }[] = [];
  const visit = (node: Layer | undefined, depth: number) => {
    if (!node || depth > 60) return;
    for (const run of node.ops ?? []) {
      const key = run.source?.stableSourceKey;
      if (run.type !== 'textRun' || !key?.startsWith(prefix) || !run.text ||
          run.rotation !== 0 || run.isVertical || run.projectionKind !== 'verbatim') continue;
      const rawOffset = key.slice(prefix.length);
      const box = run.bbox, placement = run.placement, transform = placement?.runToPage;
      if (!/^\d+$/.test(rawOffset) || !box || !near(box.y, y) || placement?.baselineY !== 0 || !transform ||
          transform.a !== 1 || transform.d !== 1 || transform.b !== 0 || transform.c !== 0 ||
          !near(transform.e, box.x) || !Number.isFinite(transform.f)) continue;
      found.push({ offset: Number(rawOffset), text: run.text, x: transform.e, baseline: transform.f - box.y });
    }
    for (const child of node.children ?? []) visit(child, depth + 1);
    visit(node.child, depth + 1);
  };
  visit(layer, 0);
  if (!found.length || found.length > maxRuns) return null;
  found.sort((a, b) => a.offset - b.offset);
  if (!near(found[0].x, firstX)) return null;
  const baseline = found[0].baseline;
  let offset = 0;
  for (const run of found) {
    if (run.offset !== offset || text.slice(offset, offset + run.text.length) !== run.text || !near(run.baseline, baseline, .01)) return null;
    offset += run.text.length;
  }
  return offset === text.length ? baseline : null;
}

function measuredPosition(pitches: ReadonlyMap<string, number>, section: number, index: number,
  leading = false): TextPosition | undefined {
  const key = `${section}:${index}`;
  const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`);
  const x = pitches.get(`x:${key}`) ?? (leading ? pitches.get(`leading-x:${key}`) : undefined);
  if (page === undefined || x === undefined || y === undefined || pitches.has(`cross-page:${key}`)) return;
  return { page, x, y };
}

/** A mixed-font line can keep its stored advance when the next line
 * starts immediately after that advance and its authored paragraph spacing. */
export function sourceHwpMixedColumnPitch(doc: HwpDocument, section: number, paragraph: number,
  pitches: ReadonlyMap<string, number>, font: number, pitch: number, after: number): number | undefined {
  const here = measuredPosition(pitches, section, paragraph, true);
  if (!here || pitches.get(`lines:${section}:${paragraph}`) !== 1 || !(after >= 0)) return;
  try {
    const paper: Paper = JSON.parse(doc.getPageInfo(here.page));
    const current: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const next: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph + 1, '[]'));
    const position = JSON.parse(doc.getPageOfPosition(section, paragraph + 1));
    if (laneAt(paper, here.x) < 0 || !position.ok || position.page !== here.page ||
        !current.stored || current.controlCount || current.lines?.length !== 1 || !next.stored || !next.lines?.length) return;
    const [line] = current.lines, following = next.lines[0];
    if (line.textStart || !near(line.height, font, .06) || !near(line.textHeight, font, .06) ||
        !near(line.height + line.spacing, pitch, .06) || line.width < font * 4 ||
        !near(here.y, bodyTop(paper) + line.verticalPosition) ||
        !near(line.verticalPosition + line.height + line.spacing + after, following.verticalPosition, .12)) return;
    const bottom = paper.footerArea?.y;
    if (bottom !== undefined && here.y + pitch + after <= bottom + .08) return line.width;
  } catch { /* The natural mixed-font line remains available without a stored advance. */ }
}

/** Only a complete stored line before a native column/page break can omit
 * its unused leading. The source layer supplies the shared paint baseline. */
export function sourceHwpColumnEndFit(doc: HwpDocument, section: number, paragraph: number,
  pitches: ReadonlyMap<string, number>, layer: (page: number) => Layer | undefined): HwpColumnEndFit | undefined {
  if (!paragraph) return;
  const here = measuredPosition(pitches, section, paragraph), before = measuredPosition(pitches, section, paragraph - 1),
    next = measuredPosition(pitches, section, paragraph + 1);
  if (!here || !before || !next || before.page !== here.page || next.page < here.page || next.page > here.page + 1 ||
      pitches.get(`lines:${section}:${paragraph}`) !== 1 || pitches.get(`lines:${section}:${paragraph - 1}`) !== 1) return;
  try {
    const source: Paper = JSON.parse(doc.getPageInfo(here.page));
    const following: Paper = next.page === here.page ? source : JSON.parse(doc.getPageInfo(next.page));
    const lane = laneAt(source, here.x), followingLane = laneAt(following, next.x);
    if (lane < 0 || laneAt(source, before.x) !== lane ||
        (lane === 0 ? next.page !== here.page || followingLane !== 1 : next.page !== here.page + 1 || followingLane !== 0) ||
        !near(next.y, bodyTop(following))) return;
    const current: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const previous: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph - 1, '[]'));
    if (!current.stored || current.controlCount || current.lines?.length !== 1 ||
        !previous.stored || previous.controlCount || previous.lines?.length !== 1) return;
    const [line] = current.lines, [prior] = previous.lines;
    const properties = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
    const pitch = line.height + line.spacing, bottom = source.footerArea?.y;
    const y = bodyTop(source) + line.verticalPosition;
    if (bottom === undefined || properties.spacingBefore || properties.spacingAfter ||
        !near(y, here.y) || !near(bodyTop(source) + prior.verticalPosition, before.y) ||
        !near(prior.verticalPosition + prior.height + prior.spacing, line.verticalPosition, .08) ||
        !near(prior.baseline, line.baseline, .08) || line.textStart || line.width < line.height * 4 ||
        !near(line.height, line.textHeight, .08) || !near(line.height, prior.height, .08) ||
        !near(pitch, line.height * properties.lineSpacing / 100, .08) ||
        bottom - y < line.height - .16 || y + pitch <= bottom + .5) return;
    const text = doc.getTextRange(section, paragraph, 0, 0xffffffff);
    if (!text.trim() || text.length > 200 || unmeasurable.test(text)) return;
    const baseline = exactSourceBaseline(layer(here.page), section, paragraph, text, here.x, here.y);
    if (baseline === null || !(baseline > 0 && baseline <= line.height)) return;
    const preceding = doc.getTextRange(section, paragraph - 1, 0, 0xffffffff);
    if (!preceding.trim() || preceding.length > 200 || unmeasurable.test(preceding) ||
        !near(line.width, prior.width, .08) || !near(here.x, before.x)) return;
    const hang = /[^\s] $/u.test(preceding);
    return { baseline, remainder: bottom - y, width: line.width, fingerprint: hwpWrapFingerprint(text),
      preceding: hwpWrapFingerprint(preceding), ...(hang ? { hang: true } : {}) };
  } catch { /* Missing source placement leaves the ordinary measured column end in use. */ }
}


/** Called after the single-column source flow proves a natural page end.
 * Bind the drawing baseline to the whole line and its actual predecessor. */
export function sourceHwpBodyEndFit(doc: HwpDocument, section: number, paragraph: number,
  pitches: ReadonlyMap<string, number>, layer: (page: number) => Layer | undefined): HwpTerminalLineFit | undefined {
  const here = measuredPosition(pitches, section, paragraph, true), before = measuredPosition(pitches, section, paragraph - 1, true);
  if (!here || !before || before.page !== here.page || !near(before.x, here.x) ||
      pitches.get(`lines:${section}:${paragraph - 1}`) !== 1) return;
  try {
    const paper: Paper = JSON.parse(doc.getPageInfo(here.page));
    const current: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const previous: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph - 1, '[]'));
    const line = current.lines?.[0], prior = previous.lines?.[0];
    if (!line || !prior || previous.lines?.length !== 1 || !near(prior.width, line.width, .08) ||
        !near(bodyTop(paper) + prior.verticalPosition, before.y) || paper.footerArea?.y === undefined) return;
    const text = doc.getTextRange(section, paragraph, 0, 0xffffffff);
    const earlier = doc.getTextRange(section, paragraph - 1, 0, 0xffffffff);
    if (!text.trim() || !earlier.trim() || text.length > 200 || earlier.length > 200 ||
        unmeasurable.test(text) || unmeasurable.test(earlier)) return;
    const baseline = exactSourceBaseline(layer(here.page), section, paragraph, text, here.x, here.y, 12);
    if (baseline === null || !(baseline > 0 && baseline <= line.height)) return;
    return { baseline, remainder: paper.footerArea.y - here.y, width: line.width,
      fingerprint: hwpWrapFingerprint(text), preceding: hwpWrapFingerprint(earlier) };
  } catch { /* The source line keeps its ordinary page fit on missing evidence. */ }
}

/** The authored page break follows a complete line whose paint fits the
 * source body when its unused leading extends outside. */
export function sourceHwpForcedBodyEnd(doc: HwpDocument, section: number, paragraph: number,
  pitches: ReadonlyMap<string, number>, layer: (page: number) => Layer | undefined): HwpSourceWrap | undefined {
  if (!paragraph) return;
  const here = measuredPosition(pitches, section, paragraph, true);
  const preceding = measuredPosition(pitches, section, paragraph - 1, true);
  const following = measuredPosition(pitches, section, paragraph + 1, true);
  if (!here || !preceding || !following || preceding.page !== here.page || following.page !== here.page + 1 ||
      pitches.get(`lines:${section}:${paragraph}`) !== 1) return;
  try {
    const paper: Paper = JSON.parse(doc.getPageInfo(here.page));
    const next: Paper = JSON.parse(doc.getPageInfo(following.page));
    if (paper.columns?.length !== 1 || next.columns?.length !== 1 || !near(following.y, bodyTop(next)) ||
        paper.footerArea?.y === undefined) return;
    const current: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const before: Lines = JSON.parse(doc.getParagraphLineMetrics(section, paragraph - 1, '[]'));
    if (!current.stored || current.controlCount || current.lines?.length !== 1 || !before.stored || before.controlCount ||
        !before.lines || before.lines.length < 2 || before.lines.length > 8) return;
    const [line] = current.lines, prior = before.lines.at(-1)!;
    const style = JSON.parse(doc.getParaPropertiesAt(section, paragraph));
    const prevStyle = JSON.parse(doc.getParaPropertiesAt(section, paragraph - 1));
    const spaceBefore = style.spacingBefore / 2, hanging = -style.indent;
    const room = paper.footerArea.y - here.y, pitch = line.height + line.spacing;
    if (line.textStart || style.alignment !== 'justify' || style.marginLeft || style.marginRight || style.spacingAfter ||
        style.keepLines || style.keepWithNext || prevStyle.spacingAfter || style.fontLineHeight || style.lineSpacingType !== 'Percent' ||
        !(spaceBefore > 0 && spaceBefore <= pitch && hanging > 0 && hanging < line.width / 4) ||
        !(room >= line.height - .08 && room < pitch - .5 && pitch <= line.height * 2) ||
        !near(line.width, prior.width, .08) || !near(line.height, line.textHeight, .08) ||
        !near(pitch, line.height * style.lineSpacing / 100, .08) ||
        !near(bodyTop(paper) + before.lines[0].verticalPosition, preceding.y) ||
        !near(bodyTop(paper) + line.verticalPosition, here.y) ||
        !near(line.verticalPosition - prior.verticalPosition - prior.height - prior.spacing, spaceBefore, .08)) return;
    const text = doc.getTextRange(section, paragraph, 0, 0xffffffff);
    const earlier = doc.getTextRange(section, paragraph - 1, 0, 0xffffffff);
    if (!text.trim() || !earlier.trim() || text.length > 160 || earlier.length > 250 ||
        unmeasurable.test(text) || unmeasurable.test(earlier)) return;
    const baseline = exactSourceBaseline(layer(here.page), section, paragraph, text, here.x, here.y, 12);
    if (baseline === null || !(baseline > 0 && baseline <= line.height) || !near(baseline, line.baseline, .08)) return;
    return { fingerprint: hwpWrapFingerprint(text), ends: [normalizeHwpWrapText(text).length],
      forcedBodyEnd: { baseline, room, width: line.width, spaceBefore, pitch, font: line.height, hanging,
        preceding: hwpWrapFingerprint(earlier), precedingLines: before.lines.length } };
  } catch { /* Without a complete source witness the ordinary line box remains. */ }
}
