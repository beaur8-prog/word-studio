import type { HwpBodyPictureFlow } from '../image-floating';
import { hwpShortCellLabel, hwpWrapFingerprint, normalizeHwpWrapText, type HwpSourceWrap } from '../hwp-text-advances';
import type { HwpDocument } from '@rhwp/core';
import { encodeHwpBodyLine, type HwpBodyLine, type HwpStoredMixedLinePair } from '../hwp-body-line';
import type { CellPath } from './cell-content';
import { encodeHwpCellLine, type HwpAuthoredCellTerminal } from '../hwp-cell-line';
import { paragraphBorderBoxAttribute, parseParagraphBorderBox } from '../paragraph-border-box';
import { hwpBorderWidthsMillimeters } from './cell-borders';

const escape = (text: string) => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);

export const cellParagraphKey = (section: number, paragraph: number, path: CellPath) =>
  `${section}:${paragraph}:` + path.map(c => `${c.controlIndex}:${c.cellIndex}:${c.cellParaIndex}`).join('/');

/** Recover percentage leading when the original cell actually gave its
 * smaller line a different advance. It remains responsive after editing. */
export function sourceMixedCellLeading(doc: HwpDocument, section: number, paragraph: number, path: CellPath,
  properties: Record<string, any>, base: number, html: string): string {
  const percent = Number(properties.lineSpacing);
  if (properties.lineSpacingType !== 'Percent' || properties.fontLineHeight || !(percent > 100) || !Number.isFinite(base) ||
    /\t|<(?:br|img|svg|table|sup|sub|u|s|del|mark)\b|<!--\s*rhwp:/i.test(html)) return '';
  const shapes = [...html.matchAll(/font-size:\s*([\d.]+)(pt|px)/gi)].map(match =>
    Number(match[1]) * (match[2].toLowerCase() === 'pt' ? 4 / 3 : 1));
  if (new Set(shapes).size < 2) return '';
  type Metrics = {
    stored: boolean; controlCount: number;
    lines?: { height: number; textHeight: number; spacing: number; verticalPosition: number }[];
  };
  let metrics: Metrics;
  try { metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, JSON.stringify(path))); }
  catch { return ''; }
  const { stored, controlCount, lines } = metrics;
  if (!stored || controlCount || !lines || lines.length < 2) return '';
  const font = shapes.reduce((largest, size) => Math.max(largest, size), 0);
  if (!lines.some(line => line.height < font - .1) || Math.abs(base - font * percent / 100) > .08 ||
    !lines.every((line, index) => line.height > 0 &&
      Math.abs(line.height - line.textHeight) < .08 &&
      shapes.some(size => Math.abs(size - line.height) < .08) &&
      Math.abs(line.height + line.spacing - line.height * percent / 100) < .08 &&
      (!index || Math.abs(line.verticalPosition - lines[index - 1].verticalPosition -
        lines[index - 1].height - lines[index - 1].spacing) < .08))) return '';
  return `--hwp-body-line:${encodeHwpBodyLine({ percent, font, base, cell: true })};`;
}

/** Preserve a witnessed import pitch that the standard DOCX spacing unit
 * cannot express. Stored lines must corroborate the observed uniform flow. */
export function sourcePreciseColumnPitch(doc: HwpDocument, section: number, paragraph: number,
  pitches: ReadonlyMap<string, number>, font: number, base: number, percent: number): boolean {
  const key = `${section}:${paragraph}`;
  const count = pitches.get(`lines:${key}`);
  if (count === undefined || count < 2 || pitches.has(`cross-page:${key}`) ||
      Math.abs(base - font * percent / 100) > .05 ||
      Math.abs(base - Math.round(base * 15) / 15) < 1 / 64) return false;
  try {
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]')) as
      { stored: boolean; controlCount: number; lines: StoredCellLine[] };
    return stored && !controlCount && lines?.length === count && lines[0].textStart === 0 &&
      lines.every((line, index) => {
        const previous = lines[index - 1];
        const pitch = line.height + line.spacing;
        return Math.abs(line.height - font) <= .04 && Math.abs(line.textHeight - font) <= .04 &&
          Math.abs(pitch - base) <= .04 && (!previous ||
            line.textStart > previous.textStart && Math.abs(line.verticalPosition - previous.verticalPosition -
              previous.height - previous.spacing) <= .04);
      });
  } catch { return false; }
}

/** Keep native paragraph policy separate from rendered font spacing. Word's
 * omitted autoSpaceDE/DN defaults enable extra East Asian/Latin/number gaps. */
export function paragraphAutoSpace(tag: string, properties: Record<string, any>): string {
  const attributes = [['autoSpaceKrEn', 'de'], ['autoSpaceKrNum', 'dn']].flatMap(([key, suffix]) =>
    typeof properties[key] === 'boolean' ? [`data-word-auto-space-${suffix}="${properties[key]}"`] : []);
  return attributes.length ? tag.replace(/>$/, ` ${attributes.join(' ')}>`): tag;
}

export function physicalLineHeight(properties: Record<string, any>, char?: Record<string, any>): number {
  const fontPx = Number(char?.fontSize) / 75;
  const line = Number(properties.lineSpacing);
  return properties.lineSpacingType === 'Percent' ? fontPx * line / 100
    : properties.lineSpacingType === 'Fixed' ? line
    : properties.lineSpacingType === 'SpaceOnly' ? fontPx + line
    : properties.lineSpacingType === 'Minimum' ? Math.max(fontPx, line) : NaN;
}

/** Recover the initial section controls' line before a separate inline
 * table line only when the actual native table paints below both reserves. */
export function sourceInitialTableLead(doc: HwpDocument, table: { index: number; textStart: number; y: number; height: number },
  font: number, percent: number, bodyTop: number) {
  try {
    const { stored, lines, tables } = JSON.parse(doc.getParagraphLineMetrics(0, 0, '[]'));
    if (!stored || lines?.length !== 2 || tables?.length !== 1 || !tables[0].inline || tables[0].index !== table.index) return;
    const [controlLine, tableLine] = lines;
    const { outerTop, outerBottom } = tables[0];
    const advance = tableLine.verticalPosition - controlLine.verticalPosition;
    if (![advance, font, percent, outerTop, outerBottom, table.y, table.height, bodyTop,
      controlLine.baseline, controlLine.textHeight, controlLine.spacing, tableLine.textHeight].every(Number.isFinite) ||
      !(advance > font && font > 0 && percent > 100 && outerTop > 0 && outerBottom >= 0 && table.height > 0) ||
      table.textStart <= 0 || tableLine.textStart !== table.textStart || controlLine.textStart !== 0 || controlLine.verticalPosition !== 0 ||
      !(controlLine.baseline > 0 && controlLine.baseline <= font) ||
      Math.abs(controlLine.textHeight - font) > .08 || Math.abs(advance - font - controlLine.spacing) > .08 ||
      Math.abs(advance - font * percent / 100) > .08 ||
      Math.abs(tableLine.textHeight - table.height - outerTop - outerBottom) > .1 ||
      Math.abs(table.y - bodyTop - advance - outerTop) > .1) return;
    return { advance, outerTop };
  } catch { /* Some native streams do not retain compatible stored lines. */ }
}

/** A wide floating table can carry its visible host text above its first
 * row even when the logical control starts that paragraph. Only project the
 * text first when its native baseline and the first row share that seam. */
export function sourceHostBeforeFloatingTable(doc: HwpDocument, section: number, paragraph: number, control: number,
  pitches: ReadonlyMap<string, number>): boolean {
  const key = `${section}:${paragraph}`;
  const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`);
  if (page === undefined || y === undefined || pitches.get(`lines:${key}`) !== 1) return false;
  const props = JSON.parse(doc.getTableProperties(section, paragraph, control));
  if (props.treatAsChar || props.hasCaption || props.vertRelTo !== 'Para' || props.vertAlign !== 'Top' ||
    props.textWrap !== 'TopAndBottom' || !(props.vertOffset > 0)) return false;
  const { stored, controlCount, lines, tables } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  if (!stored || controlCount !== 1 || lines?.length !== 1 || tables?.length !== 1 || tables[0].inline) return false;
  const { baseline, height, verticalPosition } = lines[0];
  const boxes = JSON.parse(doc.getTableCellBboxes(section, paragraph, control)) as { pageIndex: number; y: number }[];
  const first = boxes[0];
  if (!first || first.pageIndex !== page || !boxes.some(box => box.pageIndex > page)) return false;
  const source = JSON.parse(doc.getPageInfo(page)), gap = first.y - y;
  return source.columns?.length === 1 && baseline > 0 && height >= baseline && gap >= baseline - .1 && gap <= height + .1 &&
    Math.abs(y - source.marginTop - source.marginHeader - verticalPosition) <= .15;
}

/** A wide paragraph-relative table can start at the next body while its
 * entire visible host stays on the old sheet. Verify both physical origins. */
export function sourceHostOnPriorPage(doc: HwpDocument, section: number, paragraph: number, control: number,
  pitches: ReadonlyMap<string, number>) {
  const key = `${section}:${paragraph}`;
  const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`);
  if (page === undefined || y === undefined || pitches.has(`cross-page:${key}`)) return;
  const props = JSON.parse(doc.getTableProperties(section, paragraph, control));
  if (props.treatAsChar || props.hasCaption || props.vertRelTo !== 'Para' || props.vertAlign !== 'Top' ||
      props.horzRelTo !== 'Para' || props.textWrap !== 'TopAndBottom' || !(props.vertOffset > 0)) return;
  const { stored, controlCount, lines, tables } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  if (!stored || controlCount !== 1 || !lines?.length || lines.length < 2 || lines.length !== pitches.get(`lines:${key}`) ||
      tables?.length !== 1 || tables[0].inline) return;
  const boxes = JSON.parse(doc.getTableCellBboxes(section, paragraph, control)) as
    { pageIndex: number; x: number; y: number; w: number; h: number }[];
  const first = boxes[0], outerTop = Number(props.outerTop) / 75, outerBottom = Number(props.outerBottom) / 75;
  if (!first || !(outerTop >= 0 && outerBottom >= 0)) return;
  let left = first.x, right = first.x + first.w, end = first.y + first.h;
  for (const box of boxes) {
    if (box.pageIndex !== page + 1 || !(box.w > 0 && box.h > 0)) return;
    left = Math.min(left, box.x);
    right = Math.max(right, box.x + box.w);
    end = Math.max(end, box.y + box.h);
  }
  const source = JSON.parse(doc.getPageInfo(page)), following = JSON.parse(doc.getPageInfo(page + 1));
  const sourceTop = Number(source.marginTop) + Number(source.marginHeader);
  const sourceBottom = Number(source.footerArea?.y ?? (source.height - source.marginBottom - source.marginFooter));
  const followTop = Number(following.marginTop) + Number(following.marginHeader);
  const lastLine = lines.at(-1);
  if (source.columns?.length !== 1 || following.columns?.length !== 1 ||
      !(Math.abs(y - sourceTop - lines[0].verticalPosition) <= .15 && Math.abs(first.y - followTop - outerTop) <= .15) ||
      !(lastLine.height > 0 && sourceTop + lastLine.verticalPosition + lastLine.height <= sourceBottom + .15)) return;
  const x = left - Number(following.columns[0].x);
  if (!(x >= 0 && right <= following.columns[0].x + following.columns[0].width + .15)) return;
  let blankUsesTableGap = false;
  const after = `${section}:${paragraph + 2}`;
  if (paragraph + 2 < doc.getParagraphCount(section) && pitches.get(`page:${after}`) === page + 1 &&
    doc.getTextRange(section, paragraph + 1, 0, 0xffffffff) === '') {
    const next = JSON.parse(doc.getParaPropertiesAt(section, paragraph + 1));
    const blank = JSON.parse(doc.getParagraphLineMetrics(section, paragraph + 1, '[]'));
    const text = JSON.parse(doc.getParagraphLineMetrics(section, paragraph + 2, '[]'));
    const built = blank.lines?.[0];
    if (next.spacingBefore > 0 && !next.spacingAfter && !next.pageBreakBefore && blank.stored && blank.controlCount === 0 &&
        blank.lines?.length === 1 && text.stored && text.controlCount === 0 && text.lines?.length &&
        Math.abs(followTop + built.verticalPosition - end - outerBottom) <= .15 &&
        Math.abs(text.lines[0].verticalPosition - built.verticalPosition - built.height - built.spacing) <= .1)
      blankUsesTableGap = true;
  }
  return { x, top: outerTop, bottom: outerBottom, blankUsesTableGap };
}

type InlineTableBox = { pageIndex: number; cellIdx: number; x: number; y: number; w: number; h: number };

function sourceInlineTable(doc: HwpDocument, section: number, paragraph: number, control: number, pageBreak: number,
  horizontalReference: 'Column' | 'Para' = 'Column') {
  if (paragraph < 0 || doc.getTextRange(section, paragraph, 0, 0xffffffff).trim()) return;
  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  if (!metrics?.stored || metrics.controlCount !== 1 || metrics.lines?.length !== 1) return;
  const table = JSON.parse(doc.getTableProperties(section, paragraph, control));
  if (!table?.treatAsChar || table.hasCaption || table.pageBreak !== pageBreak ||
      table.textWrap !== 'TopAndBottom' || table.horzRelTo !== horizontalReference ||
      table.vertRelTo !== 'Para' || table.vertAlign !== 'Top') return;
  const dims = JSON.parse(doc.getTableDimensions(section, paragraph, control));
  const boxes: InlineTableBox[] = JSON.parse(doc.getTableCellBboxes(section, paragraph, control));
  const first = boxes[0];
  if (!first || boxes.length !== dims.cellCount || new Set(boxes.map(box => box.cellIdx)).size !== boxes.length ||
      boxes.some(box => box.pageIndex !== first.pageIndex || !(box.w > 0 && box.h > 0))) return;
  const page = JSON.parse(doc.getPageInfo(first.pageIndex));
  const top = Number(page.marginTop) + Number(page.marginHeader);
  const bottom = Number(page.footerArea?.y ?? (page.height - page.marginBottom - page.marginFooter));
  const start = Math.min(...boxes.map(box => box.y)), end = Math.max(...boxes.map(box => box.y + box.h));
  if (page.columns?.length !== 1 || !(start >= top - .2 && end <= bottom + .2)) return;
  const [column] = page.columns;
  const fillsColumn = Math.abs(Math.min(...boxes.map(box => box.x)) - column.x) <= .15 &&
    Math.abs(Math.max(...boxes.map(box => box.x + box.w)) - column.x - column.width) <= .15;
  return { page: first.pageIndex, top, bottom, start, end, rows: dims.rowCount, line: metrics.lines[0], fillsColumn,
    outerTop: Number(table.outerTop) / 75, outerBottom: Number(table.outerBottom) / 75 };
}

/** The stored inline line reserves the logical cell and outside spacing.
 * Both neighboring blank lines attest to that advance independently. */
export function sourceInlineStrokeHeight(doc: HwpDocument, section: number, paragraph: number, control: number) {
  try {
    const source = sourceInlineTable(doc, section, paragraph, control, 2, 'Para');
    if (!source || doc.getParagraphLength(section, paragraph - 1) || doc.getParagraphLength(section, paragraph + 1)) return;
    const before = JSON.parse(doc.getParagraphLineMetrics(section, paragraph - 1, '[]'));
    const after = JSON.parse(doc.getParagraphLineMetrics(section, paragraph + 1, '[]'));
    if (!before.stored || !after.stored || before.controlCount || after.controlCount ||
        before.lines?.length !== 1 || after.lines?.length !== 1) return;
    const { start, end, top, line, outerTop, outerBottom } = source;
    const previous = before.lines[0], following = after.lines[0], height = end - start;
    if (Math.abs(start - top - line.verticalPosition - outerTop) > .12 ||
        Math.abs(line.height - height - outerTop - outerBottom) > .12 ||
        Math.abs(line.verticalPosition - previous.verticalPosition - previous.height - previous.spacing) > .12 ||
        Math.abs(following.verticalPosition - line.verticalPosition - line.height - line.spacing) > .12) return;
    return Math.round(height * 100) / 100;
  } catch { return; }
}

/** Keep a stored inline control intact when the native layout placed it
 * on the next page and its preceding object and optional blank fill that
 * earlier page. The blank remains an authored paragraph. */
export function sourceKeptInlineTable(doc: HwpDocument, section: number, paragraph: number, control: number,
  previousDistance: 1 | 2 = 1, pitches?: Map<string, number>) {
  try {
    const current = sourceInlineTable(doc, section, paragraph, control, 1);
    if (!current || Math.abs(current.start - current.top - current.outerTop) > .2 ||
        Math.abs(current.line.verticalPosition) > .1) return;
    const previous = sourceInlineTable(doc, section, paragraph - previousDistance, 0, 1);
    if (!previous || current.page !== previous.page + 1) return;
    let used = previous.end + previous.outerBottom;
    if (previousDistance === 2) {
      // Keep the boundary after a mixed body page. Successive page-leading
      // tables can otherwise cascade an earlier reflow across later breaks.
      if (previous.start <= previous.top + previous.outerTop + .2) return;
      const index = paragraph - 1, key = `${section}:${index}`;
      if (doc.getParagraphLength(section, index) || pitches?.get(`page:${key}`) !== previous.page) return;
      const origin = pitches?.get(`top:${key}`);
      const properties = JSON.parse(doc.getParaPropertiesAt(section, index));
      const metrics = JSON.parse(doc.getParagraphLineMetrics(section, index, '[]'));
      if (!metrics.stored || metrics.controlCount || metrics.lines?.length !== 1 || origin === undefined ||
          properties.pageBreakBefore || properties.lineSpacingType !== 'Percent' || properties.indent ||
          properties.marginLeft || properties.marginRight || properties.spacingBefore || properties.spacingAfter) return;
      const blank = metrics.lines[0], line = previous.line;
      if (!Number.isFinite(blank.height) || blank.height <= 0 || !Number.isFinite(blank.spacing) || blank.spacing < 0 ||
          Math.abs(blank.verticalPosition - line.verticalPosition - line.height - line.spacing) > .12 ||
          Math.abs(origin - previous.top - blank.verticalPosition) > .12 || origin < used ||
          origin + blank.height + blank.spacing > previous.bottom + .2) return;
      used = origin + blank.height + blank.spacing;
    }
    if (used + current.outerTop + current.end - current.start > previous.bottom + .5) return current.page + 1;
  } catch { return; }
}

/** Witness a whole inline table after a source body paragraph and its
 * terminal blank. The stored single blank must fit, with no room for the
 * original table's complete row after it. */
export function sourceKeptInlineTableAfterBody(doc: HwpDocument, section: number, paragraph: number, control: number) {
  if (paragraph < 2) return;
  try {
    if (doc.getParagraphLength(section, paragraph - 1) ||
        !doc.getTextRange(section, paragraph - 2, 0, 0xffffffff).trim()) return;
    const table = sourceInlineTable(doc, section, paragraph, control, 2, 'Para');
    if (!table || table.rows !== 1 || !table.page || Math.abs(table.line.verticalPosition) > .1 ||
        Math.abs(table.start - table.top - table.outerTop) > .2 || table.outerTop < 0 || table.outerBottom < 0 ||
        Math.abs(table.line.height - (table.end - table.start) - table.outerTop - table.outerBottom) > .15) return;

    const preceding = JSON.parse(doc.getPageInfo(table.page - 1));
    const top = Number(preceding.marginTop) + Number(preceding.marginHeader);
    const bottom = Number(preceding.footerArea?.y ?? (preceding.height - preceding.marginBottom - preceding.marginFooter));
    if (preceding.columns?.length !== 1 || !Number.isFinite(top) || !(bottom > top)) return;

    const [body, blank] = [paragraph - 2, paragraph - 1].map(index => ({
      page: JSON.parse(doc.getPageOfPosition(section, index)),
      properties: JSON.parse(doc.getParaPropertiesAt(section, index)),
      metrics: JSON.parse(doc.getParagraphLineMetrics(section, index, '[]')),
    }));
    if ([body, blank].some(({ page, properties, metrics }) => !page.ok || page.page !== table.page - 1 ||
        properties.pageBreakBefore || properties.keepLines || properties.keepWithNext || !metrics.stored ||
        metrics.controlCount || metrics.lines?.length !== 1)) return;
    if (blank.properties.lineSpacingType !== 'Percent' || blank.properties.indent || blank.properties.marginLeft ||
        blank.properties.marginRight || blank.properties.spacingBefore || blank.properties.spacingAfter) return;

    const [last] = body.metrics.lines, [empty] = blank.metrics.lines;
    const after = Number(body.properties.spacingAfter || 0) / 2;
    const blankEnd = top + empty.verticalPosition + empty.height + empty.spacing;
    if (!(last.verticalPosition >= 0 && last.height > 0 && last.spacing >= 0 && after >= 0 &&
          empty.height > 0 && empty.spacing >= 0 && empty.textStart === 0) ||
        Math.abs(empty.verticalPosition - last.verticalPosition - last.height - last.spacing - after) > .15 ||
        Math.abs(empty.width - table.line.width) > .15 || !(blankEnd <= bottom + .2)) return;
    if (blankEnd + table.outerTop + table.end - table.start > bottom + .5) return table.page + 1;
  } catch { return; }
}

/** The last blank still owns its full advance. When the next paragraph
 * starts an authored page, only the blank line's em must fit before it. */
export function sourceTerminalEmptyLineAfterTable(doc: HwpDocument, section: number, paragraph: number,
  pitches: Map<string, number>, font: number, base: number, percent: number): HwpBodyLine | null {
  if (!(font > 0 && percent > 100 && percent <= 200 && base > font && base <= font * 2)) return null;
  const key = String(section) + ':' + paragraph, next = String(section) + ':' + (paragraph + 1);
  const page = pitches.get('page:' + key), y = pitches.get('top:' + key);
  if (page === undefined || y === undefined || pitches.get('page:' + next) !== page + 1) return null;
  try {
    if (doc.getParagraphLength(section, paragraph + 1)) return null;
    const preceding = sourceInlineTable(doc, section, paragraph - 1, 0, 0);
    if (!preceding || preceding.page !== page || !(y + font <= preceding.bottom + .1 && y + base > preceding.bottom + .5)) return null;
    const here = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const following = JSON.parse(doc.getParagraphLineMetrics(section, paragraph + 1, '[]'));
    if (!here?.stored || here.controlCount || here.lines?.length !== 1 ||
        !following?.stored || following.controlCount || following.lines?.length !== 1 ||
        Math.abs(following.lines[0].verticalPosition) > .1) return null;
    const line = here.lines[0], before = preceding.line;
    if (Math.abs(line.height - font) > .08 || Math.abs(line.textHeight - font) > .08 ||
        Math.abs(line.height + line.spacing - base) > .08 ||
        Math.abs(line.verticalPosition - before.verticalPosition - before.height - before.spacing) > .12 ||
        Math.abs(line.verticalPosition - (y - preceding.top)) > .12) return null;
    return { percent, font, base, empty: true };
  } catch { return null; }
}

/** A stored empty line can end a page by its em when only the extra
 * percentage leading lies outside the native body. Verify that boundary
 * before carrying the behavior into the editable document. */
export function sourceTerminalEmptyLine(doc: HwpDocument, section: number, paragraph: number,
  pitches: Map<string, number>, font: number, base: number, percent: number): HwpBodyLine | null {
  const key = `${section}:${paragraph}`;
  const previous = `${section}:${paragraph - 1}`, next = `${section}:${paragraph + 1}`;
  const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`), previousY = pitches.get(`top:${previous}`);
  if (page === undefined || y === undefined || previousY === undefined || pitches.get(`page:${previous}`) !== page ||
      pitches.get(`page:${next}`) !== page + 1 || Math.abs(y - previousY - base) > .12 ||
      !(font > 0 && percent > 100 && percent <= 200 && base > font)) return null;
  try {
    // Require a continuous same-pitch empty run leading into real body text
    // on the next page. A separate space at an object or another empty first
    // line lacks the uninterrupted native flow observed here.
    if (doc.getParagraphLength(section, paragraph - 1) || !doc.getParagraphLength(section, paragraph + 1)) return null;
    const info = JSON.parse(doc.getPageInfo(page)), nextPage = JSON.parse(doc.getPageInfo(page + 1));
    const bottom = Number(info.height) - Number(info.marginBottom) - Number(info.marginFooter);
    const nextTop = Number(nextPage.marginTop) + Number(nextPage.marginHeader);
    if (info.columns?.length !== 1 || nextPage.columns?.length !== 1 ||
      !(Math.abs((pitches.get(`top:${next}`) ?? NaN) - nextTop) < .12 &&
        y + font <= bottom + .1 && y + base > bottom + .5)) return null;
    for (const index of [paragraph - 1, paragraph]) {
      const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, index, '[]'));
      if (!stored || controlCount || lines?.length !== 1) return null;
      const { height, textHeight, spacing } = lines[0];
      if (![height, textHeight, spacing].every(Number.isFinite) || Math.abs(height - font) > .08 ||
        Math.abs(textHeight - font) > .08 || Math.abs(height + spacing - base) > .08) return null;
    }
    return { percent, font, base, empty: true };
  } catch { return null; }
}

export function hwpBodyTextAtLineTop(doc: HwpDocument, section: number, paragraph: number, font: number, advance: number, mode: 'body' | 'continued' | 'column-end' = 'body'): boolean {
  type Line = { height: number; textHeight: number; baseline: number; spacing: number; verticalPosition: number };
  const storedLines = (index: number): Line[] | undefined => {
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, index, '[]')) as {
      stored: boolean; controlCount: number; lines?: Line[];
    };
    if (!stored || controlCount || !lines?.length) return;
    const even = lines.every((line, i) =>
      (i === 0 || Math.abs(line.verticalPosition - lines[i - 1].verticalPosition - advance) < .1) &&
      Math.abs(line.height - font) < .1 && Math.abs(line.textHeight - font) < .1 &&
      line.baseline > 0 && line.baseline <= line.textHeight &&
      Math.abs(line.height + line.spacing - advance) < .1);
    return even ? lines : undefined;
  };
  try {
    // Consecutive source lines confirm that the native em-height text box
    // precedes its extra leading, including across a paragraph boundary.
    const lines = storedLines(paragraph);
    if (!lines) return false;
    if (mode === 'column-end') return lines.length === 1;
    if (lines.length > 1) return true;
    if (!paragraph) return false;
    const previous = storedLines(paragraph - 1);
    if (!previous || (previous.length < 2 && mode !== 'continued')) return false;
    const last = previous[previous.length - 1];
    return Math.abs(lines[0].verticalPosition - last.verticalPosition - advance) < .1 &&
      Math.abs(lines[0].baseline - last.baseline) < .1;
  } catch { return false; }
}

/** A natural page end can fit the last line's em, but not its unused leading. */
/** A one-line mixed-size paragraph can reserve the original physical
 * pitch when the stored advance naturally leads to its next paragraph. */
export function sourceStoredMixedLine(doc: HwpDocument, section: number, paragraph: number,
  pitches: Map<string, number>, font: number, base: number): boolean {
  const key = `${section}:${paragraph}`;
  const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`);
  if (page === undefined || y === undefined || pitches.has(`cross-page:${key}`)) return false;
  try {
    const here = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const next = JSON.parse(doc.getParagraphLineMetrics(section, paragraph + 1, '[]'));
    const info = JSON.parse(doc.getPageInfo(page));
    if (!here.stored || here.controlCount || here.lines?.length !== 1 || !next.stored ||
        next.controlCount || !next.lines?.length || info.columns?.length !== 1) return false;
    const line = here.lines[0], bodyTop = Number(info.marginTop) + Number(info.marginHeader);
    const bodyBottom = Number(info.height) - Number(info.marginBottom) - Number(info.marginFooter);
    return line.textStart === 0 && Math.abs(line.height - font) <= .06 && Math.abs(line.textHeight - font) <= .06 &&
      Math.abs(line.height + line.spacing - base) <= .06 &&
      Math.abs(bodyTop + line.verticalPosition - y) <= .15 && y + base <= bodyBottom &&
      Math.abs(next.lines[0].verticalPosition - line.verticalPosition - base) <= .12;
  } catch { return false; }
}

/** A mixed pair can use unequal physical advances even when the first line
 * alone sets the paragraph's usual pitch. Confirm both stored starts and the
 * real neighboring paragraphs before carrying those advances into print. */
export function sourceMixedBodyLinePair(doc: HwpDocument, section: number, paragraph: number,
  pitches: Map<string, number>, font: number, base: number, percent: number): HwpStoredMixedLinePair | undefined {
  const key = `${section}:${paragraph}`, before = `${section}:${paragraph - 1}`, after = `${section}:${paragraph + 1}`;
  const page = pitches.get(`page:${key}`);
  if (page === undefined || pitches.get(`page:${before}`) !== page || pitches.get(`page:${after}`) !== page ||
      pitches.has(`cross-page:${key}`) || pitches.get(`lines:${key}`) !== 2) return;
  try {
    const papers = JSON.parse(doc.getPageInfo(page));
    if (papers.columns?.length !== 1) return;
    const bodyTop = papers.marginTop + papers.marginHeader;
    const read = (index: number) => JSON.parse(doc.getParagraphLineMetrics(section, index, '[]'));
    const here = read(paragraph), prior = read(paragraph - 1), next = read(paragraph + 1);
    if (!here.stored || here.controlCount || here.lines?.length !== 2 ||
        !prior.stored || prior.controlCount || !prior.lines?.length ||
        !next.stored || next.controlCount || !next.lines?.length) return;
    const [first, second] = here.lines, previous = prior.lines.at(-1), following = next.lines[0];
    const firstPitch = first.height + first.spacing, secondPitch = second.height + second.spacing;
    if (!((first.textStart === 0 && second.textStart > 0) &&
          Math.abs(first.height - font) <= .06 && first.height > second.height &&
          Math.abs(firstPitch - base) <= .06 && Math.abs(first.width - second.width) <= .06 &&
          [first, second].every(line => line.baseline > 0 && line.baseline <= line.height &&
            Math.abs(line.height - line.textHeight) <= .06 && Math.abs(line.height * (percent / 100 - 1) - line.spacing) <= .06))) return;
    const y = (at: string) => pitches.get(`top:${at}`) ?? Infinity;
    if (Math.abs(bodyTop + prior.lines[0].verticalPosition - y(before)) > .15 ||
        Math.abs(bodyTop + first.verticalPosition - y(key)) > .15 ||
        Math.abs(bodyTop + following.verticalPosition - y(after)) > .15 ||
        Math.abs(previous.verticalPosition + previous.height + previous.spacing - first.verticalPosition) > .08 ||
        Math.abs(first.verticalPosition + firstPitch - second.verticalPosition) > .08 ||
        Math.abs(second.verticalPosition + secondPitch - following.verticalPosition) > .08) return;
    return [{ pitch: firstPitch, baseline: first.baseline }, { pitch: secondPitch, baseline: second.baseline }];
  } catch { /* Missing stored measurements use ordinary paragraph flow. */ }
}

/** In a mixed-size paragraph the browser can add leading even when both
 * stored native lines consumed the same advance. Verify the neighboring
 * lines and their authored margins before carrying that pitch into DOCX. */
export function sourceUniformMixedLinePair(doc: HwpDocument, section: number, paragraph: number,
  pitches: Map<string, number>, font: number, base: number, properties: Record<string, any>): boolean {
  const key = (p: number) => `${section}:${p}`;
  const page = pitches.get(`page:${key(paragraph)}`);
  if (page === undefined || pitches.get(`lines:${key(paragraph)}`) !== 2 ||
      [paragraph - 1, paragraph, paragraph + 1].some(p => pitches.get(`page:${key(p)}`) !== page ||
        pitches.has(`cross-page:${key(p)}`))) return false;
  try {
    const paper = JSON.parse(doc.getPageInfo(page));
    if (paper.columns?.length !== 1) return false;
    const read = (p: number) => JSON.parse(doc.getParagraphLineMetrics(section, p, '[]'));
    const prior = read(paragraph - 1), own = read(paragraph), next = read(paragraph + 1);
    if ([prior, own, next].some(p => !p.stored || p.controlCount || !p.lines?.length) || own.lines.length !== 2) return false;
    const [first, second] = own.lines, previous = prior.lines.at(-1), following = next.lines[0];
    const previousProps = JSON.parse(doc.getParaPropertiesAt(section, paragraph - 1));
    const nextProps = JSON.parse(doc.getParaPropertiesAt(section, paragraph + 1));
    const before = Math.max(properties.spacingBefore || 0, previousProps.spacingAfter || 0) / 2;
    const after = Math.max(properties.spacingAfter || 0, nextProps.spacingBefore || 0) / 2;
    const y = (p: number) => pitches.get(`top:${key(p)}`) ?? Infinity;
    const bodyTop = paper.marginTop + paper.marginHeader;
    return first.textStart === 0 && second.textStart > 0 && Math.abs(first.width - second.width) <= .06 &&
      [first, second].every(line => Math.abs(line.height - font) <= .06 &&
        Math.abs(line.textHeight - font) <= .06 && line.baseline > 0 && line.baseline <= font &&
        Math.abs(line.height + line.spacing - base) <= .06) &&
      Math.abs(first.baseline - second.baseline) <= .06 &&
      Math.abs(bodyTop + first.verticalPosition - y(paragraph)) <= .15 &&
      Math.abs(bodyTop + following.verticalPosition - y(paragraph + 1)) <= .15 &&
      Math.abs(first.verticalPosition - previous.verticalPosition - previous.height - previous.spacing - before) <= .08 &&
      Math.abs(second.verticalPosition - first.verticalPosition - base) <= .08 &&
      Math.abs(following.verticalPosition - second.verticalPosition - base - after) <= .08;
  } catch { return false; }
}

export function sourceTerminalTextAtLineTop(doc: HwpDocument, section: number, paragraph: number,
  pitches: Map<string, number>, font: number, advance: number): boolean {
  if (!(font > 0 && advance > font && advance <= font * 2)) return false;
  const key = `${section}:${paragraph}`, prior = `${section}:${paragraph - 1}`, next = `${section}:${paragraph + 1}`;
  const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`), nextY = pitches.get(`top:${next}`);
  if (page === undefined || y === undefined || nextY === undefined || pitches.get(`lines:${key}`) !== 1 ||
    pitches.get(`page:${prior}`) !== page || pitches.get(`cross-page:${prior}`) ||
    pitches.get(`page:${next}`) !== page + 1 || Math.abs((pitches.get(`height:${key}`) ?? 0) - font) > sourceCellBoundsTolerance)
    return false;
  try {
    const paper = JSON.parse(doc.getPageInfo(page)), following = JSON.parse(doc.getPageInfo(page + 1));
    if (paper.columns?.length !== 1 || following.columns?.length !== 1 ||
        Math.abs(nextY - following.marginTop - following.marginHeader) > sourceCellBoundsTolerance ||
        JSON.parse(doc.getParaPropertiesAt(section, paragraph + 1)).pageBreakBefore) return false;
    const bottom = Number(paper.footerArea?.y ?? (paper.height - paper.marginBottom - paper.marginFooter));
    const clearance = bottom - y - font;
    if (!(clearance >= -.08 && clearance <= 1.5 && y + advance > bottom + .5)) return false;
    const previous = JSON.parse(doc.getParagraphLineMetrics(section, paragraph - 1, '[]'));
    const here = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const last = previous.lines?.at(-1), line = here.lines?.[0];
    return !!(previous.stored && !previous.controlCount && last && here.stored && !here.controlCount &&
      here.lines.length === 1 && line?.textStart === 0 && line.baseline > 0 && line.baseline <= font &&
      Math.abs(line.height - font) < .08 && Math.abs(line.textHeight - font) < .08 &&
      Math.abs(line.height + line.spacing - advance) < .08 &&
      Math.abs(line.verticalPosition - last.verticalPosition - last.height - last.spacing) <= sourceCellBoundsTolerance &&
      Math.abs(y - paper.marginTop - paper.marginHeader - line.verticalPosition) <= sourceCellBoundsTolerance);
  } catch { return false; }
}

/** A text line follows the authored trailing leading of an intact inline
 * table. Keep its text at the measured line origin when only its unused
 * leading would fall outside the source body. */
export function sourceTextAfterInlineTableFits(doc: HwpDocument, section: number, paragraph: number,
  pitches: ReadonlyMap<string, number>, font: number, advance: number): boolean {
  if (paragraph < 1 || paragraph + 1 >= doc.getParagraphCount(section) ||
    !(font > 0 && advance > font && advance <= font * 2 + .08)) return false;
  const key = `${section}:${paragraph}`, next = `${section}:${paragraph + 1}`;
  const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`), nextY = pitches.get(`top:${next}`);
  if (page === undefined || y === undefined || nextY === undefined || pitches.get(`lines:${key}`) !== 1 ||
      pitches.get(`page:${next}`) !== page + 1 || Math.abs((pitches.get(`height:${key}`) ?? 0) - font) > sourceCellBoundsTolerance)
    return false;
  try {
    const source = sourceInlineTable(doc, section, paragraph - 1, 0, 0, 'Para');
    if (!source || source.rows < 2 || source.page !== page || !source.fillsColumn || source.outerTop !== 0 || source.outerBottom !== 0 ||
        !(source.line.spacing > 0) || Math.abs(source.start - source.top - source.line.verticalPosition) > .15 ||
        Math.abs(source.end - source.start - source.line.height) > .15 ||
        Math.abs(y - source.end - source.line.spacing) > .15) return false;
    const followingPage = JSON.parse(doc.getPageInfo(page + 1));
    if (followingPage.columns?.length !== 1 || Math.abs(nextY - followingPage.marginTop - followingPage.marginHeader) > .15) return false;
    const remaining = source.bottom - y - font;
    if (!(remaining >= -.08 && remaining <= font / 4 && y + advance > source.bottom + .5)) return false;
    const here = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const following = JSON.parse(doc.getParagraphLineMetrics(section, paragraph + 1, '[]'));
    const [line] = here.lines ?? [];
    return !!(here.stored && !here.controlCount && here.lines?.length === 1 && line &&
      line.textStart === 0 && line.baseline > 0 && line.baseline <= font &&
      Math.abs(line.height - font) <= .08 && Math.abs(line.textHeight - font) <= .08 &&
      Math.abs(line.height + line.spacing - advance) <= .08 &&
      Math.abs(line.verticalPosition - source.line.verticalPosition - source.line.height - source.line.spacing) <= .12 &&
      Math.abs(y - source.top - line.verticalPosition) <= .15 &&
      following.stored && following.lines?.length === 1 && Math.abs(following.lines[0].verticalPosition) <= .08);
  } catch { return false; }
}

export function hwpParagraphBorderBox(properties: Record<string, any>): string {
  // The source feature is a connected, uniformly stroked paragraph frame.
  // Leave separated boxes, offsets, fills, and hanging indents to their own
  // placement rules rather than guessing a matching frame.
  const margins = [properties.marginLeft, properties.marginRight];
  const offsets = properties.borderSpacing;
  if (properties.borderConnect !== true || properties.borderIgnoreMargin !== false || properties.fillType !== 'none' ||
      properties.headType !== 'None' || properties.spacingBefore || properties.spacingAfter || properties.indent < 0 ||
      !Array.isArray(offsets) || offsets.length !== 4 || offsets.some(offset => offset !== 0) ||
      margins.some(value => !Number.isFinite(value) || value < 0 || value > 31 * 4 / 3)) return '';
  const first = properties.borderLeft;
  if (!first || first.type !== 1 || !Number.isInteger(first.width) ||
      !/^#[a-f0-9]{6}$/i.test(first.color) ||
      ['Right', 'Top', 'Bottom'].some(side => {
        const edge = properties['border' + side];
        return !edge || edge.type !== first.type || edge.width !== first.width || edge.color !== first.color;
      })) return '';
  const box = parseParagraphBorderBox(JSON.stringify({
    size: Math.round(hwpBorderWidthsMillimeters[first.width] * 72 / 25.4 * 8),
    color: first.color.slice(1), leftSpace: Math.round(margins[0] * 3 / 4), rightSpace: Math.round(margins[1] * 3 / 4),
  }));
  return box ? ' ' + paragraphBorderBoxAttribute + '="' + escape(JSON.stringify(box)) + '"' : '';
}

export function connectedHwpBorderGroup(doc: HwpDocument, section: number, first: number, count: number,
  properties: Record<string, any>, continues: (index: number) => boolean, supported: (index: number) => boolean) {
  const border = hwpParagraphBorderBox(properties);
  let end = first;
  if (!border) return { end, attribute: '' };
  let complete = supported(first);
  while (end + 1 < count && continues(end + 1)) {
    const next = JSON.parse(doc.getParaPropertiesAt(section, end + 1));
    if (next.pageBreakBefore || next.borderFillId !== properties.borderFillId || hwpParagraphBorderBox(next) !== border) break;
    end++;
    if (!supported(end)) complete = false;
  }
  // A singleton or a subset of a source group must not acquire new caps.
  return { end, attribute: complete && end > first ? border : '' };
}

export function paragraphCss(properties: Record<string, any>, char?: Record<string, any>, cap?: number, includeSpacing = true): string {
  const fontPx = Number(char?.fontSize) / 75;
  const line = Number(properties.lineSpacing);
  // The query API mixes units: margins/indent and fixed leading are
  // already physical pixels, but before/after retain ParaShape's 2x scale.
  // See build_para_properties_json and style_resolver at the corpus revision.
  const height = physicalLineHeight(properties, char);
  const indent = Number(properties.indent) || 0;
  const left = (Number(properties.marginLeft) || 0) + Math.max(0, -indent);
  const spacing = Number.isFinite(cap) && Number.isFinite(height) && height > cap! && cap! >= fontPx ? `${cap}px`
    : properties.lineSpacingType === 'Percent' ? `${line}%` : Number.isFinite(height) ? `${height}px` : '';
  return `margin-left:${left}px;margin-right:${(Number(properties.marginRight) || 0)}px;text-indent:${indent}px;` +
    (['left', 'center', 'right', 'justify'].includes(properties.alignment) ? `text-align:${properties.alignment};` : '') +
    (includeSpacing ? `margin-top:${(Number(properties.spacingBefore) || 0) / 2}px;margin-bottom:${(Number(properties.spacingAfter) || 0) / 2}px;` : '') +
    (char?.fontSize ? `font-size:${char.fontSize / 100}pt;font-family:${escape(JSON.stringify(char.fontFamily || 'serif'))};` : '') +
    (spacing ? `line-height:${spacing};` : '') +
    // HWP often permits a single first/last line at a page boundary. Retain
    // that explicit policy instead of enabling Word's two-line protection.
    (typeof properties.widowOrphan === 'boolean' ? `widows:${properties.widowOrphan ? 2 : 1};orphans:${properties.widowOrphan ? 2 : 1};` : '') +
    // The native numeric API uses 0 for Korean word boundaries and 1 for
    // character boundaries. Do not infer this from HWPX's Latin enum names.
    (properties.koreanBreakUnit === 0 ? 'word-break:keep-all;' : '');
}

/** Use a stored single line and a painted suffix with one authored scale.
 * This keeps the source advance available when a substitute face wraps it. */
export function sourceScaledCellRun(doc: HwpDocument, section: number, paragraph: number, controlIndex: number, cellIndex: number,
  outer: number, width: number, painted: HwpPaintedLine, runs: readonly Record<string, any>[]): HwpSourceWrap | undefined {
  const address = JSON.stringify([{ controlIndex, cellIndex, cellParaIndex: 0 }]);
  const text = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
  if (text.length > 160 || text.length !== [...text].length || /[\p{Cc}\p{Cf}\u0590-\u08ff\ufffc]/u.test(text)) return;
  const end = normalizeHwpWrapText(text).length;
  if (end < 2) return;
  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address));
  if (!metrics.stored || metrics.controlCount || metrics.lines?.length !== 1 ||
      Math.abs(metrics.lines[0].width - width) > sourceCellBoundsTolerance) return;
  const shapes = JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, address, 0, text.length)) as
    { startOffset: number; endOffset: number }[];
  if (shapes.length < 2 || shapes.at(-1)?.endOffset !== text.length) return;
  const properties = (offset: number) => JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, address, offset));
  if (!shapes.slice(0, -1).every(shape => properties(shape.startOffset).ratios?.every((ratio: number) => ratio === 100))) return;
  const { startOffset: start } = shapes.at(-1)!;
  const shape = properties(start), scale = shape.ratios?.[0], size = shape.fontSize / 75;
  if (!Number.isInteger(scale) || scale < 50 || scale >= 100 || !shape.ratios.every((ratio: number) => ratio === scale) ||
      !shape.relativeSizes?.every((percent: number) => percent === 100) || !shape.fontFamily || size < 4 || size > 120) return;
  const glyphs = runs.filter(run => run.secIdx === section && run.parentParaIdx === paragraph &&
    (run.cellPath?.length ?? 0) <= 1 && (run.cellPath?.[0]?.controlIndex ?? run.controlIdx) === controlIndex &&
    (run.cellPath?.[0]?.cellIndex ?? run.cellIdx) === cellIndex &&
    (run.cellPath?.[0]?.cellParaIndex ?? run.cellParaIdx) === 0 && typeof run.text === 'string');
  if (glyphs.length < 2 || glyphs.some(run => run.y !== painted.y) ||
      glyphs.map(run => run.text).join('') !== text) return;
  const suffix = glyphs.at(-1)!;
  const left = glyphs[0].x, endX = suffix.x + suffix.w;
  if (suffix.charStart !== start || suffix.text !== text.slice(start) || !Number.isFinite(left) ||
      !Number.isFinite(suffix.w) || suffix.w <= 0 || endX - left > width + .1 || suffix.w >= width) return;
  return { fingerprint: hwpWrapFingerprint(text), ends: [end], cellWidth: { outer, text: Math.round(width * 1000) / 1000 },
    scaleRun: { start, scale, width: suffix.w, size, font: shape.fontFamily } };
}

/** A figure crossing the entire HWP body can reserve blank paragraph
 * lines even though its own anchor contains no inline text. */
export function sourceBodyPictureFlow(doc: HwpDocument, section: number, paragraph: number,
  picture: { page: number; x: number; y: number; w: number; h: number }, pitches: Map<string, number>): HwpBodyPictureFlow | undefined {
  const next = paragraph + 1, nextKey = `${section}:${next}`;
  if (next >= doc.getParagraphCount(section) || pitches.get(`page:${nextKey}`) !== picture.page ||
      doc.getTextRange(section, paragraph, 0, 0xffffffff).trim() ||
      doc.getTextRange(section, next, 0, 0xffffffff).trim()) return;
  const page = JSON.parse(doc.getPageInfo(picture.page)), bodyTop = page.marginTop + page.marginHeader;
  if (Math.abs(picture.x - page.bodyLeft) > sourceCellBoundsTolerance ||
      Math.abs(picture.x + picture.w - page.bodyRight) > sourceCellBoundsTolerance ||
      picture.y + picture.h > page.footerArea.y + sourceCellBoundsTolerance) return;
  const first = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  const following = JSON.parse(doc.getParagraphLineMetrics(section, next, '[]'));
  if (!first.stored || !following.stored || first.controlCount !== 1 || following.controlCount ||
      first.lines?.length !== 1 || following.lines?.length !== 1 || first.lines[0].width !== 0) return;
  const anchorY = bodyTop + first.lines[0].verticalPosition, followingEm = following.lines[0].textHeight;
  const followingY = pitches.get(`top:${nextKey}`);
  if (followingY === undefined || !(bodyTop <= anchorY && anchorY < picture.y && followingY > anchorY) ||
      !(followingEm >= 4 && followingEm <= 120) ||
      Math.abs(followingY + followingEm - picture.y - picture.h) > sourceCellBoundsTolerance) return;
  return { anchorY: Math.round(anchorY * 1000) / 1000, height: picture.h,
    followingEm: Math.round(followingEm * 1000) / 1000 };
}

/** Preserve the source outer margin of a page-leading inline group when
 * the stored line and painted position confirm the same amount of space. */
export function sourceInlineGroupReserve(doc: HwpDocument, section: number, paragraph: number, control: number,
  group: { page: number; y: number; h: number }, previewHeight: number) {
  try {
    const shape = JSON.parse(doc.getShapeProperties(section, paragraph, control));
    if (!shape.treatAsChar || shape.outerMarginBottom || !(shape.outerMarginTop > 0)) return;
    const paper = JSON.parse(doc.getPageInfo(group.page));
    if (paper.columns?.length !== 1 || !paper.footerArea) return;
    const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    if (!metrics.stored || metrics.controlCount !== 1 || metrics.lines?.length !== 1) return;
    const { verticalPosition, height } = metrics.lines[0];
    const bodyTop = paper.marginTop + paper.marginHeader, shapeHeight = shape.height / 75;
    const padding = (previewHeight - shapeHeight) / 2, reserve = shape.outerMarginTop / 75;
    const close = (a: number, b: number) => Math.abs(a - b) <= sourceCellBoundsTolerance;
    if (!close(verticalPosition, 0) || !close(group.h, shapeHeight) || !close(height, shapeHeight + reserve) ||
        !(padding >= 0 && padding < reserve) || !close(bodyTop + reserve, group.y) ||
        !(group.y + group.h + padding <= paper.footerArea.y)) return;
    return reserve - padding;
  } catch { /* Preserve ordinary inline flow without the complete source layout. */ }
}

/** Native paper artwork shares the opening paragraph with live text. The
 * stored text lines show which font determined each percentage advance. */
export function sourcePaperHeadingLine(doc: HwpDocument, section: number, paragraph: number, pageIndex: number,
  text: string, percent: number, font: number, base: number): HwpBodyLine | undefined {
  if (!text.trim() || !(percent > 100 && percent <= 200 && font > 0 && base <= 240) ||
      Math.abs(base - font * percent / 100) > .06) return;
  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  const lines = metrics.lines as (StoredCellLine & { width: number })[] | undefined;
  if (!metrics.stored || !lines || lines.length < 2 || lines.length > 12 || !lines.some(line => line.textHeight < font - .1)) return;
  const page = JSON.parse(doc.getPageInfo(pageIndex)), top = Number(page.marginTop) + Number(page.marginHeader);
  const width = Number(page.columns?.[0]?.width);
  if (!(width > font && width <= 10000) || !(lines[0].verticalPosition > 0)) return;
  const painted = (JSON.parse(doc.getPageTextLayout(pageIndex)).runs ?? []).filter((run: any) =>
    run.secIdx === section && run.paraIdx === paragraph && run.parentParaIdx === undefined && !run.headerFooter && run.text?.trim());
  if (!painted.length || !lines.every((line, index) => {
    const y = top + line.verticalPosition, pitch = line.height + line.spacing;
    return Math.abs(line.width - width) <= .1 && Math.abs(line.height - line.textHeight) <= .1 &&
      Math.abs(pitch - line.textHeight * percent / 100) <= .1 &&
      (!index || Math.abs(lines[index - 1].verticalPosition + lines[index - 1].height + lines[index - 1].spacing -
        line.verticalPosition) <= .1) &&
      painted.some((run: any) => Math.abs(run.y - y) <= .1 && Math.abs(run.h - line.textHeight) <= .1);
  })) return;
  return { percent, font, base, paperHeading: { width: lines[0].width, fingerprint: hwpWrapFingerprint(text) } };
}

export function sourceTightInlinePicture(doc: HwpDocument, section: number, paragraph: number, control: number,
  position: { page: number; x: number; y: number; w: number; h: number }): boolean {
  try {
    const props = JSON.parse(doc.getPictureProperties(section, paragraph, control));
    if (!props.treatAsChar || props.textWrap !== 'TopAndBottom' || props.allowOverlap || !props.restrictInPage ||
        props.hasCaption || props.vertOffset || props.horzOffset || props.outerMarginTop || props.outerMarginBottom ||
        Math.abs(Number(props.height) / 75 - position.h) > .16) return false;
    const page = JSON.parse(doc.getPageInfo(position.page));
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
    const line = lines?.[0], room = Number(page.footerArea?.y) - position.y - position.h;
    return stored && controlCount === 1 && lines.length === 1 && line.textStart === 0 &&
      Math.abs(line.height - position.h) <= .16 && Math.abs(line.textHeight - position.h) <= .16 &&
      room >= 0 && room < line.spacing &&
      page.columns?.some((column: { x: number; width: number }) =>
        position.x >= column.x - .1 && position.x + position.w <= column.x + column.width + .1);
  } catch { return false; }
}

/** Retain the empty host's left inset when its one inline table is laid out
 * independently inside a source column. */
export function sourceInlineColumnTableIndent(doc: HwpDocument, section: number, paragraph: number, control: number,
  position: { page: number; x: number; y: number; w: number; h: number }, margin: number): number | undefined {
  const { stored, controlCount, tables, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, '[]'));
  const line = lines?.[0];
  const table = tables?.[0];
  if (!stored || controlCount !== 1 || lines?.length !== 1 || tables?.length !== 1 ||
      !table?.inline || table.index !== control || line.textStart !== 0 ||
      Math.abs(line.width - position.w) > sourceCellBoundsTolerance ||
      Math.abs(line.height - position.h) > sourceCellBoundsTolerance) return;
  const page = JSON.parse(doc.getPageInfo(position.page));
  if (page.columns?.length !== 2 || position.y < Number(page.marginTop) + Number(page.marginHeader) ||
      position.y + position.h > Number(page.footerArea?.y) + sourceCellBoundsTolerance) return;
  const column = page.columns.find((c: { x: number; width: number }) =>
    Math.abs(position.x - c.x - margin) <= sourceCellBoundsTolerance &&
    Math.abs(position.x + position.w - c.x - c.width) <= sourceCellBoundsTolerance);
  if (column) return margin;
}

export type HwpPaintedLine = { y: number; h: number; page: number; charStart?: number; shortLabelInk?: number };
type StoredCellLine = { verticalPosition: number; textStart: number; height: number; textHeight: number; spacing: number };
export const sourceCellBoundsTolerance = .16;

export function sourceLinePitches(doc: HwpDocument, splitCellLines?: Map<string, HwpPaintedLine[]>,
  bodyLines?: Map<string, HwpPaintedLine[]>): Map<string, number> {
  const lines = new Map<string, HwpPaintedLine[]>();
  const texts = new Map<string, () => string>();
  const pitches = new Map<string, number>();
  const cellPages = new Map<string, number>();
  for (let page = 0; page < doc.pageCount(); page++) for (const run of JSON.parse(doc.getPageTextLayout(page)).runs ?? []) {
    if (run.parentParaIdx === undefined && typeof run.secIdx === 'number' && typeof run.paraIdx === 'number') pitches.set(`flow:${run.secIdx}:${run.paraIdx}`, 1);
    if (run.headerFooter || run.paraIdx === 0xffffffff || typeof run.secIdx !== 'number' || typeof run.paraIdx !== 'number') continue;
    const key = run.parentParaIdx === undefined ? `${run.secIdx}:${run.paraIdx}` : cellParagraphKey(run.secIdx, run.parentParaIdx, run.cellPath?.length ? run.cellPath : [{ controlIndex: run.controlIdx, cellIndex: run.cellIdx, cellParaIndex: run.cellParaIdx }]);
    if (!texts.has(key)) texts.set(key, () => run.parentParaIdx === undefined
      ? doc.getTextRange(run.secIdx, run.paraIdx, 0, doc.getParagraphLength(run.secIdx, run.paraIdx))
      : doc.getTextInCellByPath(run.secIdx, run.parentParaIdx, JSON.stringify(run.cellPath?.length ? run.cellPath : [{ controlIndex: run.controlIdx, cellIndex: run.cellIdx, cellParaIndex: run.cellParaIdx }]), 0, 0xffffffff));
    if (!pitches.has(`top:${key}`)) { pitches.set(`top:${key}`, run.y); pitches.set(`page:${key}`, page); }
    pitches.set(`height:${key}`, Math.max(pitches.get(`height:${key}`) ?? 0, run.h));
    if (bodyLines && run.parentParaIdx === undefined && run.text && run.h > 0) {
      const painted = bodyLines.get(key) ?? [], previous = painted.at(-1);
      // In a hanging paragraph, the first run for a physical line can be
      // only authored spaces. Its character start is still meaningful.
      if (!previous || previous.page !== page || Math.abs(previous.y - run.y) > Math.min(previous.h, run.h) * .45)
        painted.push({ y: run.y, h: run.h, page, charStart: run.charStart });
      bodyLines.set(key, painted);
    }
    if (!run.text?.trim()) {
      // The source layer still gives leading spaces a horizontal origin.
      // Keep it separately from the column planner's first visible run.
      if (run.parentParaIdx === undefined && run.charStart === 0 && Number.isFinite(run.x) &&
          page === pitches.get(`page:${key}`) && run.y === pitches.get(`top:${key}`) && !pitches.has(`leading-x:${key}`))
        pitches.set(`leading-x:${key}`, run.x);
      continue;
    }
    if (run.parentParaIdx === undefined) {
      if (!pitches.has(`x:${key}`) && run.charStart === 0 && page === pitches.get(`page:${key}`) &&
          run.y === pitches.get(`top:${key}`) && Number.isFinite(run.x)) pitches.set(`x:${key}`, run.x);
      if (page !== pitches.get(`page:${key}`)) pitches.set(`cross-page:${key}`, 1);
    }
    if (run.parentParaIdx !== undefined && (run.cellPath?.length ?? 0) <= 1) {
      const cell = run.cellPath?.[0], control = cell?.controlIndex ?? run.controlIdx, index = cell?.cellIndex ?? run.cellIdx;
      if (Number.isInteger(control) && Number.isInteger(index)) {
        const key = `${run.secIdx}:${run.parentParaIdx}:${control}:${index}`;
        const prior = cellPages.get(key);
        if (prior === undefined) cellPages.set(key, page);
        else if (prior !== page) pitches.set(`continues:${key}`, 1);
      }
    }
    const group = lines.get(key) ?? [], previous = group.at(-1);
    if (!previous || previous.page !== page || Math.abs(previous.y - run.y) > Math.min(previous.h, run.h) * .45) {
      const line: HwpPaintedLine = { y: run.y, h: run.h, page, charStart: run.charStart };
      if (!previous && run.parentParaIdx !== undefined && run.charStart === 0 &&
          hwpShortCellLabel(run.text) && Number.isFinite(run.w) && run.w > 0) line.shortLabelInk = run.w;
      group.push(line);
    } else if (previous.shortLabelInk !== undefined) delete previous.shortLabelInk;
    lines.set(key, group);
  }
  for (const [key, group] of lines) {
    pitches.set(`lines:${key}`, group.length);
    if (splitCellLines && !key.includes('/') && key.split(':').length === 5) splitCellLines.set(key, group);
    let text: string[] | undefined;
    const values = group.slice(1).flatMap((line, i) => {
      const previous = group[i];
      if (line.page !== previous.page) return [];
      const distance = line.y - previous.y, height = Math.max(line.h, previous.h);
      // Retain the existing rejection of large control/object gaps before
      // normalizing blank lines; they must not become new pitch candidates.
      if (distance < height * .9 || distance >= height * 4) return [];
      let steps = 1;
      if (line.charStart !== undefined && previous.charStart !== undefined && line.charStart > previous.charStart) {
        try {
          // Paint-run character starts use native Unicode scalar offsets.
          // JS string slicing would stop early after supplementary Unicode
          // and mistake several explicit line pitches for one enlarged line.
          text ??= Array.from(texts.get(key)?.() ?? '');
          // Empty explicit lines have no painted run. Their combined distance
          // is several line pitches, rather than an enlarged single pitch.
          steps = Math.max(1, text.slice(previous.charStart, line.charStart).filter(char => char === '\n').length);
        } catch { /* Optional text metrics retain the measured fallback. */ }
      }
      const value = distance / steps;
      return value >= height * .9 && value < height * 4 ? [value] : [];
    }).sort((a, b) => a - b);
    if (values.length) pitches.set(key, Math.round(values[Math.floor(values.length / 2)] * 100) / 100);
  }
  return pitches;
}

function sourceCellBeforeOnPage(doc: HwpDocument, address: [number, number, number, number],
  pitches: Map<string, number>, page: number, count: number) {
  const [section, paragraph, controlIndex, cellIndex] = address;
  if (count < 2) return;
  let position = 0, terminalHeight = 0, paintedOrigin: number | undefined;
  const before: number[] = [];
  for (let index = 0; index < count; index++) {
    const key = `${section}:${paragraph}:${controlIndex}:${cellIndex}:${index}`;
    if (pitches.get(`page:${key}`) !== page) return;
    const properties = JSON.parse(doc.getCellParaPropertiesAt(...address, index));
    const gap = Number(properties.spacingBefore) / 2;
    if (!(gap > 0 && Number.isFinite(gap)) || Number(properties.spacingAfter) !== 0 ||
        properties.headType !== 'None' || properties.lineSpacingType !== 'Percent') return;
    const path = JSON.stringify([{ controlIndex, cellIndex, cellParaIndex: index }]);
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path)) as
      { stored: boolean; controlCount: number; lines: StoredCellLine[] };
    if (!stored || controlCount || !lines?.length || lines.length !== pitches.get(`lines:${key}`)) return;
    const font = Number(JSON.parse(doc.getCellCharPropertiesAt(...address, index, 0)).fontSize) / 75;
    const advance = physicalLineHeight(properties, { fontSize: font * 75 });
    const paintedTop = pitches.get(`top:${key}`);
    if (!(font > 0 && advance >= font) || paintedTop === undefined ||
        !(Math.abs((pitches.get(`height:${key}`) ?? NaN) - font) <= .1) ||
        (lines.length > 1 && !(Math.abs((pitches.get(key) ?? NaN) - advance) <= .1))) return;
    position += gap;
    paintedOrigin ??= paintedTop - lines[0].verticalPosition;
    if (!(Math.abs(paintedTop - paintedOrigin - position) <= .1)) return;
    for (const [lineIndex, line] of lines.entries()) {
      if (!(Math.abs(line.verticalPosition - position) <= .1 && Math.abs(line.height - font) <= .1 &&
          Math.abs(line.textHeight - font) <= .1 && Math.abs(line.height + line.spacing - advance) <= .1) ||
          (lineIndex === 0 ? line.textStart !== 0 : !(line.textStart > lines[lineIndex - 1].textStart))) return;
      terminalHeight = line.verticalPosition + line.textHeight;
      position = line.verticalPosition + line.height + line.spacing;
    }
    before.push(gap);
  }
  return paintedOrigin === undefined ? undefined : { before, terminalHeight, paintedOrigin };
}

/** An automatic row can be shorter than the residual table grid. Prove
 * its intrinsic blank content against the continuous stored line positions
 * and physical cell bounds before leaving its height to that content. */
export function sourceBlankAutoCellFits(doc: HwpDocument, address: [number, number, number, number],
  innerWidth: number, innerHeight: number): boolean {
  if (!(innerWidth > 0 && innerHeight > 0)) return false;
  const [section, paragraph, controlIndex, cellIndex] = address;
  const count = doc.getCellParagraphCount(...address);
  if (count < 2) return false;
  let nextTop = 0, bottom = 0;
  for (let cellParaIndex = 0; cellParaIndex < count; cellParaIndex++) {
    const path = JSON.stringify([{ controlIndex, cellIndex, cellParaIndex }]);
    if (doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff) !== '') return false;
    const properties = JSON.parse(doc.getCellParaPropertiesAt(...address, cellParaIndex));
    if (properties.lineSpacingType !== 'Percent' || properties.spacingBefore || properties.spacingAfter ||
        properties.headType !== 'None' || properties.pageBreakBefore || properties.keepWithNext) return false;
    const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
    if (!metrics.stored || metrics.controlCount !== 0 || metrics.lines?.length !== 1) return false;
    const [line] = metrics.lines;
    if (!(line.height > 0 && line.spacing >= 0 && line.textStart === 0 &&
        Math.abs(line.width - innerWidth) <= sourceCellBoundsTolerance &&
        Math.abs(line.verticalPosition - nextTop) <= sourceCellBoundsTolerance &&
        Math.abs(line.height - line.textHeight) <= sourceCellBoundsTolerance)) return false;
    bottom = line.verticalPosition + line.height;
    nextTop = bottom + line.spacing;
    if (bottom - innerHeight > sourceCellBoundsTolerance) return false;
  }
  return Math.abs(bottom - innerHeight) <= sourceCellBoundsTolerance;
}

/** Recover before-spacing where the stored and painted lines and the intact
 * native cell all confirm it. A line pitch measures only within a paragraph. */
export function sourceIntactCellBefore(doc: HwpDocument, address: [number, number, number, number],
  pitches: Map<string, number>, page: number, innerHeight: number): number[] | undefined {
  const result = sourceCellBeforeOnPage(doc, address, pitches, page, doc.getCellParagraphCount(...address));
  return result && Math.abs(result.terminalHeight - innerHeight) <= .1 ? result.before : undefined;
}

/** An inline table can occupy the second fragment without a painted paragraph
 * of its own. Verify that fragment before recovering the text gaps above it. */
export function sourceBeforeNestedCellTail(doc: HwpDocument, address: [number, number, number, number],
  pitches: Map<string, number>, first: { pageIndex: number; y: number; h: number },
  tail: { pageIndex: number; y: number; h: number; w: number }, top: number, bottom: number) {
  try {
    const [section, paragraph, controlIndex, cellIndex] = address;
    const last = doc.getCellParagraphCount(...address) - 1;
    const prefix = sourceCellBeforeOnPage(doc, address, pitches, first.pageIndex, last);
    if (!prefix || tail.pageIndex !== first.pageIndex + 1 || prefix.terminalHeight + top + bottom > first.h + .1 ||
        Math.abs(prefix.paintedOrigin - first.y - top) > .1) return;
    const cell = [{ controlIndex, cellIndex, cellParaIndex: last }];
    const path = JSON.stringify(cell);
    if (doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff) !== '') return;
    const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
    const properties = JSON.parse(doc.getCellParaPropertiesAt(...address, last));
    if (!metrics.stored || metrics.controlCount !== 1 || metrics.lines?.length !== 1 || metrics.tables?.length !== 1 ||
        properties.lineSpacingType !== 'Percent' || properties.spacingAfter || properties.headType !== 'None') return;
    const [table] = metrics.tables, [line] = metrics.lines;
    const gap = Number(properties.spacingBefore) / 2;
    if (table.inline !== true || !Number.isInteger(table.index) || table.index < 0 ||
        !(gap > 0) || line.textStart !== 0 || Math.abs(line.verticalPosition - gap) > .1 ||
        !(table.outerTop >= 0 && table.outerBottom >= 0)) return;
    const nested = JSON.stringify([...cell, { controlIndex: table.index, cellIndex: 0, cellParaIndex: 0 }]);
    const size = JSON.parse(doc.getTableDimensionsByPath(section, paragraph, nested));
    const boxes = JSON.parse(doc.getTableCellBboxesByPath(section, paragraph, nested));
    if (size.rowCount !== 1 || size.colCount !== 1 || size.cellCount !== 1 || boxes.length !== 1) return;
    const [box] = boxes;
    const boxHeight = box.h + table.outerTop + table.outerBottom;
    const geometry = [box.y, box.h, box.w, boxHeight, line.height, line.textHeight, line.verticalPosition, gap, tail.h];
    if (!geometry.every(Number.isFinite) || box.pageIndex !== tail.pageIndex || box.rowSpan !== 1 || box.colSpan !== 1 ||
        !(box.h > 0 && box.w > 0 && box.w <= tail.w) ||
        Math.abs(box.y - tail.y - top - table.outerTop) > .12 ||
        Math.abs(line.height - boxHeight) > .12 || Math.abs(line.textHeight - boxHeight) > .12 ||
        Math.abs(tail.h - boxHeight - gap - top - bottom) > .12) return;
    return prefix.before;
  } catch { /* Missing optional native geometry leaves the cell unchanged. */ }
}

/** Verify the authored before-space across real line-level cell breaks. Stored
 * offsets restart at the top of the next fragment, including mid-paragraph. */
export function sourceSplitCellBefore(doc: HwpDocument, address: [number, number, number, number],
  painted: Map<string, HwpPaintedLine[]>, boxes: readonly { pageIndex: number; y: number; h: number }[],
  paddingTop: number, paddingBottom: number, usePaintedPositions = false): { before: number[]; fillsFragments: boolean } | undefined {
  const [section, paragraph, controlIndex, cellIndex] = address;
  const count = doc.getCellParagraphCount(...address);
  if (count < 2 || boxes.length < 2 || usePaintedPositions && boxes.length > 3) return;
  const before: number[] = [];
  const byPage = new Map(boxes.map(box => [box.pageIndex, box]));
  const tolerance = sourceCellBoundsTolerance;
  let previous: { page: number; position: number; advance: number; font: number; room: number } | undefined;
  let verifiedCuts = 0;
  for (let index = 0; index < count; index++) {
    const properties = JSON.parse(doc.getCellParaPropertiesAt(...address, index));
    const gap = Number(properties.spacingBefore) / 2;
    if (!(Number.isFinite(gap) && gap > 0) || Number(properties.spacingAfter) !== 0 ||
      properties.headType !== 'None' || properties.lineSpacingType !== 'Percent') return;
    const observed = painted.get(`${section}:${paragraph}:${controlIndex}:${cellIndex}:${index}`);
    const path = JSON.stringify([{ controlIndex, cellIndex, cellParaIndex: index }]);
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path)) as
      { stored: boolean; controlCount: number; lines: StoredCellLine[] };
    if (!stored || controlCount || !lines?.length) return;
    const literal = usePaintedPositions ? Array.from(doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff)) : [];
    // A terminal paragraph containing only ordinary spaces can have a stored
    // line with no painted ink. Verify its position against the preceding
    // measured line on the same sheet and leave the literal spaces untouched.
    if (!observed?.length && !(index === count - 1 && previous?.page === boxes.at(-1)!.pageIndex &&
      lines.length === 1 && !/[^ ]/.test(doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff)))) return;
    if (observed?.length && observed.length !== lines.length || usePaintedPositions && !observed?.length) return;
    const char = JSON.parse(doc.getCellCharPropertiesAt(...address, index, 0));
    const font = Number(char.fontSize) / 75;
    const advance = physicalLineHeight(properties, char);
    if (!(font > 0 && advance >= font)) return;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i], visual = observed?.[i];
      const page = visual?.page ?? previous!.page, box = byPage.get(page);
      if (!box || (i ? !(line.textStart > lines[i - 1].textStart) : line.textStart !== 0)) return;
      const movedToNextPage = previous?.page !== page;
      if (previous && movedToNextPage) {
        if (page !== previous.page + 1) return;
        const old = previous;
        // A new paragraph carries its authored gap, whereas a continued
        // line starts at the top of the continuation.
        if (i ? old.room <= old.advance - old.font + tolerance : old.room + tolerance < gap + font) verifiedCuts++;
      }
      const top = (movedToNextPage || !previous ? 0 : previous.position + previous.advance) + (i ? 0 : gap);
      const room = box.h - paddingTop - top - font - paddingBottom;
      // Some cells restart stored offsets independently of physical sheets.
      // Use actual text starts only when the caller has corroborated the physical fragments.
      const start = visual?.charStart ?? -1;
      const verifiedOffset = usePaintedPositions ? start >= line.textStart &&
        start < (lines[i + 1]?.textStart ?? literal.length) &&
        literal.slice(line.textStart, start).every(char => char === ' ' || char === '\u00a0')
        : Math.abs(line.verticalPosition - top) <= tolerance;
      if (!(verifiedOffset &&
        Math.abs(line.height - font) <= tolerance && Math.abs(line.textHeight - font) <= tolerance &&
        (!visual || Math.abs(visual.y - box.y - paddingTop - top) <= tolerance && Math.abs(visual.h - font) <= tolerance) &&
        Math.abs(line.height + line.spacing - advance) <= tolerance && room >= -tolerance)) return;
      previous = { page, position: top, advance, font, room };
    }
    before.push(gap);
  }
  const fillsFragments = verifiedCuts === boxes.length - 1 && !!previous &&
    previous.page === boxes.at(-1)!.pageIndex && Math.abs(previous.room) <= tolerance;
  if (usePaintedPositions && verifiedCuts !== boxes.length - 1) return;
  return { before, fillsFragments };
}

/** Confirm authored paragraph spacing in a repeated row where both cells
 * keep contributing across page boundaries. A peer may finish earlier. */
export function sourcePaintedCellBefore(doc: HwpDocument, address: [number, number, number, number],
  painted: ReadonlyMap<string, HwpPaintedLine[]>, pitches: ReadonlyMap<string, number>,
  boxes: readonly { pageIndex: number; y: number; h: number }[], paddingTop: number, paddingBottom: number,
  checkAuthoredBreaks = false, peerFillsFirst = false) {
  const [section, paragraph, controlIndex, cellIndex] = address;
  const byPage = new Map(boxes.map(box => [box.pageIndex, box]));
  const tolerance = sourceCellBoundsTolerance;
  const before: number[] = [];
  let previous: { page: number; y: number; height: number; pitch: number; remaining: number } | undefined;
  const cuts: { paragraph: number; continuedLine: boolean }[] = [];
  let terminalBlank = false;
  for (let index = 0, count = doc.getCellParagraphCount(...address); index < count; index++) {
    const properties = JSON.parse(doc.getCellParaPropertiesAt(...address, index));
    const gap = Number(properties.spacingBefore) / 2;
    if (!(gap > 0 && Number.isFinite(gap)) || properties.spacingAfter !== 0 ||
        properties.headType !== 'None' || properties.lineSpacingType !== 'Percent') return;
    const key = `${section}:${paragraph}:${controlIndex}:${cellIndex}:${index}`;
    const path = JSON.stringify([{ controlIndex, cellIndex, cellParaIndex: index }]);
    const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path)) as
      { stored: boolean; controlCount: number; lines?: StoredCellLine[] };
    if (!stored || controlCount || !lines?.length) return;
    const text = doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
    if (/[\u0000-\u0009\u000b-\u001f\ufffc]/u.test(text)) return;
    const literal = Array.from(text);
    if (text.includes('\n') && (!checkAuthoredBreaks || literal.some((char, offset) =>
      char === '\n' && !lines.some(line => line.textStart === offset + 1)))) return;
    terminalBlank = !literal.length;
    let visible = painted.get(key);
    if (!visible?.length && !literal.length && lines.length === 1 && index === count - 1) {
      // Reader reports the terminal blank line box with the other paragraph
      // positions. It has no glyphs and still supplies its authored height.
      const page = pitches.get(`page:${key}`), y = pitches.get(`top:${key}`), h = pitches.get(`height:${key}`);
      if (page !== undefined && previous?.page === page && y !== undefined && h !== undefined)
        visible = [{ page, y, h, charStart: 0 }];
    }
    if (visible?.length !== lines.length) return;
    for (const [i, line] of lines.entries()) {
      const { page, y, h, charStart } = visible[i], box = byPage.get(page);
      const pitch = line.height + line.spacing;
      const next = lines[i + 1]?.textStart ?? literal.length;
      if (!box || !(h > 0 && Number.isFinite(y) && Number.isFinite(pitch) && pitch >= line.height) ||
          !Number.isInteger(line.textStart) ||
          Math.abs(h - line.height) > tolerance || Math.abs(line.textHeight - line.height) > tolerance ||
          (i ? !(line.textStart > lines[i - 1].textStart) : line.textStart !== 0) ||
          (literal.length ? charStart === undefined || charStart < line.textStart || charStart >= next ||
            !literal.slice(line.textStart, charStart).every(ch => ch === ' ' || ch === '\u00a0') : charStart !== 0)) return;
      const continuation = !previous || previous.page !== page;
      if (previous && continuation) {
        // A second cell can have used the first fragment's remaining
        // height even though this cell has room for another paragraph.
        const sharedBoundary = peerFillsFirst && !cuts.length && i === 0;
        if (page !== previous.page + 1 || !sharedBoundary &&
            previous.remaining >= previous.pitch - previous.height + (i ? 0 : gap) + h + tolerance) return;
        cuts.push({ paragraph: index, continuedLine: i > 0 });
      }
      const origin = continuation ? box.y + paddingTop : previous!.y + previous!.pitch;
      const remaining = box.y + box.h - paddingBottom - y - h;
      if (Math.abs(y - origin - (i ? 0 : gap)) > tolerance || remaining < -tolerance) return;
      previous = { page, y, height: h, pitch, remaining };
    }
    before.push(gap);
  }
  if (!before.length || !previous || previous.page - boxes[0].pageIndex !== cuts.length) return;
  return { before, cuts, terminalBlank, fillsRow: cuts.length === boxes.length - 1 && Math.abs(previous.remaining) <= tolerance };
}

/** Whole paragraphs can continue at a boundary filled by both cells or
 * by alternating cells on successive fragments. */
export function sourceWholeParagraphPeerRow(
  profiles: readonly ReturnType<typeof sourcePaintedCellBefore>[],
  filled: readonly (readonly boolean[])[],
) {
  const [left, right] = profiles;
  if (!left || !right || left.terminalBlank || right.terminalBlank || left.fillsRow === right.fillsRow ||
      [...left.cuts, ...right.cuts].some(cut => cut.continuedLine)) return false;
  const main = left.fillsRow ? 0 : 1, peer = 1 - main;
  const primary = left.fillsRow ? left : right, shorter = left.fillsRow ? right : left;
  if (shorter.cuts.length !== 1 || shorter.cuts[0].paragraph !== 1) return false;
  if (filled.length === 2) {
    return primary.cuts.length === 1 && primary.cuts[0].paragraph === 1 && primary.before.length >= 4 &&
      shorter.before.length === 2 && filled[0][0] && filled[0][1] && filled[1][main] && !filled[1][peer];
  }
  if (filled.length !== 3 || primary.cuts.length !== 2) return false;
  const [first, second] = primary.cuts.map(cut => cut.paragraph);
  return Math.min(first, second - first, primary.before.length - second) >= 2 && shorter.before.length >= 3 &&
    filled[0][main] && !filled[0][peer] && filled[1][peer] && !filled[1][main] && filled[2][main] && !filled[2][peer];
}

/** A second story can carry a complete inline table on the middle sheet of
 * a row whose line positions are independently proven by its text cell. */
export function sourceInlinePeerInMiddleFragment(doc: HwpDocument, address: [number, number, number, number],
  middle: { pageIndex: number; x: number; y: number; w: number; h: number }) {
  const [section, paragraph, controlIndex, cellIndex] = address;
  const count = doc.getCellParagraphCount(...address);
  if (count < 3 || count > 32) return false;
  let found = false;
  for (let index = 0; index < count; index++) {
    const path = [{ controlIndex, cellIndex, cellParaIndex: index }];
    const story = JSON.stringify(path);
    const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, story));
    if (!metrics.stored || !metrics.lines?.length) return false;
    if (!metrics.controlCount) continue;
    const table = metrics.tables?.[0];
    if (found || !index || index === count - 1 || metrics.controlCount !== 1 || metrics.tables?.length !== 1 ||
        metrics.lines.length !== 1 || !table?.inline ||
        doc.getTextInCellByPath(section, paragraph, story, 0, 0xffffffff).trim()) return false;
    const child = JSON.stringify([...path, { controlIndex: table.index, cellIndex: 0, cellParaIndex: 0 }]);
    const cells: Array<{ cellIdx: number; pageIndex: number; x: number; y: number; w: number; h: number }> =
      JSON.parse(doc.getTableCellBboxesByPath(section, paragraph, child));
    const { cellCount } = JSON.parse(doc.getTableDimensionsByPath(section, paragraph, child));
    const tolerance = sourceCellBoundsTolerance;
    if (!cells.length || cells.length !== cellCount || new Set(cells.map(cell => cell.cellIdx)).size !== cellCount ||
        !cells.every(cell => cell.pageIndex === middle.pageIndex && cell.w > 0 && cell.h > 0 &&
          cell.x >= middle.x - tolerance && cell.y >= middle.y - tolerance &&
          cell.x + cell.w <= middle.x + middle.w + tolerance && cell.y + cell.h <= middle.y + middle.h + tolerance)) return false;
    found = true;
  }
  return found;
}

export function sourceColumnCellWrap(doc: HwpDocument, section: number, paragraph: number, path: CellPath,
  observed: HwpPaintedLine[] | undefined, box: { pageIndex: number; y: number; h: number; w: number },
  stories: number): HwpSourceWrap | undefined {
  if (!observed || observed.length < 2 || observed.length > 80 || observed.length > 2 && !stories ||
      !observed.every(line => line.page === box.pageIndex &&
      line.y >= box.y - .1 && line.y + line.h <= box.y + box.h + .1)) return;
  const address = JSON.stringify(path);
  const { stored, controlCount, lines } = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address)) as
    { stored: boolean; controlCount: number; lines?: (StoredCellLine & { width: number })[] };
  if (!stored || controlCount || lines?.length !== observed.length || lines[0].textStart !== 0 ||
      !lines.every((line, i) =>
        line.textStart === observed[i].charStart && line.width === lines[0].width && line.width > 0 && line.width <= box.w)) return;
  const text = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
  if (!text || text.length > (lines.length === 2 ? 256 : 1024) ||
      lines.slice(1).some((line, i) => !Number.isInteger(line.textStart) || line.textStart <= lines[i].textStart || line.textStart >= text.length) ||
      /[\u0000-\u001f\ufffc]/u.test(text)) return;
  const ends = [...lines.slice(1).map(line => normalizeHwpWrapText(text.slice(0, line.textStart)).length),
    normalizeHwpWrapText(text).length];
  if (ends.some((end, i) => end <= (ends[i - 1] ?? 0))) return;
  const proof = { fingerprint: hwpWrapFingerprint(text), ends, cellWidth: { outer: box.w, text: lines[0].width }, columnCellHeight: box.h };
  if (lines.length === 2) return proof;
  const pitch = lines[0].height + lines[0].spacing;
  if (!(pitch > 0) || !lines.every((line, i) =>
      Math.abs(line.height + line.spacing - pitch) <= .08 && Math.abs(line.height - observed[i].h) <= .1 &&
      (!i || Math.abs(line.verticalPosition - lines[i - 1].verticalPosition - pitch) <= .08 &&
        Math.abs(observed[i].y - observed[i - 1].y - pitch) <= .15))) return;
  return { ...proof, columnCellPitch: pitch, ...(stories > 1 && { columnCellStories: stories }) };
}

/** Shared cell line-box semantics, independent of the table nesting depth. */
export function cellParagraphCss(properties: Record<string, any>, char: Record<string, any>, options: {
  pitch?: number; inkHeight?: number; cap?: number; before?: number; final: boolean; lines?: number; simple: boolean;
  linePolicyFont?: number;
  authoredTerminal?: HwpAuthoredCellTerminal;
}): string {
  let pitch = options.pitch;
  let policy = '';
  if (options.final && options.lines !== undefined && options.simple) {
    const inkHeight = options.inkHeight ?? Number(char.fontSize) / 75;
    const requested = pitch ?? physicalLineHeight(properties, char);
    if (inkHeight > 0 && requested > inkHeight) {
      if (options.lines <= 1) pitch = inkHeight;
      const font = Number(char.fontSize) / 75;
      if (['Percent', 'SpaceOnly', 'Minimum', 'Fixed'].includes(properties.lineSpacingType) && Math.abs(inkHeight - font) < .5) {
        pitch = options.lines <= 1 ? inkHeight : requested;
        policy = `--hwp-cell-line:${encodeHwpCellLine({ rule: properties.lineSpacingType, font: options.linePolicyFont ?? font, advance: requested, end: inkHeight, base: options.lines <= 1 ? inkHeight : requested,
          ...(options.lines === 1 && options.authoredTerminal ? { authoredTerminal: options.authoredTerminal } : {}) })};`;
      }
    }
  }
  return paragraphCss(properties, char, options.cap, false) + (options.before ? `margin-top:${options.before}px;` : '') +
    (pitch ? `line-height:${pitch}px;` : '') + policy;
}
