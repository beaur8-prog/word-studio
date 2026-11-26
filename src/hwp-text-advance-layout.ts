import { hwpNestedCellFaceAttribute } from './hwp-nested-table-layout';
import { characterScaleProperty, parseCharacterScale } from './character-scale';
import { hwpBodyFallbackSelector, hwpCellFaceFallbackSelector, hwpCellWidthFallbackSelector, hwpCompactCellFace, hwpCompactHeaderFace, hwpShortCellLabel, hwpSingleBodyFace, hwpSourceSpacesAttribute, hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap, type HwpSourceWrap } from './hwp-text-advances';
import { parseHwpBodyLine } from './hwp-body-line';
import { measuredHwpLineInk } from './hwp-line-ink';
import { paragraphLineOffsets } from './paragraph-fragments';

const sourceFaces: Record<string, readonly string[]> = {
  'malgun gothic': ['Malgun Gothic', 'MalgunGothic', '맑은 고딕'],
  '한양중고딕': ['한양중고딕'],
  '휴먼명조': ['휴먼명조'],
  '돋움': ['돋움', 'Dotum'],
};
const familyName = (value: string) => {
  const family = value.split(',')[0].trim().replace(/^["']|["']$/g, '').toLowerCase();
  return family === '맑은 고딕' ? 'malgun gothic' : family === 'dotum' ? '돋움' : family;
};
const knownFaces = new Map<string, Promise<boolean>>();
const paragraphs = `p[${hwpSourceSpacesAttribute}="true"]`;
export const hwpSourceAdvanceRun = 'data-hwp-source-advance-run';
const unsupported = 'li,.document-furniture,.document-textbox,.table-wrap-frame,[data-document-note],[data-column-count]:not([data-column-count="1"])';

function fontUnavailable(family: string, aliases: readonly string[] = sourceFaces[family]) {
  if ([...document.fonts].some(font => font.status === 'loaded' && familyName(font.family) === family))
    return Promise.resolve(false);
  const present = knownFaces.get(family);
  if (present) return present;
  const probe = new FontFace('Studio HWP source face check', aliases.map(name =>
    `local(${JSON.stringify(name)})`).join(','));
  const missing = probe.load().then(() => false, () => true);
  knownFaces.set(family, missing);
  return missing;
}

type InlineAdjustment = { text: Text; from: number; to: number; style: string; hangul: boolean };
type ParagraphAdjustment = { paragraph: HTMLElement; runs: InlineAdjustment[] };

export function hwpSourceAdvanceParagraph(paragraph: HTMLElement, allowPrintFragment = false, allowPrintTerminal = false) {
  return paragraph.matches(paragraphs) && !paragraph.querySelector('span[style*="--word-character-scale"]') &&
    plainHwpSourceParagraph(paragraph, allowPrintFragment, allowPrintTerminal);
}

export function plainHwpSourceParagraph(paragraph: HTMLElement, allowPrintFragment = false, allowPrintTerminal = false) {
  return paragraph.tagName === 'P' && !paragraph.closest(unsupported) &&
    (allowPrintFragment || !paragraph.hasAttribute('data-print-paragraph-fragment')) &&
    !paragraph.matches('[data-hwp-print-source-flow],[data-hwp-cell-continuation]') &&
    (allowPrintTerminal || !paragraph.hasAttribute('data-hwp-cell-terminal')) &&
    !paragraph.querySelector('img,svg,br:not(.ProseMirror-trailingBreak),sup,sub,[data-flow-widget],[data-note-reference],[data-document-field],[data-docx-tab],[data-word-overlap]' +
      (allowPrintTerminal ? '' : ',[data-hwp-cell-prefix]'));
}

function candidateText(root: Iterable<HTMLElement>) {
  const selected: { paragraph: HTMLElement; text: Text; css: CSSStyleDeclaration; family: string }[] = [];
  for (const paragraph of root) {
    if (!hwpSourceAdvanceParagraph(paragraph)) continue;
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      if (!/[가-힣 ]/u.test(text.data)) continue;
      const css = getComputedStyle(text.parentElement!);
      const family = familyName(css.fontFamily);
      if (Object.hasOwn(sourceFaces, family) && css.direction === 'ltr' && css.writingMode === 'horizontal-tb')
        selected.push({ paragraph, text, css, family });
    }
  }
  return selected;
}

function createAdvanceMeasurements() {
  const sample = document.createElement('span');
  sample.style.cssText = 'position:fixed;left:-100000px;visibility:hidden;white-space:pre;letter-spacing:0;word-spacing:0';
  const measurements = new Map<string, { hangul: number; space: number }>();
  function collect(selected: ReturnType<typeof candidateText>, availability: ReadonlyMap<string, boolean>) {
    if (!selected.length) return [];
    const programs = new Map<HTMLElement, InlineAdjustment[]>();
    try {
      for (const { paragraph, text, css, family } of selected) {
        if (!availability.get(family)) continue;
        const size = parseFloat(css.fontSize);
        if (!(size > 0) || !['normal', '0px'].includes(css.wordSpacing)) continue;
        const props = [css.fontFamily, css.fontSize, css.fontWeight, css.fontStyle,
          text.parentElement?.closest<HTMLElement>('[lang]')?.lang || ''];
        const key = JSON.stringify(props);
        let measured = measurements.get(key);
        if (!measured) {
          if (!sample.isConnected) document.body.append(sample);
          [sample.style.fontFamily, sample.style.fontSize, sample.style.fontWeight, sample.style.fontStyle, sample.lang] = props;
          sample.textContent = '한글가나다라사아';
          const hangul = sample.getBoundingClientRect().width / 8;
          sample.textContent = '        ';
          measured = { hangul, space: sample.getBoundingClientRect().width / 8 };
          measurements.set(key, measured);
        }
        const runs = programs.get(paragraph) ?? [];
        for (const match of text.data.matchAll(/[가-힣]+| +/gu)) {
          const hangul = match[0][0] !== ' ', width = hangul ? measured.hangul : measured.space;
          const delta = size * (hangul ? 1 : .5) - width;
          if (!(width > 0) || Math.abs(delta) < 1 / 64) continue;
          const style = hangul ? `letter-spacing:${(parseFloat(css.letterSpacing) || 0) + delta}px` : `word-spacing:${delta}px`;
          runs.push({ text, from: match.index!, to: match.index! + match[0].length, style, hangul });
        }
        if (runs.length) programs.set(paragraph, runs);
      }
    } finally { sample.remove(); }
    return [...programs].map(([paragraph, runs]) => ({ paragraph, runs }));
  }
  return collect;
}

async function knownMissingFaces(families: Iterable<string>) {
  return new Map(await Promise.all([...new Set(families)].map(async family => [family, await fontUnavailable(family)] as const)));
}

async function candidates(root: HTMLElement): Promise<ParagraphAdjustment[]> {
  const selected = candidateText(root.querySelectorAll<HTMLElement>(paragraphs));
  if (!selected.length) return [];
  return createAdvanceMeasurements()(selected, await knownMissingFaces(selected.map(run => run.family)));
}

function replaceRuns(runs: InlineAdjustment[]) {
  const byText = new Map<Text, InlineAdjustment[]>();
  for (const run of runs) {
    const group = byText.get(run.text);
    if (group) group.push(run);
    else byText.set(run.text, [run]);
  }
  const undo: (() => void)[] = [];
  for (const [text, program] of byText) {
    const result = document.createDocumentFragment();
    let start = 0;
    for (const { from, to, style } of program) {
      if (from > start) result.append(document.createTextNode(text.data.slice(start, from)));
      const span = document.createElement('span');
      span.style.cssText = style;
      span.textContent = text.data.slice(from, to);
      result.append(span);
      start = to;
    }
    if (start < text.length) result.append(document.createTextNode(text.data.slice(start)));
    const content = [...result.childNodes];
    text.replaceWith(result);
    undo.push(() => { content[0].replaceWith(text); content.slice(1).forEach(node => node.remove()); });
  }
  return () => undo.splice(0).forEach(restore => restore());
}

function ordinarySingleCell(paragraph: HTMLElement) {
  const cell = paragraph.parentElement;
  if (!(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1) return false;
  const table = cell.closest('table');
  return table?.rows.length === 1 && table.rows[0].cells.length === 1 && !table.querySelector('table') &&
    !table.parentElement?.closest('table') && !table.matches('[data-word-table-width-auto="true"]');
}

type SourcePaper = { index: number; box: DOMRect; bottom: number; scale: number; liveEditor?: boolean; topMargin?: number };

function bodyContinuationOffset(paragraph: HTMLElement, paper: SourcePaper, paragraphTop: number) {
  const { box, scale, topMargin } = paper;
  if (topMargin === undefined || !Number.isFinite(topMargin)) return;
  // The gutter of an earlier continued body paragraph ends at this sheet's
  // text origin. Use its actual edge when the paper and live flow round their
  // accumulated coordinates independently, within the normal fit tolerance.
  for (let sibling = paragraph.previousElementSibling; sibling; sibling = sibling.previousElementSibling) {
    if (sibling.getBoundingClientRect().bottom < box.top) return;
    if (!sibling.matches('p') || sibling.getBoundingClientRect().top >= box.top) continue;
    for (const widget of sibling.querySelectorAll('.pagination-cell-gap.ProseMirror-widget[aria-hidden="true"]')) {
      const gutter = widget.getBoundingClientRect();
      const shift = (gutter.bottom - box.top) / scale - topMargin;
      if (gutter.top < box.top && gutter.bottom <= paragraphTop && Math.abs(shift) <= .5) return shift;
    }
    return;
  }
}

/** A measured paragraph can share a top-level native cell with a child
 * table. Validate the paragraph and cell widths before growing its flow. */
function sourceContainerText(paragraph: HTMLElement) {
  if (!hwpSourceAdvanceParagraph(paragraph)) return false;
  const cell = paragraph.parentElement;
  if (!(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' ||
      cell.rowSpan !== 1 || cell.colSpan !== 1) return false;
  const row = cell.parentElement as HTMLTableRowElement, table = cell.closest('table');
  if (!table || table.parentElement?.closest('table') || table.rows.length !== 1 || row.cells.length !== 1 ||
      row.dataset.rowSplit !== 'allow' || row.dataset.wordRowRule === 'exact' ||
      table.matches('[data-word-table-width-auto="true"]') || !cell.querySelector(':scope > table')) return false;
  const source = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
  if (!source?.cellWidth || hwpWrapFingerprint(paragraph.textContent || '') !== source.fingerprint ||
      !(parseFloat(getComputedStyle(paragraph).textIndent) < 0)) return false;
  const css = getComputedStyle(cell);
  const tolerance = (parseFloat(css.borderLeftWidth) || 0) + (parseFloat(css.borderRightWidth) || 0) + .05 + 1 / 15;
  return Math.abs(cell.getBoundingClientRect().width - source.cellWidth.outer) <= tolerance &&
    Math.abs(paragraph.getBoundingClientRect().width - source.cellWidth.text) <= tolerance;
}

export function sourceFlowAllowance(paragraph: HTMLElement, paper?: SourcePaper) {
  if (!hwpSourceAdvanceParagraph(paragraph)) return;
  const metadata = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
  const text = paragraph.textContent ?? '';
  if (!metadata || hwpWrapFingerprint(text) !== metadata.fingerprint) return;
  const indent = parseFloat(getComputedStyle(paragraph).textIndent);
  if (!(indent < 0)) return;
  if (ordinarySingleCell(paragraph)) return Infinity;
  if (!paper || !metadata.origin || metadata.origin.page !== paper.index ||
      !paragraph.parentElement?.matches('.print-content,.section-editable,.tiptap')) return;
  const precision = .05 + 1 / 64, box = paragraph.getBoundingClientRect();
  const { x, y } = metadata.origin;
  const nativeLine = paper.liveEditor && parseHwpBodyLine(paragraph.dataset.hwpBodyLine);
  // The reader stores tenths of a pixel. Saved Word spacing adds one twip,
  // and layout rounds paragraph and paper positions independently.
  const verticalTolerance = nativeLine && nativeLine.align === 'top' ? precision + 1 / 15 + 1 / 64 : precision;
  if (Math.abs((box.left - paper.box.left) / paper.scale + indent - x) > precision) return;
  const drift = (box.top - paper.box.top) / paper.scale - y;
  if (Math.abs(drift) <= verticalTolerance) return (paper.bottom - box.top) / paper.scale;
  if (!nativeLine || nativeLine.align !== 'top' || Math.abs(drift) > .5 + verticalTolerance) return;
  const offset = bodyContinuationOffset(paragraph, paper, box.top);
  if (offset !== undefined && Math.abs(drift - offset) <= verticalTolerance)
    return (paper.bottom - box.top) / paper.scale;
}

function sourceFlowCandidate(program: ParagraphAdjustment, allowance: number, contains: (paragraph: HTMLElement) => boolean,
  exactEnds?: 'body' | 'container') {
  const { paragraph, runs } = program;
  const metadata = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
  const box = paragraph.getBoundingClientRect();
  if (!metadata || !runs.length || paragraphLineOffsets(paragraph).length + 2 !== metadata.ends.length) return;
  const restore = replaceRuns(runs);
  const after = paragraph.getBoundingClientRect();
  const sourceLinesFit = exactEnds === 'container' || exactEnds === 'body' && metadata.origin
    ? matchesStoredWrap(paragraph, paragraph.textContent || '', metadata.ends)
    : paragraphLineOffsets(paragraph).length + 1 === metadata.ends.length;
  if (after.width !== box.width || after.height <= box.height || after.height > allowance ||
      !contains(paragraph) || !sourceLinesFit) {
    restore();
    return;
  }
  return { runs, restore, baseHeight: box.height };
}

function restoreSingleBodyLine(paragraph: HTMLElement, paper: SourcePaper, native: HwpSourceWrap, range: Range) {
  const before = paragraph.getBoundingClientRect(), css = getComputedStyle(paragraph);
  const pitch = parseFloat(css.lineHeight), size = parseFloat(css.fontSize);
  const precision = .05 + 1 / 15 + 2 / 64;
  if (!native.bodyWidth || !paragraph.hasAttribute('data-exact-line-height') || css.textAlign !== 'justify' ||
      css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' || css.transform !== 'none' ||
      !(pitch >= size && size > 0 && Math.abs(before.height - 2 * pitch) <= precision) ||
      Math.abs(before.width - parseFloat(css.textIndent) - native.bodyWidth) > precision ||
      !sourceFlowAllowance(paragraph, paper) || paragraphLineOffsets(paragraph).length !== 1 ||
      paragraph.querySelector(':not(span)')) return;
  const styled = [paragraph, ...paragraph.querySelectorAll<HTMLElement>('span')].map(node =>
    ({ node, family: node.style.fontFamily, style: getComputedStyle(node) }));
  if (styled.some(({ family, style }) => familyName(family || style.fontFamily) !== familyName(hwpSingleBodyFace) ||
      (family || style.fontFamily).includes(',') || parseFloat(style.fontSize) !== size ||
      style.fontStyle !== 'normal' || style.fontWeight !== '400' || style.textDecorationLine !== 'none')) return;

  for (const { node, family } of styled) node.style.fontFamily = `${family || JSON.stringify(hwpSingleBodyFace)}, "Noto Sans KR"`;
  const restore = () => {
    for (const { node, family } of styled) node.style.fontFamily = family;
    delete paragraph.dataset.hwpPrintSourceFlow;
  };
  const after = paragraph.getBoundingClientRect(), body = paragraph.parentElement!.getBoundingClientRect();
  range.selectNodeContents(paragraph);
  const boxes = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
  // break-spaces exposes leading preserved spaces as a separate range.
  // Measure the visible line without moving or removing those spaces.
  const ink = measuredHwpLineInk(paragraph, 1, size, true);
  if (after.width !== before.width || after.height >= before.height ||
      Math.abs(after.height - pitch) > precision || after.top !== before.top ||
      paragraphLineOffsets(paragraph).length || !boxes.length || !ink ||
      boxes.some(rect => rect.left < body.left - precision || rect.right > body.right + precision) ||
      ink.top < after.top - precision || ink.bottom > Math.min(after.bottom, paper.bottom) + precision) {
    restore();
    return;
  }
  paragraph.dataset.hwpPrintSourceFlow = 'true';
  return restore;
}

/** Try a naturally missing line before flow is split. Body text must still
 * start at its authored paper coordinates. A one-cell ordinary table keeps
 * its own width as its row grows and follows the usual table pagination. */
export async function prepareHwpSourceFlow(root: HTMLElement) {
  if (!root.querySelector(`p[${hwpSourceWrapAttribute}],p[${hwpNestedCellFaceAttribute}]`)) return;
  const bodyFallback = root.querySelector(`${hwpBodyFallbackSelector},p[${hwpNestedCellFaceAttribute}]`);
  const [availability, missingSingleBodyFace, missingCompactCellFace, missingCompactHeaderFace] = await Promise.all([
    knownMissingFaces(Object.keys(sourceFaces)),
    bodyFallback ? fontUnavailable(familyName(hwpSingleBodyFace),
      [hwpSingleBodyFace, 'KoPubBatang Light', 'KoPubBatang-Light']) : false,
    root.querySelector(`${hwpCellFaceFallbackSelector},${hwpCellWidthFallbackSelector}`) ? fontUnavailable(familyName(hwpCompactCellFace),
      [hwpCompactCellFace, 'KoPubDotum Light', 'KoPubDotum-Light']) : false,
    root.querySelector(`${hwpCellFaceFallbackSelector}[${hwpSourceWrapAttribute}*='"header":true']`)
      ? fontUnavailable(familyName(hwpCompactHeaderFace), [hwpCompactHeaderFace, 'KoPubDotum Bold', 'KoPubDotum-Bold']) : false,
  ]);
  const measurements = createAdvanceMeasurements();
  const range = document.createRange();
  function restoreMissingWrap(paragraph: HTMLElement) {
    if (!paragraph.hasAttribute(hwpSourceWrapAttribute)) return;
    const sheet = paragraph.closest<HTMLElement>('.print-page'), body = paragraph.closest('.print-content');
    const paper = sheet && body && { index: Number(sheet.dataset.page), box: sheet.getBoundingClientRect(),
      bottom: body.getBoundingClientRect().bottom + .5, scale: 1 };
    const native = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
    if (native?.singleBodyFallback)
      return missingSingleBodyFace && paper ? restoreSingleBodyLine(paragraph, paper, native, range) : undefined;
    const containerText = sourceContainerText(paragraph);
    const allowance = containerText ? Infinity : sourceFlowAllowance(paragraph, paper || undefined);
    if (!allowance || !(allowance > 0)) return;
    const program = measurements(candidateText([paragraph]), availability)[0];
    if (!program) return;
    const candidate = sourceFlowCandidate(program, allowance, p => {
      const bounds = p.parentElement!.getBoundingClientRect();
      range.selectNodeContents(p);
      return [...range.getClientRects()].every(box => !(box.width > 0 && box.height > 0) ||
        box.left >= bounds.left - .5 && box.right <= bounds.right + .5);
    }, containerText ? 'container' : undefined);
    if (!candidate) return;
    paragraph.dataset.hwpPrintSourceFlow = 'true';
    return () => { delete paragraph.dataset.hwpPrintSourceFlow; candidate.restore(); };
  }

  // The final print pass cannot compare a shortened cell paragraph against
  // the complete source fingerprint. Check its natural breaks before cutting.
  function preserveSplitCellWrap(paragraph: HTMLElement, verifiedOverhangContinuation = false) {
    const cell = paragraph.parentElement;
    if (!(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' ||
        cell.rowSpan !== 1 || cell.colSpan !== 1 || cell.closest('table')?.parentElement?.closest('td,th') ||
        !hwpSourceAdvanceParagraph(paragraph)) return;
    const reference = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
    const text = paragraph.textContent || '';
    if (!reference || reference.ends.length < 2 || hwpWrapFingerprint(text) !== reference.fingerprint ||
        matchesStoredWrap(paragraph, text, reference.ends)) return;
    const before = geometry(paragraph);
    const preserve = (restore: () => void, allowWidth: boolean) => {
      const bounds = cell.getBoundingClientRect();
      range.selectNodeContents(paragraph);
      if (geometry(paragraph).some((value, i) => (!allowWidth || i > 0) && value !== before[i]) ||
          !matchesStoredWrap(paragraph, text, reference.ends) ||
          [...range.getClientRects()].some(rect => rect.width > 0 && rect.height > 0 &&
            (rect.left < bounds.left - .5 || rect.right > bounds.right + .5))) {
        restore();
        return;
      }
      paragraph.dataset.hwpPrintSourceFlow = 'true';
      return () => { delete paragraph.dataset.hwpPrintSourceFlow; restore(); };
    };

    const program = measurements(candidateText([paragraph]), availability)[0];
    if (program) for (const mode of advanceModes) {
      const runs = runsInMode(program.runs, mode);
      if (!runs.length) continue;
      const accepted = preserve(replaceRuns(runs), false);
      if (accepted) return accepted;
    }
    const adjustment = sourceCellWidthAdjustment(paragraph);
    if (!adjustment) return;
    const width = parseFloat(getComputedStyle(paragraph).width) + adjustment.delta;
    if (!(width > 0)) return;
    const oldWidth = paragraph.style.width;
    const restoreWidth = () => { paragraph.style.width = oldWidth; };
    paragraph.style.width = `${width}px`;
    const fittedWidth = preserve(restoreWidth, true);
    if (fittedWidth || !verifiedOverhangContinuation || !program?.runs.length) return fittedWidth;
    // The ordinary final pass can use both measurements together. On a
    // continuation moved by the source boundary, retain them before cutting
    // only when the original line ends and the enclosing row still match.
    const restoreRuns = replaceRuns(program.runs);
    paragraph.style.width = `${width}px`;
    return preserve(() => { restoreWidth(); restoreRuns(); }, true);
  }

  return { restoreMissingWrap, preserveSplitCellWrap, missingSingleBodyFace, missingCompactCellFace, missingCompactHeaderFace };
}

function geometry(paragraph: HTMLElement) {
  return [paragraph, paragraph.closest('td,th'), paragraph.closest('tr')].flatMap(element => {
    if (!element) return [];
    const box = element.getBoundingClientRect();
    return [box.width, box.height, box.left, box.top];
  });
}

function inkWithinPage(paragraph: HTMLElement) {
  const paper = paragraph.closest('.print-page')?.getBoundingClientRect();
  const cell = paragraph.closest('td,th')?.getBoundingClientRect();
  if (!paper) return false;
  const range = document.createRange();
  range.selectNodeContents(paragraph);
  return [...range.getClientRects()].every(box => !(box.width > 0 && box.height > 0) ||
    box.left >= (cell ?? paper).left - .5 && box.right <= (cell ?? paper).right + .5 &&
    box.top >= paper.top && box.bottom <= paper.bottom);
}

export function matchesStoredWrap(paragraph: HTMLElement, text: string, ends: readonly number[]) {
  const breaks = paragraphLineOffsets(paragraph);
  return breaks.length + 1 === ends.length && normalizeHwpWrapText(text).length === ends.at(-1) &&
    breaks.every((position, index) => normalizeHwpWrapText(text.slice(0, position)).length === ends[index]);
}

const advanceModes = ['both', 'hangul', 'space'] as const;

function runsInMode(runs: InlineAdjustment[], mode: typeof advanceModes[number]) {
  if (mode === 'both') return runs;
  const selected = runs.filter(run => run.hangul === (mode === 'hangul'));
  return selected.length === runs.length ? [] : selected;
}

export function hasUnchangedSourceWrap(paragraph: HTMLElement) {
  const text = paragraph.textContent || '';
  const reference = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
  return !!reference && reference.ends.length > 1 && hwpWrapFingerprint(text) === reference.fingerprint &&
    matchesStoredWrap(paragraph, text, reference.ends);
}

/** The original break offsets decide which measured native advance can
 * restore ordinary wrapping. No break is added to the paragraph. */
function chooseAdvances(programs: ParagraphAdjustment[], contained: (p: HTMLElement) => boolean) {
  const pending = programs.flatMap(program => {
    const { paragraph } = program;
    const text = paragraph.textContent || '';
    const metadata = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
    const reference = metadata && hwpWrapFingerprint(text) === metadata.fingerprint ? metadata.ends : undefined;
    if (reference && matchesStoredWrap(paragraph, text, reference)) return [];
    const { width, height } = paragraph.getBoundingClientRect();
    return [{ ...program, text, reference, width, height }];
  });
  const accepted = new Map<HTMLElement, { runs: InlineAdjustment[]; restore: () => void }>();
  for (const mode of advanceModes) {
    const applied = pending.flatMap(program => {
      if (accepted.has(program.paragraph) || mode !== 'both' && !program.reference) return [];
      const runs = runsInMode(program.runs, mode);
      if (!runs.length) return [];
      return [{ program, runs, restore: replaceRuns(runs) }];
    });
    const rejected: (() => void)[] = [];
    for (const { program, runs, restore } of applied) {
      const { paragraph, text, reference, width, height } = program;
      const box = paragraph.getBoundingClientRect();
      if (box.width === width && box.height === height && contained(paragraph) &&
          (!reference || matchesStoredWrap(paragraph, text, reference))) accepted.set(paragraph, { runs, restore });
      else rejected.push(restore);
    }
    rejected.forEach(restore => restore());
  }
  return accepted;
}

export type HwpCellWidthAdjustment = { delta: number; left: number; right: number };

export function sourceCellWidthAdjustment(paragraph: HTMLElement, scale = 1, active = 0): HwpCellWidthAdjustment | undefined {
  const cell = paragraph.parentElement;
  if (!(cell instanceof HTMLTableCellElement)) return;
  const reference = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
  if (!reference?.cellWidth || hwpWrapFingerprint(paragraph.textContent || '') !== reference.fingerprint) return;
  const box = paragraph.getBoundingClientRect(), cellBox = cell.getBoundingClientRect();
  const style = getComputedStyle(cell);
  const border = (parseFloat(style.borderLeftWidth) || 0) + (parseFloat(style.borderRightWidth) || 0);
  // Native boxes are rounded to tenths of a pixel. DOCX rounds authored
  // borders and widths to twips. Match the outer box before restoring its
  // text width, which excludes the native padding and paragraph indents.
  const tolerance = border + .05 + 1 / 15;
  const delta = reference.cellWidth.text - (box.width / scale - active);
  if (Math.abs(reference.cellWidth.outer - cellBox.width / scale) > tolerance ||
      delta <= .05 || delta > tolerance) return;
  return { delta, left: (cellBox.left - box.left) / scale, right: (cellBox.right - box.left) / scale };
}

export function shortHwpCellWidthParagraph(paragraph: HTMLElement) {
  const cell = paragraph.parentElement;
  if (!(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' ||
      cell.rowSpan !== 1 || cell.colSpan !== 1 || cell.children.length !== 1 ||
      cell.closest('table')?.parentElement?.closest('td,th') || paragraph.closest(unsupported) ||
      paragraph.querySelector('img,svg,br:not(.ProseMirror-trailingBreak),sup,sub,[data-docx-tab],[data-word-overlap]') ||
      paragraph.style.transform) return false;
  const text = paragraph.textContent ?? '';
  const reference = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
  return !!reference?.cellWidth && reference.ends.length === 1 && hwpShortCellLabel(text) &&
    reference.ends[0] === text.length && reference.fingerprint === hwpWrapFingerprint(text);
}

function tryShortCellWidth(paragraph: HTMLElement, adjustment: HwpCellWidthAdjustment) {
  const text = paragraph.textContent ?? '';
  const box = paragraph.getBoundingClientRect();
  const width = parseFloat(getComputedStyle(paragraph).width) + adjustment.delta;
  if (!hwpShortCellLabel(text) || !(width > 0) || matchesStoredWrap(paragraph, text, [text.length])) return;
  const prior = paragraph.style.width;
  paragraph.style.width = `${width}px`;
  const restore = () => { paragraph.style.width = prior; };
  const span = document.createRange();
  span.selectNodeContents(paragraph);
  const left = box.left + adjustment.left, right = box.left + adjustment.right;
  if (!(paragraph.getBoundingClientRect().height < box.height) || !matchesStoredWrap(paragraph, text, [text.length]) ||
      [...span.getClientRects()].some(rect => rect.width > 0 && rect.height > 0 && (rect.left < left || rect.right > right))) {
    restore();
    return;
  }
  return { restore, baseHeight: box.height };
}

export function shortHwpCellInkFits(paragraph: HTMLElement) {
  const cell = paragraph.parentElement;
  if (!(cell instanceof HTMLTableCellElement) || paragraphLineOffsets(paragraph).length) return false;
  const bounds = cell.getBoundingClientRect(), range = document.createRange();
  range.selectNodeContents(paragraph);
  const ink = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
  return ink.length > 0 && ink.every(rect => rect.left >= bounds.left && rect.right <= bounds.right &&
    rect.top >= bounds.top - .5 && rect.bottom <= bounds.bottom + .5);
}

export function mixedHwpCellRunParagraph(paragraph: HTMLElement, scale = 1) {
  const cell = paragraph.parentElement, table = cell?.closest('table');
  if (!(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' || cell.rowSpan !== 1 || cell.colSpan !== 1 ||
      cell.children.length !== 1 || !table || table.dataset.wordTableWidthAuto === 'true' ||
      table.parentElement?.closest('td,th') || paragraph.closest(unsupported) ||
      paragraph.querySelector('img,svg,br:not(.ProseMirror-trailingBreak),sup,sub,[data-docx-tab],[data-word-overlap]')) return false;
  const ref = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
  if (!ref?.scaleRun || !ref.cellWidth || ref.fingerprint !== hwpWrapFingerprint(paragraph.textContent || '')) return false;
  const style = getComputedStyle(cell), pstyle = getComputedStyle(paragraph);
  if (pstyle.writingMode !== 'horizontal-tb' || pstyle.direction !== 'ltr' || pstyle.transform !== 'none') return false;
  const tolerance = (parseFloat(style.borderLeftWidth) || 0) + (parseFloat(style.borderRightWidth) || 0) + .16;
  const box = paragraph.getBoundingClientRect();
  const fullWidth = box.width / scale - Math.min(0, parseFloat(pstyle.textIndent) || 0);
  return Math.abs(ref.cellWidth.outer - cell.getBoundingClientRect().width / scale) <= tolerance &&
    Math.abs(ref.cellWidth.text - fullWidth) <= tolerance;
}

function tryMixedHwpCellRun(paragraph: HTMLElement) {
  const reference = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
  const run = reference?.scaleRun, text = paragraph.textContent || '';
  if (!run || reference.fingerprint !== hwpWrapFingerprint(text) || run.start >= text.length ||
      normalizeHwpWrapText(text).length !== reference.ends[0] || !paragraphLineOffsets(paragraph).length) return;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let offset = 0;
  let suffix: Text | undefined;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    let mark = node.parentElement, scale: number | null = null;
    while (mark && mark !== paragraph) {
      scale = parseCharacterScale(mark.style.getPropertyValue(characterScaleProperty));
      if (scale !== null) break;
      mark = mark.parentElement;
    }
    if (offset === run.start && offset + node.length === text.length && scale === run.scale) {
      suffix = node;
    } else if (scale !== null && scale !== 100) return;
    offset += node.length;
  }
  if (!suffix || !suffix.parentElement) return;
  const css = getComputedStyle(suffix.parentElement);
  if (familyName(css.fontFamily) !== familyName(run.font) || Math.abs(parseFloat(css.fontSize) - run.size) > .05) return;
  const prior = paragraph.getBoundingClientRect();
  const span = document.createElement('span');
  span.style.cssText = 'display:inline-block;white-space:pre';
  suffix.replaceWith(span);
  span.append(suffix);
  const restore = () => span.replaceWith(suffix!);
  const natural = parseFloat(getComputedStyle(span).width);
  const factor = run.width / natural;
  if (!(factor >= .5 && factor < 1 && natural - run.width >= 1 / 64)) {
    restore();
    return;
  }
  const style = `display:inline-block;white-space:pre;margin-right:${run.width - natural}px;transform:scaleX(${factor});transform-origin:left center`;
  span.style.cssText = style;
  const box = paragraph.getBoundingClientRect(), range = document.createRange();
  const left = box.left + Math.min(0, parseFloat(getComputedStyle(paragraph).textIndent) || 0);
  const ink = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let content = false;
  while (ink.nextNode()) {
    range.selectNodeContents(ink.currentNode);
    for (const rect of range.getClientRects()) {
      if (!rect.width || !rect.height) continue;
      content = true;
      if (rect.left < left - .5 || rect.right > box.right + .5) {
        restore();
        return;
      }
    }
  }
  if (!content || !(box.height < prior.height) || paragraphLineOffsets(paragraph).length) {
    restore();
    return;
  }
  return { restore, plan: { baseHeight: prior.height, runs: [{ from: run.start, to: text.length, style }] } };
}

export function layoutMixedHwpCellRuns(table: HTMLTableElement) {
  for (const paragraph of table.querySelectorAll<HTMLElement>('p[data-hwp-source-wrap]')) {
    if (!mixedHwpCellRunParagraph(paragraph)) continue;
    const cell = paragraph.parentElement!, width = cell.getBoundingClientRect().width;
    const trial = tryMixedHwpCellRun(paragraph);
    if (trial && Math.abs(cell.getBoundingClientRect().width - width) > 1 / 64) trial.restore();
  }
}

/** Match the measured source content width before testing where the rows fit.
 * Accept the change only when the unchanged text naturally fits in one line. */
export function layoutShortHwpCellWidths(table: HTMLTableElement) {
  for (const p of table.querySelectorAll<HTMLParagraphElement>('p[data-hwp-source-wrap]')) {
    if (!shortHwpCellWidthParagraph(p)) continue;
    const adjustment = sourceCellWidthAdjustment(p);
    if (!adjustment) continue;
    const cell = p.parentElement!;
    const width = cell.getBoundingClientRect().width;
    const trial = tryShortCellWidth(p, adjustment);
    if (!trial) continue;
    if (Math.abs(cell.getBoundingClientRect().width - width) > 1 / 64 || !shortHwpCellInkFits(p))
      trial.restore();
  }
}

type WidthTrial = ParagraphAdjustment & { reference: HwpSourceWrap; text: string; target: number };

function sourceWidthTrial(program: ParagraphAdjustment, adjustment: HwpCellWidthAdjustment): WidthTrial | undefined {
  const { paragraph } = program;
  const reference = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute));
  const text = paragraph.textContent || '';
  if (!reference || matchesStoredWrap(paragraph, text, reference.ends)) return;
  const target = parseFloat(getComputedStyle(paragraph).width) + adjustment.delta;
  if (target > 0) return { ...program, reference, text, target };
}

function chooseSourceWidths<T extends WidthTrial>(trials: T[], fits: (trial: T) => boolean) {
  const selected = new Map<HTMLElement, { trial: T; runs: InlineAdjustment[]; restore: () => void }>();
  // Box geometry and the fallback font can each account for part of an
  // inexact line. Keep an isolated width match before trying both together.
  for (const useAdvances of [false, true]) {
    const active = trials.flatMap(trial => {
      const { paragraph, target } = trial;
      if (selected.has(paragraph) || useAdvances && !trial.runs.length) return [];
      const runs = useAdvances ? trial.runs : [];
      const oldWidth = paragraph.style.width;
      const restoreRuns = replaceRuns(runs);
      paragraph.style.width = `${target}px`;
      return [{ trial, runs, restore: () => { paragraph.style.width = oldWidth; restoreRuns(); } }];
    });
    const accepted: typeof active = [], rejected: typeof active = [];
    for (const item of active) (fits(item.trial) ? accepted : rejected).push(item);
    for (const item of rejected) item.restore();
    if ([...selected.values(), ...accepted].some(item => !fits(item.trial))) {
      for (const item of accepted) item.restore();
    } else {
      for (const item of accepted) selected.set(item.trial.paragraph, item);
    }
  }
  return [...selected.values()];
}

function correctSourceCellWidths(programs: ParagraphAdjustment[]) {
  const attempts = programs.flatMap(program => {
    const adjustment = sourceCellWidthAdjustment(program.paragraph);
    const trial = adjustment && sourceWidthTrial(program, adjustment);
    return trial ? [{ ...trial, box: geometry(program.paragraph) }] : [];
  });
  return chooseSourceWidths(attempts, ({ paragraph, reference, text, box }) =>
    geometry(paragraph).slice(1).every((value, index) => value === box[index + 1]) &&
    inkWithinPage(paragraph) && matchesStoredWrap(paragraph, text, reference.ends));
}

/** Keep all previous physical paragraph and row boundaries. The measured
 * advance affects text inside that space and leaves the source font intact. */
export async function layoutHwpSourceAdvances(root: HTMLElement) {
  const programs = await candidates(root);
  const original = new Map(programs.map(({ paragraph }) => [paragraph, geometry(paragraph)]));
  const accepted = chooseAdvances(programs, inkWithinPage);
  const widths = correctSourceCellWidths(programs);
  const expanded = new Set(widths.map(({ trial }) => trial.paragraph));
  if ([...accepted.keys()].some(paragraph =>
    geometry(paragraph).some((size, index) => (!expanded.has(paragraph) || index > 0) &&
      size !== original.get(paragraph)![index]) || !inkWithinPage(paragraph))) {
    for (const { restore } of widths) restore();
    accepted.forEach(item => item.restore());
  }
}

export type HwpSourceAdvanceRun = { from: number; to: number; style: string };
export type HwpSourceAdvancePlan = { runs: HwpSourceAdvanceRun[]; baseHeight?: number; widthDelta?: number };

export async function planHwpSourceAdvanceDecorations(root: HTMLElement, allowances: ReadonlyMap<HTMLElement, number>,
  widths: ReadonlyMap<HTMLElement, HwpCellWidthAdjustment>, mixedRuns: ReadonlySet<HTMLElement>) {
  const programs = await candidates(root);
  const offsets = new Map<Text, number>();
  for (const { paragraph } of programs) {
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    let offset = 0;
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      offsets.set(node, offset);
      offset += node.length;
    }
  }
  const range = document.createRange();
  const contained = (paragraph: HTMLElement) => {
    const box = paragraph.getBoundingClientRect();
    const hanging = Math.min(0, parseFloat(getComputedStyle(paragraph).textIndent) || 0);
    range.selectNodeContents(paragraph);
    return [...range.getClientRects()].every(rect => !(rect.width > 0 && rect.height > 0) ||
      rect.left >= box.left + hanging - .5 && rect.right <= box.right + .5);
  };
  const flow = new Set<HTMLElement>();
  const result = new Map<HTMLElement, HwpSourceAdvancePlan>();
  const savePlan = (paragraph: HTMLElement, runs: InlineAdjustment[], baseHeight?: number, widthDelta?: number) => {
    result.set(paragraph, { baseHeight, widthDelta, runs: runs.map(({ text, from, to, style }) => {
      const start = offsets.get(text)!;
      return { from: start + from, to: start + to, style };
    }) });
  };
  for (const paragraph of mixedRuns) {
    const trial = tryMixedHwpCellRun(paragraph);
    if (!trial) continue;
    trial.restore();
    result.set(paragraph, trial.plan);
  }
  for (const [paragraph, adjustment] of widths) {
    const reference = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
    if (reference?.ends.length !== 1) continue;
    const trial = tryShortCellWidth(paragraph, adjustment);
    if (!trial) continue;
    trial.restore();
    savePlan(paragraph, [], trial.baseHeight, adjustment.delta);
  }
  for (const program of programs) {
    const allowance = allowances.get(program.paragraph);
    if (!allowance) continue;
    const candidate = sourceFlowCandidate(program, allowance, contained, 'body');
    if (candidate) {
      candidate.restore();
      flow.add(program.paragraph);
      savePlan(program.paragraph, candidate.runs, candidate.baseHeight);
    }
  }
  const accepted = chooseAdvances(programs.filter(({ paragraph }) => !flow.has(paragraph)), contained);
  for (const [paragraph, { runs, restore }] of accepted) {
    restore();
    savePlan(paragraph, runs);
  }
  const trials = programs.flatMap(program => {
    const { paragraph } = program;
    const adjustment = widths.get(paragraph);
    if (!adjustment || result.has(paragraph)) return [];
    const trial = sourceWidthTrial(program, adjustment);
    return trial ? [{ ...trial, adjustment, box: paragraph.getBoundingClientRect() }] : [];
  });
  const ownedWidths = chooseSourceWidths(trials, ({ paragraph, adjustment, text, reference, box }) => {
    range.selectNodeContents(paragraph);
    const left = box.left + adjustment.left - .5, right = box.left + adjustment.right + .5;
    return box.height === paragraph.getBoundingClientRect().height &&
      [...range.getClientRects()].every(rect => !(rect.width > 0 && rect.height > 0) || rect.left >= left && rect.right <= right) &&
      matchesStoredWrap(paragraph, text, reference.ends);
  });
  for (const { trial, runs, restore } of ownedWidths) {
    restore();
    savePlan(trial.paragraph, runs, undefined, trial.adjustment.delta);
  }
  return result;
}
