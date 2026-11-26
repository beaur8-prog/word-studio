import { editorLayoutResume, editorLayoutSuspended } from './editor-layout-task';
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import type { Node as ModelNode } from '@tiptap/pm/model';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { activeHwpBodyLine, encodeHwpBodyLine, hwpBodyTextLine, parseHwpBodyLine, type HwpBodyLine } from './hwp-body-line';
import { cssPixels } from './hwp-cell-line';
import { followsWithHwpColumnBreak, measureHwpColumnEnd } from './hwp-column-end';
import { measureSourceColumnEnd, sourceColumnPaintStyle, sourceColumnTail, type SourceColumnPlacement } from './hwp-column-source-layout';
import { documentScale } from './document-scale';
import { paragraphLineOffsets } from './paragraph-fragments';
import { horizontalInkFits, lineOpportunities, sourceLineBoundaries } from './hwp-source-cell-lines';
import { hwpSingleBodyFace, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } from './hwp-text-advances';
import { matchesStoredWrap } from './hwp-text-advance-layout';
import { measuredHwpInkLines, measuredHwpLineInk, retainedTextBands } from './hwp-line-ink';
import { imageFloating } from './image-floating';

// A view-only page spacer carries line-height:0; it is not an authored
// run override and must not disable the paragraph's native line policy.
const viewFurniture = '.pagination-gap,.pagination-cell-gap,[data-flow-widget]';
const objects = 'br,svg,table,sup,sub,[data-note-reference],[data-page-number],[data-docx-tab],[data-docx-bookmark],[data-word-overlap],[data-symbol-font]';
function eligible(p: HTMLElement, line: HwpBodyLine) {
  if (line.align === 'column-end' || line.precisePitch) return false;
  if (line.paperHeading) {
    const images = [...p.querySelectorAll<HTMLImageElement>('img:not(.ProseMirror-separator)')];
    const style = getComputedStyle(p), scale = documentScale(p);
    const width = p.getBoundingClientRect().width / scale - Math.min(0, parseFloat(style.textIndent) || 0);
    if (!p.parentElement?.matches('[data-column-count="2"]') || p.closest('li,td,th,[data-document-textbox]') ||
        p.matches('[data-print-paragraph-fragment],[data-print-soft-end]') ||
        hwpWrapFingerprint(p.textContent ?? '') !== line.paperHeading.fingerprint ||
        Math.abs(width - line.paperHeading.width) > .05 + 1 / 15 + 2 / 64 ||
        !images.length || !images.every(image => {
          const native = imageFloating(image.dataset.imageFloating)?.hwpPaperDrawing;
          return native && Math.abs(Number(image.getAttribute('width')) - native.width) <= .16 &&
            Math.abs(Number(image.getAttribute('height')) - native.height) <= .16;
        })) return false;
  } else if (p.querySelector('img')) return false;
  if (Boolean(p.closest('td,th')) !== Boolean(line.cell)) return false;
  if (line.align && (p.closest('li,[data-column-count]:not([data-column-count="1"])') || p.querySelector('u,s,del,mark'))) return false;
  // Column float widgets and caret separators are paint-only in the editor.
  const authoredObjects = p.querySelectorAll(objects);
  const hasObjects = line.paperHeading
    ? [...authoredObjects].some(node => !node.closest(viewFurniture) && !node.matches('.ProseMirror-trailingBreak'))
    : authoredObjects.length > 0;
  return !!p.textContent?.trim() && !hasObjects &&
    ![...p.querySelectorAll<HTMLElement>('[style]')].some(span => {
      if (span.closest(viewFurniture)) return false;
      const size = cssPixels(span.style.fontSize);
      return !!span.style.lineHeight || (line.align === 'top' && size !== null && Math.abs(size - line.font) >= .06);
    });
}
function applyStoredMixedPitch(paragraph: HTMLElement, line: HwpBodyLine) {
  if (!line.storedSingleLine || paragraph.style.height || paragraph.closest('li,[data-column-count]:not([data-column-count="1"])') ||
      paragraph.hasAttribute('data-print-paragraph-fragment')) return;
  const box = paragraph.getBoundingClientRect(), scale = documentScale(paragraph);
  const height = line.base * scale, lower = box.top + height;
  if (!(box.height > height + 1 / 64) || paragraphLineOffsets(paragraph).length) return;
  const range = document.createRange(); range.selectNodeContents(paragraph);
  const glyphBoxes = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
  if (glyphBoxes.length && glyphBoxes.every(rect => rect.top >= box.top && rect.bottom <= lower)) {
    paragraph.style.height = `${line.base}px`;
    paragraph.dataset.hwpStoredMixedPitch = 'true';
  }
}

/** Keep the native two-line strut at the unchanged width when the
 * laid-out lines already clear its lower edge. */
function storedPairHeight(paragraph: HTMLElement, line: HwpBodyLine, active = false) {
  const pair = line.mixedLinePair, uniform = line.storedUniformPair;
  if (!pair && !uniform || paragraph.dataset.hwpBodyLayout !== (uniform ? 'mixed-top' : 'true') ||
      paragraph.style.height && !active ||
      paragraph.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
      paragraph.querySelector('a,[id],mark,u,s,del')) return;
  const source = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), text = paragraph.textContent ?? '';
  if (!source?.bodyWidth || source.ends.length !== 2 || source.fingerprint !== hwpWrapFingerprint(text) ||
      source.ends[1] !== normalizeHwpWrapText(text).length) return;
  // The native pair consumes the same pitch on both lines. During page
  // fitting the browser can wrap at a different word from the final source
  // projection. Only compact the unchanged two-line text when its boxes fit.
  if (paragraphLineOffsets(paragraph).length !== 1) return;
  const box = paragraph.getBoundingClientRect(), style = getComputedStyle(paragraph), scale = documentScale(paragraph);
  const indent = parseFloat(style.textIndent) || 0, height = pair ? pair[0].pitch + pair[1].pitch : 2 * line.base;
  const difference = box.height / scale - height;
  const maxLeadingDifference = uniform ? 2 + 1 / 16 : 1;
  if (style.textAlign !== 'justify' || style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' ||
      style.transform !== 'none' || indent > 0 || Math.abs(box.width / scale - indent - source.bodyWidth) > .15 ||
      (active ? Math.abs(difference) > 1 / 64 : difference <= 1 / 64 || difference > maxLeadingDifference)) return;
  const range = document.createRange();
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let seen = false;
  while (walker.nextNode()) {
    const text = walker.currentNode;
    if (!text.textContent?.trim()) continue;
    range.selectNodeContents(text);
    for (const glyphs of range.getClientRects()) {
      if (!(glyphs.width > 0 && glyphs.height > 0)) continue;
      seen = true;
      // Uniform mixed-top runs can have font boxes above their actual
      // paint. Compacting below does not change their original top position.
      if ((!uniform && glyphs.top < box.top) || glyphs.bottom > box.top + height * scale ||
          glyphs.left < box.left + indent * scale - .0625 || glyphs.right > box.right + .0625) return;
    }
  }
  if (seen) return height;
}

const mixedLineSnapshots = new WeakMap<HTMLElement, Node[]>();
const bodyRunSpans = new WeakSet<HTMLElement>();

function layoutStoredMixedPair(paragraph: HTMLElement, line: HwpBodyLine) {
  const pair = line.mixedLinePair, body = paragraph.parentElement;
  if (!pair || !body?.matches('.print-content') || paragraph.style.height ||
      paragraph.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
      paragraph.querySelector('a,[id],mark,u,s,del')) return;
  const source = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), text = paragraph.textContent ?? '';
  if (!source?.bodyWidth || source.ends.length !== 2 || source.fingerprint !== hwpWrapFingerprint(text) ||
      source.ends[1] !== normalizeHwpWrapText(text).length || paragraphLineOffsets(paragraph).length !== 1) return;
  const before = paragraph.getBoundingClientRect(), bodyBox = body.getBoundingClientRect(), scale = documentScale(paragraph);
  const style = getComputedStyle(paragraph), indent = parseFloat(style.textIndent) || 0;
  const pitch = pair[0].pitch + pair[1].pitch;
  if (!(scale > 0) || style.textAlign !== 'justify' || style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' ||
      style.transform !== 'none' || indent > 0 || Math.abs(before.width / scale - indent - source.bodyWidth) > .15 ||
      Math.abs(before.height / scale - pitch) > 1 || before.top + pitch * scale > bodyBox.bottom) return;
  const offsets = sourceLineBoundaries(paragraph, source.ends.slice(0, 1));
  if (!offsets) return;
  const range = document.createRange(), projected = lineOpportunities(paragraph, offsets, range);
  const lines = [...projected.children].filter((element): element is HTMLSpanElement => element instanceof HTMLSpanElement);
  if (lines.length !== 2) return;
  for (const [index, element] of lines.entries()) {
    const { pitch: advance } = pair[index];
    Object.assign(element.style, { display: 'block', height: `${advance}px`, textIndent: index ? '0px' : `${indent}px` });
    element.style.setProperty('line-height', `${advance}px`, 'important');
  }
  const original = [...paragraph.childNodes];
  paragraph.replaceChildren(...lines);
  const marker = document.createElement('span');
  marker.style.cssText = 'display:inline-block;vertical-align:baseline;width:0;height:0;font-size:0;padding:0;border:0;margin:0';
  marker.style.setProperty('line-height', '0px', 'important');
  const safe = lines.every((element, index) => {
    const frame = element.getBoundingClientRect(), runs = [...element.querySelectorAll<HTMLElement>('[data-hwp-body-run]')];
    if (!runs.length || Math.abs(frame.height / scale - pair[index].pitch) > .025) return false;
    element.append(marker);
    const offset = pair[index].baseline - (marker.getBoundingClientRect().bottom - frame.top) / scale;
    marker.remove();
    if (!Number.isFinite(offset) || Math.abs(offset) > 4) return false;
    for (const run of runs) Object.assign(run.style, { position: 'relative', top: `${offset}px` });
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
    let visible = false, fontBoxOutside = false;
    while (walker.nextNode()) {
      const text = walker.currentNode;
      if (!text.textContent?.trim()) continue;
      range.selectNodeContents(text);
      const boxes = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
      if (boxes.length !== 1) return false;
      const [box] = boxes;
      if (!(box.left >= bodyBox.left - .0625 && box.right <= bodyBox.right + .0625)) return false;
      fontBoxOutside ||= box.top < frame.top || box.bottom > frame.bottom;
      visible = true;
    }
    if (!visible || !fontBoxOutside) return visible;
    const ink = measuredHwpLineInk(element, scale, line.font);
    return !!ink && ink.top >= frame.top && ink.bottom <= frame.bottom;
  });
  if (!safe || paragraph.textContent !== text) paragraph.replaceChildren(...original);
  else mixedLineSnapshots.set(paragraph, original);
}

/** Static print/measurement DOM only. The editor uses decorations below, so
 * neither wrapping spans nor measured line boundaries enter the document. */
export function layoutHwpBodyLines(root: HTMLElement, paintSourcePairs = false) {
  for (const p of root.querySelectorAll<HTMLElement>('p[data-hwp-inline-table-spaces="true"]')) {
    p.toggleAttribute('data-hwp-inline-table-spaces-layout', /^ +$/.test(p.textContent ?? ''));
  }
  for (const p of root.matches('[data-hwp-body-line]') ? [root] : root.querySelectorAll<HTMLElement>('[data-hwp-body-line]')) {
    const line = parseHwpBodyLine(p.dataset.hwpBodyLine); if (!line || line.empty) continue;
    const earlier = mixedLineSnapshots.get(p);
    if (earlier) { p.replaceChildren(...earlier); mixedLineSnapshots.delete(p); }
    if (p.dataset.hwpStoredMixedPitch === 'true') {
      p.style.removeProperty('height');
      delete p.dataset.hwpStoredMixedPitch;
    }
    const retained = new Set<HTMLElement>();
    for (const span of p.querySelectorAll<HTMLElement>('[data-hwp-body-run]')) {
      // A plain wrapper has no geometry of its own outside the body-line
      // styles below. Keep it across fits; moves and edited runs still
      // recheck the paragraph policy and any measured pitch.
      if (bodyRunSpans.has(span) && span.attributes.length === 1 && span.childNodes.length === 1 &&
          span.firstChild instanceof Text && span.firstChild.length) {
        retained.add(span);
        continue;
      }
      delete span.dataset.hwpBodyRun; span.style.removeProperty('line-height');
      if (![...span.attributes].some(a => a.name !== 'style') && !span.style.length) span.replaceWith(...span.childNodes);
    }
    delete p.dataset.hwpBodyLayout; p.style.removeProperty('--hwp-body-ratio'); p.style.removeProperty('--hwp-body-ink-offset');
    p.style.lineHeight = `${line.base}px`;
    if (!eligible(p, line)) {
      for (const span of retained) span.replaceWith(span.firstChild!);
      continue;
    }
    const texts: Text[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      if (text.length && (!text.parentElement || !retained.has(text.parentElement))) texts.push(text);
    }
    for (const text of texts) {
      const span = document.createElement('span'); span.dataset.hwpBodyRun = 'true';
      bodyRunSpans.add(span);
      text.replaceWith(span); span.append(text);
    }
    p.dataset.hwpBodyLayout = line.paperHeading ? 'paper-heading' : line.align ?? 'true';
    if (line.align === 'top') continue;
    p.style.setProperty('--hwp-body-ratio', String(line.percent / 100));
    if (line.align === 'mixed-top') p.style.setProperty('--hwp-body-ink-offset', `${(line.font - line.base) / 2}px`);
    // A paragraph-wide strut based on the large first run otherwise keeps
    // every later, smaller line at the large font's height. All real text
    // retains its font size and supplies its own percentage line box.
    p.style.lineHeight = '0';
    applyStoredMixedPitch(p, line);
    if (paintSourcePairs && line.storedUniformPair) {
      const height = storedPairHeight(p, line);
      if (height !== undefined) {
        p.style.height = `${height}px`;
        p.dataset.hwpStoredMixedPitch = 'true';
      }
    } else if (paintSourcePairs) layoutStoredMixedPair(p, line);
  }
}

/** Recover an original body line only while its text, hanging indent and
 * available paper area still agree. The line breaks exist only in the copy. */
export function restoreHwpAuthoredBodyLines(paragraph: HTMLParagraphElement) {
  const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), authored = proof?.authoredBodyLines;
  const text = paragraph.textContent || '';
  if (!proof || !authored || hwpWrapFingerprint(text) !== proof.fingerprint ||
      paragraph.closest('td,th,[data-column-count],[data-document-note],[data-document-textbox]') ||
      paragraph.querySelector('br,img,svg,a,[data-docx-tab],[data-document-field],[data-note-reference],[data-docx-bookmark]')) return;
  const { origin } = authored;
  const page = paragraph.closest<HTMLElement>('.print-page'), body = paragraph.closest('.print-content');
  if (!page || !body || Number(page.dataset.page) !== origin.page) return;
  const paper = page.getBoundingClientRect(), bounds = body.getBoundingClientRect(), box = paragraph.getBoundingClientRect();
  const style = getComputedStyle(paragraph);
  const hanging = -parseFloat(style.textIndent), lineCount = paragraphLineOffsets(paragraph).length + 1;
  const policy = parseHwpBodyLine(paragraph.dataset.hwpBodyLine);
  const pitch = parseFloat(style.lineHeight) || policy?.base || 0;
  // Mixed HWP runs may leave a zero CSS paragraph strut. The native
  // policy and the current measured lines must agree before using it.
  if (!pitch || Math.abs(box.height - lineCount * pitch) > .25) return;
  const relativeTop = box.top - paper.top - origin.y;
  // An early paragraph can be missing native vertical advance. Adding an
  // advance to a paragraph already below its source start compounds drift.
  // Account for source tenths, a DOCX twip and both CSS layout edges.
  if (relativeTop > .05 + 1 / 15 + 2 / 64 || relativeTop <= -authored.pitch) return;
  if (style.textAlign !== 'justify' || Math.abs(hanging - authored.hanging) > .1 ||
      Math.abs(parseFloat(style.marginLeft) - hanging) > .1 || parseFloat(style.marginRight) ||
      Math.abs(parseFloat(style.fontSize) - authored.size) > .07 || Math.abs(pitch - authored.pitch) > .2 ||
      Math.abs(bounds.width - authored.width) > .6 || Math.abs(box.left - hanging - paper.left - origin.x) > .6 ||
      box.bottom + pitch > bounds.bottom + .5 ||
      lineCount + 1 !== proof.ends.length) return;
  const mayRestoreStrut = paragraph.dataset.hwpBodyLayout === 'true' && policy && !policy.align &&
    !policy.storedSingleLine && !policy.mixedLinePair && policy.base > policy.font;
  const ends = proof.ends.slice(0, -1), boundaries = sourceLineBoundaries(paragraph, ends);
  if (!boundaries || normalizeHwpWrapText(text).length !== proof.ends.at(-1)) return;
  const before = paragraph.innerHTML, range = document.createRange();
  for (const boundary of boundaries.slice(1, -1).reverse()) {
    range.setStart(boundary.node, boundary.at);
    range.collapse(true);
    range.insertNode(document.createElement('br'));
  }
  const restore = () => {
    paragraph.innerHTML = before;
    if (policy) layoutHwpBodyLines(paragraph, true);
  };
  if (policy) layoutHwpBodyLines(paragraph, true);
  // The synthetic break disables percentage wrapping. With fixed leading
  // on both the paragraph and its mixed-size runs, their different baselines
  // can still expand the new line. Let the paragraph supply the strut and
  // accept it only when every rendered line keeps its ink inside that pitch.
  const useSourceStrut = mayRestoreStrut && paragraph.getBoundingClientRect().height - box.height - pitch > .2;
  if (useSourceStrut) {
    for (const inline of paragraph.querySelectorAll<HTMLElement>('span,strong,em,b,i,u,s,mark'))
      inline.style.lineHeight = '1em';
  }
  const after = paragraph.getBoundingClientRect(), actual = paragraphLineOffsets(paragraph);
  const ink = useSourceStrut ? measuredHwpInkLines(paragraph, 1, authored.size) : undefined;
  const fitsSourceStrut = !useSourceStrut || ink?.length === proof.ends.length && ink.every((line, index) =>
    line.top >= after.top + index * pitch - 1 / 64 && line.bottom <= after.top + (index + 1) * pitch + 1 / 64);
  // Print offsets include each synthetic BR. Source offsets count only
  // the paragraph's literal text.
  if (!fitsSourceStrut || paragraph.textContent !== text || Math.abs(after.height - box.height - pitch) > .2 ||
      actual.length !== ends.length || actual.some((offset, index) => normalizeHwpWrapText(text.slice(0, offset - index - 1)).length !== ends[index]) ||
      !horizontalInkFits(paragraph, range, bounds)) {
    restore();
    return;
  }
  return restore;
}

const unsafeAuthoredBodyFlow = 'img[data-image-floating],.table-wrap-frame,[data-column-count],[data-document-textbox],' +
  '[data-note-reference],[data-print-soft-end],[data-hwp-projected-body-end]';

/** Keep an intact source paragraph on its original sheet before a forced
 * break. The preceding paragraph can give back a line only if the missing
 * face's fallback naturally produces all of its recorded line ends. */
export function restoreHwpBodyBeforePageBreak(last: HTMLParagraphElement, fits: () => boolean, permitPrintBreak = true) {
  const body = last.parentElement, page = body?.closest<HTMLElement>('.print-page');
  if (!body?.matches('.print-content') || !page || body.querySelector(unsafeAuthoredBodyFlow)) return false;
  const paper = page.getBoundingClientRect(), bounds = body.getBoundingClientRect();
  const precision = .05 + 1 / 15 + 2 / 64;
  const unsupported = ':not(span),[id],[data-document-field],[data-word-rotation],[data-flow-widget]';
  const read = (paragraph: HTMLParagraphElement) => {
    const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
    const source = proof?.authoredBodyLines, text = paragraph.textContent || '';
    if (!proof || !source || source.origin.page !== Number(page.dataset.page) ||
        hwpWrapFingerprint(text) !== proof.fingerprint ||
        paragraph.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
        paragraph.querySelector(unsupported)) return;
    const box = paragraph.getBoundingClientRect(), css = getComputedStyle(paragraph);
    const hanging = -parseFloat(css.textIndent), pitch = parseFloat(css.lineHeight);
    if (css.textAlign !== 'justify' || css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' ||
        css.transform !== 'none' || Math.abs(hanging - source.hanging) > precision ||
        Math.abs(box.left - hanging - paper.left - source.origin.x) > precision ||
        Math.abs(box.width + hanging - source.width) > precision || Math.abs(pitch - source.pitch) > precision ||
        Math.abs(box.top - paper.top - source.origin.y) >= source.pitch) return;
    return { proof, source, text, box, pitch };
  };
  const ending = read(last);
  if (!ending || !matchesStoredWrap(last, ending.text, ending.proof.ends)) return false;
  const range = document.createRange();
  for (let paragraph = last.previousElementSibling; paragraph instanceof HTMLParagraphElement;
    paragraph = paragraph.previousElementSibling) {
    const witness = read(paragraph);
    if (!witness || witness.text.length > 160 || witness.proof.ends.length !== 2 ||
        normalizeHwpWrapText(witness.text).length !== witness.proof.ends[1] ||
        paragraphLineOffsets(paragraph).length !== 2 ||
        Math.abs(witness.box.height - 3 * witness.pitch) > precision) continue;
    const inline = [paragraph, ...paragraph.querySelectorAll<HTMLElement>('span')];
    const family = (value: string) => value.replace(/^['"]|['"]$/g, '');
    if (inline.some(node => {
      const css = getComputedStyle(node);
      const sourceTop = node !== paragraph && paragraph.dataset.hwpBodyLayout === 'top' && css.verticalAlign === 'top';
      return family(css.fontFamily) !== hwpSingleBodyFace || Math.abs(parseFloat(css.fontSize) - witness.source.size) > .07 ||
        css.fontStyle !== 'normal' || css.fontWeight !== '400' || (css.verticalAlign !== 'baseline' && !sourceTop) ||
        css.textDecorationLine !== 'none' || css.textShadow !== 'none';
    })) continue;
    const declarations = inline.filter(node => node.style.fontFamily);
    const original = declarations.map(node => node.style.fontFamily);
    for (const node of declarations) node.style.fontFamily += ', "Noto Sans KR"';
    const children = [...paragraph.childNodes];
    const firstEnd = witness.proof.ends[0];
    let exact = matchesStoredWrap(paragraph, witness.text, witness.proof.ends);
    // When the fallback already fits two lines, its first line can still
    // end on a different glyph before DOCX rounds the original tracking.
    // Native export reads this DOM as editable content. Only the print
    // preview can use a synthetic explicit boundary.
    if (permitPrintBreak && !exact && paragraphLineOffsets(paragraph).length === 1) {
      const cuts = sourceLineBoundaries(paragraph, [firstEnd]);
      if (cuts) {
        const lines = lineOpportunities(paragraph, cuts, range);
        lines.querySelector('wbr')!.replaceWith(document.createElement('br'));
        paragraph.replaceChildren(lines);
        const offsets = paragraphLineOffsets(paragraph);
        exact = offsets.length === 1 && offsets[0] === cuts[1].offset + 1;
      }
    }
    const box = paragraph.getBoundingClientRect();
    range.selectNodeContents(paragraph);
    const glyphs = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
    const predecessorBottom = paragraph.previousElementSibling?.getBoundingClientRect().bottom ?? bounds.top;
    const safe = box.width === witness.box.width && box.top === witness.box.top &&
        Math.abs(box.height - 2 * witness.pitch) <= precision &&
        exact && paragraph.textContent === witness.text &&
        glyphs.length && glyphs.every(rect => rect.top >= Math.max(bounds.top, predecessorBottom) && rect.bottom <= box.bottom) &&
        horizontalInkFits(paragraph, range, bounds);
    if (safe && fits()) return true;
    if (paragraph.firstChild !== children[0]) paragraph.replaceChildren(...children);
    declarations.forEach((node, index) => { node.style.fontFamily = original[index]; });
    if (safe) fits();
  }
  return false;
}

/** A measured next-page table may fit too early when its preceding body
 * has lost authored lines. Restore those lines only if they fit together
 * and the intact table is the only element driven beyond this sheet. */
export function restoreHwpBodyBeforeKeptTable(body: HTMLElement, table: HTMLTableElement) {
  const blank = table.previousElementSibling;
  if (!(blank instanceof HTMLParagraphElement) || blank.firstElementChild || blank.textContent?.trim() ||
      body.querySelector(unsafeAuthoredBodyFlow)) return false;
  const paragraphs: HTMLParagraphElement[] = [];
  for (let node = blank.previousElementSibling; node instanceof HTMLParagraphElement; node = node.previousElementSibling) {
    if (node.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]')) break;
    if (node.dataset.hwpSourceWrap?.includes('authoredBodyLines')) paragraphs.unshift(node);
  }
  const restorations: { paragraph: HTMLParagraphElement; undo: () => void }[] = [];
  for (const paragraph of paragraphs) {
    const undo = restoreHwpAuthoredBodyLines(paragraph);
    if (undo) restorations.push({ paragraph, undo });
  }
  if (!restorations.length) return false;
  const bounds = body.getBoundingClientRect(), limit = bounds.bottom + .5;
  const range = document.createRange();
  const keepsBody = table.getBoundingClientRect().bottom > limit && blank.getBoundingClientRect().bottom <= limit &&
    restorations.every(({ paragraph }) => {
      const previous = paragraph.previousElementSibling?.getBoundingClientRect();
      range.selectNodeContents(paragraph);
      const glyphs = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
      return glyphs.length > 0 && glyphs.every(rect =>
        rect.top >= Math.max(bounds.top, previous?.bottom ?? bounds.top) - .5 && rect.bottom <= limit);
    });
  if (!keepsBody) for (const { undo } of restorations.reverse()) undo();
  return keepsBody;
}

/** Use the authored wrap only if the same line boxes and neighboring
 * blocks stay in place. The break opportunities belong to the print copy. */
function restoreHwpBodyLineEnds(paragraph: HTMLParagraphElement, range: Range) {
  const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
  const source = proof?.authoredBodyLines, text = paragraph.textContent ?? '';
  const paper = paragraph.closest<HTMLElement>('.print-page'), body = paragraph.parentElement;
  if (!source || !proof || !body || !paper || source.origin.page !== Number(paper.dataset.page) ||
      hwpWrapFingerprint(text) !== proof.fingerprint ||
      paragraph.matches('[data-paragraph-border-box],[data-word-paragraph-bottom-border],[data-print-soft-end]') ||
      paragraph.querySelector(':not(span,strong,b,em,i,u),[id],[data-flow-widget],[data-document-field],[data-word-overlap]')) return;
  const lineCount = paragraphLineOffsets(paragraph).length + 1;
  if (lineCount !== proof.ends.length || matchesStoredWrap(paragraph, text, proof.ends)) return;
  const style = getComputedStyle(paragraph);
  const box = paragraph.getBoundingClientRect(), bounds = body.getBoundingClientRect(), page = paper.getBoundingClientRect();
  const hanging = -parseFloat(style.textIndent);
  if (style.textAlign !== 'justify' || style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' ||
      style.transform !== 'none' || Math.abs(hanging - source.hanging) > .1 ||
      Math.abs(parseFloat(style.marginLeft) - hanging) > .1 || parseFloat(style.marginRight) ||
      Math.abs(box.width + hanging - source.width) > .15 ||
      Math.abs(box.left - hanging - page.left - source.origin.x) > .15 ||
      Math.abs(parseFloat(style.fontSize) - source.size) > .07 ||
      Math.abs(box.height - lineCount * source.pitch) > .25 ||
      normalizeHwpWrapText(text).length !== proof.ends.at(-1)) return;
  const cuts = sourceLineBoundaries(paragraph, proof.ends.slice(0, -1));
  if (!cuts) return;

  const before = retainedTextBands(paragraph, range, bounds);
  if (!before) return;
  const next = paragraph.nextElementSibling, origin = next?.getBoundingClientRect();
  const old = [...paragraph.childNodes];
  paragraph.replaceChildren(lineOpportunities(paragraph, cuts, range));
  const changed = paragraph.getBoundingClientRect(), following = next?.getBoundingClientRect(), paint = retainedTextBands(paragraph, range, bounds);
  if (!paint || paint.size !== before.size || [...paint].some(band => !before.has(band)) ||
      box.x !== changed.x || box.y !== changed.y || box.width !== changed.width || box.height !== changed.height ||
      origin && following && (origin.x !== following.x || origin.y !== following.y ||
        origin.width !== following.width || origin.height !== following.height) ||
      paragraph.textContent !== text || !matchesStoredWrap(paragraph, text, proof.ends) ||
      !horizontalInkFits(paragraph, range, bounds)) paragraph.replaceChildren(...old);
}

/** Recover source lines on an already placed sheet only when the rest of
 * its ordinary body still fits. These alternate lines may not create a
 * new page or push a later section across its source boundary. */
export function layoutHwpAuthoredBodyLines(root: HTMLElement) {
  const candidates = root.querySelectorAll<HTMLParagraphElement>(
    '.print-content > p[data-hwp-source-wrap*="authoredBodyLines"]',
  );
  if (!candidates.length) return;
  const simple = new WeakMap<HTMLElement, boolean>();
  const range = document.createRange();
  for (const paragraph of candidates) {
    const body = paragraph.parentElement!;
    let eligible = simple.get(body);
    if (eligible === undefined) {
      eligible = !body.querySelector(unsafeAuthoredBodyFlow);
      simple.set(body, eligible);
    }
    if (!eligible || paragraph.hasAttribute('data-print-paragraph-fragment')) continue;
    restoreHwpBodyLineEnds(paragraph, range);
    const limit = body.getBoundingClientRect();
    const restore = restoreHwpAuthoredBodyLines(paragraph);
    if (!restore) continue;
    const previous = paragraph.previousElementSibling?.getBoundingClientRect();
    const siblings = [...body.children];
    const moving = siblings.slice(siblings.indexOf(paragraph));
    range.selectNodeContents(paragraph);
    const ink = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
    if (!ink.length || ink.some(box => box.top < Math.max(limit.top, previous?.bottom ?? limit.top) - .5 ||
          box.bottom > limit.bottom + .5) ||
        moving.some(element => element.getBoundingClientRect().bottom > limit.bottom + .5)) restore();
  }
}

/** A complete source line at a natural page end may use its ink extent.
 * Its full paragraph still occupies flow space for the following content. */
export function sourceBodyEndBottom(paragraph: HTMLElement, body: HTMLElement): number | undefined {
  const policy = parseHwpBodyLine(paragraph.dataset.hwpBodyLine), source = policy?.bodyEndFit;
  const previous = paragraph.previousElementSibling;
  if (!source || !policy || paragraph.parentElement !== body || paragraph.dataset.hwpBodyLayout !== 'mixed-top' ||
      !(previous instanceof HTMLParagraphElement) || paragraphLineOffsets(paragraph).length || paragraphLineOffsets(previous).length ||
      [paragraph, previous].some(node => node.tagName !== 'P' || node.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
        node.querySelector(objects + ',a,mark,u,s,del'))) return;
  if (hwpWrapFingerprint(paragraph.textContent ?? '') !== source.fingerprint ||
      hwpWrapFingerprint(previous.textContent ?? '') !== source.preceding) return;
  const box = paragraph.getBoundingClientRect(), prior = previous.getBoundingClientRect(), bounds = body.getBoundingClientRect();
  const style = getComputedStyle(paragraph), before = getComputedStyle(previous), scale = documentScale(paragraph);
  const indent = parseFloat(style.textIndent) || 0, room = (bounds.bottom - box.top) / scale;
  if (!(scale > 0) || !(room > 0 && room < policy.base) ||
      Math.abs(room - source.remainder) > .25 || Math.abs(box.width / scale - Math.min(0, indent) - source.width) > .15 ||
      Math.abs(box.top - prior.bottom) > .08 * scale || Math.abs(box.height / scale - policy.base) > .1 ||
      [style, before].some(css => css.transform !== 'none' || css.backgroundColor !== 'rgba(0, 0, 0, 0)' ||
        css.textShadow !== 'none' || parseFloat(css.borderTopWidth) || parseFloat(css.borderBottomWidth))) return;
  const limit = policy.maxFont ?? policy.font;
  const ink = measuredHwpLineInk(paragraph, scale, limit), precedingInk = measuredHwpLineInk(previous, scale, limit);
  if (!ink || !precedingInk) return;
  const delta = box.top + source.baseline * scale - ink.baseline;
  const top = ink.top + delta, bottom = ink.bottom + delta;
  // The outline measurement already includes a sample pixel safety margin.
  // Require both pieces of real ink to stay clear of each other and the body.
  if (!(delta < -.05 * scale && delta >= -4 * scale && top >= bounds.top && bottom <= bounds.bottom &&
        top - precedingInk.bottom >= .5 * scale)) return;
  paragraph.style.transform = `translateY(${delta / scale}px)`;
  paragraph.dataset.hwpProjectedBodyEnd = 'true';
  return bottom;
}

export function sourceForcedBodyEndBottom(paragraph: HTMLElement, body: HTMLElement, projectPaint = true): number | undefined {
  if (paragraph.dataset.hwpPrintBeforeExplicitBreak !== 'true' || paragraph.parentElement !== body) return;
  const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), source = proof?.forcedBodyEnd;
  const previous = paragraph.previousElementSibling;
  if (!proof || !source || !(previous instanceof HTMLParagraphElement) ||
      [paragraph, previous].some(node => node.tagName !== 'P' ||
        node.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
        node.querySelector(objects + ',a,mark,u,s,del,sup,sub')) ||
      hwpWrapFingerprint(paragraph.textContent ?? '') !== proof.fingerprint ||
      hwpWrapFingerprint(previous.textContent ?? '') !== source.preceding ||
      paragraphLineOffsets(paragraph).length !== 0 || paragraphLineOffsets(previous).length + 1 !== source.precedingLines) return;
  const box = paragraph.getBoundingClientRect(), prior = previous.getBoundingClientRect(), bounds = body.getBoundingClientRect();
  const css = getComputedStyle(paragraph), previousStyle = getComputedStyle(previous);
  const scale = documentScale(paragraph), room = (bounds.bottom - box.top) / scale;
  if (!(scale > 0 && room > 0 && room < source.pitch) || Math.abs(room - source.room) > .6 ||
      Math.abs(box.width / scale + source.hanging - source.width) > .2 ||
      Math.abs(parseFloat(css.textIndent) + source.hanging) > .08 ||
      Math.abs((box.top - prior.bottom) / scale - source.spaceBefore) > .15 ||
      Math.abs(box.height / scale - source.pitch) > .16 ||
      [css, previousStyle].some(style => style.transform !== 'none' || style.backgroundColor !== 'rgba(0, 0, 0, 0)' ||
        style.textShadow !== 'none' || parseFloat(style.borderTopWidth) || parseFloat(style.borderBottomWidth))) return;
  const ink = measuredHwpLineInk(paragraph, scale, source.font, true);
  if (!ink) return;
  const range = document.createRange();
  range.selectNodeContents(previous);
  const previousInk = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
  if (!previousInk.length || !horizontalInkFits(paragraph, range, bounds)) return;
  const delta = box.top + source.baseline * scale - ink.baseline;
  const top = ink.top + delta, bottom = ink.bottom + delta;
  if (!(delta < -.05 * scale && delta > (source.font - source.pitch) * scale &&
        top >= Math.max(bounds.top, ...previousInk.map(rect => rect.bottom)) + .5 * scale && bottom <= bounds.bottom)) return;
  if (projectPaint) {
    paragraph.style.transform = `translateY(${delta / scale}px)`;
    paragraph.dataset.hwpProjectedBodyEnd = 'true';
  }
  return bottom;
}

/** A stored two-line HWP page tail may leave its empty terminal leading
 * outside the content box. Keep the layout and accept only its contained ink. */
export function sourceAuthoredTailInkBottom(paragraph: HTMLElement, body: HTMLElement) {
  const wrap = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap);
  const proof = wrap?.authoredTailInk, paper = body.closest<HTMLElement>('.print-page');
  if (!proof || !paper || paragraph.parentElement !== body || proof.origin.page !== Number(paper.dataset.page) ||
      paragraph.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
      paragraph.querySelector('br,img,svg,mark,s,del,sup,sub,a,[data-note-reference],[data-document-field],[data-docx-tab],[data-word-overlap]') ||
      hwpWrapFingerprint(paragraph.textContent ?? '') !== wrap.fingerprint ||
      paragraphLineOffsets(paragraph).length !== 1) return;
  const page = paper.getBoundingClientRect(), bounds = body.getBoundingClientRect(), box = paragraph.getBoundingClientRect();
  const css = getComputedStyle(paragraph);
  const trailing = box.bottom - bounds.bottom;
  if (css.textAlign !== 'justify' || css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' ||
      css.transform !== 'none' || css.textShadow !== 'none' || css.backgroundColor !== 'rgba(0, 0, 0, 0)' ||
      Math.abs(parseFloat(css.fontSize) - proof.size) > .08 || Math.abs(parseFloat(css.lineHeight) - proof.pitch) > .12 ||
      Math.abs(parseFloat(css.textIndent) + proof.hanging) > .15 || Math.abs(parseFloat(css.marginRight) - proof.rightMargin) > .15 ||
      Math.abs(box.width + proof.hanging - proof.width) > .2 || Math.abs(box.height - 2 * proof.pitch) > .15 ||
      Math.abs(box.top - page.top - proof.origin.y) >= proof.pitch ||
      !(trailing > 0 && trailing <= proof.pitch - proof.size)) return;

  const range = document.createRange(), walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let bottom = -Infinity;
  while (walker.nextNode()) {
    const node = walker.currentNode;
    if (!node.textContent?.trim() || !node.parentElement) continue;
    const style = getComputedStyle(node.parentElement);
    if (style.textShadow !== 'none' || style.backgroundColor !== 'rgba(0, 0, 0, 0)') return;
    range.selectNodeContents(node);
    const reserve = node.parentElement.closest('u,[data-word-underline]') ? 1 : 0;
    for (const rect of range.getClientRects()) {
      if (!(rect.width > .1 && rect.height > 1)) continue;
      if (rect.top < bounds.top || rect.left < bounds.left - .15 || rect.right > bounds.right + .15 ||
          rect.bottom + reserve > bounds.bottom) return;
      bottom = Math.max(bottom, rect.bottom + reserve);
    }
  }
  if (Number.isFinite(bottom)) return bottom;
}

export function clearSourceBodyEndPaint(paragraph: HTMLElement) {
  if (paragraph.dataset.hwpProjectedBodyEnd === 'true') {
    paragraph.style.removeProperty('transform');
    delete paragraph.dataset.hwpProjectedBodyEnd;
  }
}

type ColumnBounds = { room: number; width: number; height: number; hanging: boolean; scale: number };
type CachedColumn = { bounds: ColumnBounds; predecessor: ModelNode; placement: SourceColumnPlacement | null };
const columnLayoutRounding = 4 / 64;

function sameColumnBounds(a: ColumnBounds, b: ColumnBounds, tolerance: number) {
  return a.hanging === b.hanging && a.scale === b.scale &&
    Math.abs(a.room - b.room) <= tolerance && Math.abs(a.width - b.width) <= tolerance &&
    Math.abs(a.height - b.height) <= tolerance;
}
const key = new PluginKey('hwpBodyLineMetrics');
export const HwpBodyLineMetrics = Extension.create({
  name: 'hwpBodyLineMetrics',
  addGlobalAttributes: () => [{ types: ['paragraph', 'heading'], attributes: {
    hwpBodyLine: { default: null,
      parseHTML: el => { const line = parseHwpBodyLine(el.dataset.hwpBodyLine || el.style.getPropertyValue('--hwp-body-line')); return line ? encodeHwpBodyLine(line) : null; },
      renderHTML: attrs => { const line = activeHwpBodyLine(attrs); return line ? { 'data-hwp-body-line': encodeHwpBodyLine(line) } : {}; },
    },
  } }, { types: ['paragraph'], attributes: {
    hwpInlineTableSpaces: { default: false,
      parseHTML: el => el.dataset.hwpInlineTableSpaces === 'true',
      renderHTML: attrs => attrs.hwpInlineTableSpaces ? { 'data-hwp-inline-table-spaces': 'true' } : {},
    },
  } }],
  addProseMirrorPlugins: () => [new Plugin({
    key,
    state: { init: () => DecorationSet.empty, apply: (tr, old) => {
      const explicit = tr.getMeta(key); if (explicit) return explicit;
      const mapped = old.map(tr.mapping, tr.doc);
      if (!tr.docChanged) return mapped;
      // Restore natural room as soon as this paragraph's content or formatting
      // changes, including the first transaction of a composed edit.
      return mapped.remove(mapped.find().filter(decoration => decoration.spec.mixedPitchNode &&
        tr.doc.nodeAt(decoration.from) !== decoration.spec.mixedPitchNode));
    } },
    props: { decorations: state => key.getState(state) },
    view(view) {
      let frame = 0, signature = '', lastDocument = view.state.doc;
      let sourceColumns = new WeakMap<ModelNode, CachedColumn>();
      const schedule = () => {
        // Pagination and resize can dispatch again before this frame runs.
        if (frame || editorLayoutSuspended()) return;
        frame = requestAnimationFrame(() => {
          frame = 0;
          if (editorLayoutSuspended() || !view.dom.isConnected || view.composing) return;
          const decorations: Decoration[] = [], values: string[] = [];
          let papers: DOMRect[] | undefined;
          view.state.doc.descendants((node, pos) => {
            if (node.type.name === 'paragraph' && node.attrs.hwpInlineTableSpaces && /^ +$/.test(node.textContent)) {
              decorations.push(Decoration.node(pos, pos + node.nodeSize, { 'data-hwp-inline-table-spaces-layout': 'true' }));
              values.push(`${pos}:table-spaces`);
              return;
            }
            // Nodes without a body policy can include an entire table or
            // section. Avoid serializing those descendants just to reject it.
            if (!node.attrs.hwpBodyLine) return;
            const line = hwpBodyTextLine(node.toJSON()); if (!line) return;
            const dom = view.nodeDOM(pos);
            if (!(dom instanceof HTMLElement)) return;
            if (line.align === 'column-end') {
              if (!followsWithHwpColumnBreak(dom)) return;
              const previous = dom.previousElementSibling;
              const predecessor = view.state.doc.resolve(pos).nodeBefore;
              const scale = documentScale(dom.parentElement ?? dom);
              const tail = line.columnFit && sourceColumnTail(dom, line);
              if (tail && predecessor?.type.name === 'paragraph' && predecessor.textContent === tail.textContent) {
                decorations.push(Decoration.inline(pos - 2, pos - 1, { nodeName: 'span', 'data-hwp-column-tail': 'true' }));
                values.push(`${pos}:column-tail`);
              }

              const active = dom.dataset.hwpColumnPaint === 'true';
              const ordinary = active ? null : measureHwpColumnEnd(dom, line, scale);
              let placement: SourceColumnPlacement | null = null;
              if (line.columnFit && previous instanceof HTMLElement && predecessor?.type.name === 'paragraph' &&
                  previous.textContent === predecessor.textContent) {
                const stack = dom.closest<HTMLElement>('.document-stack'), owner = dom.closest<HTMLElement>('[data-bottom-margin]');
                const inset = Number(owner?.dataset.bottomMargin), preceding = previous.getBoundingClientRect();
                papers ??= [...(stack?.querySelectorAll('.section-paper,.paper-backgrounds .paper') ?? [])]
                  .map(paper => paper.getBoundingClientRect()).filter(box => box.height > 0);
                const paper = papers.find(box => preceding.top >= box.top - .05 && preceding.top < box.bottom);
                if (paper && Number.isFinite(inset) && previous.getClientRects().length === 1) {
                  const bottom = paper.bottom - inset * scale;
                  const bounds: ColumnBounds = { room: (bottom - preceding.bottom) / scale,
                    width: preceding.width / scale, height: preceding.height / scale, scale,
                    hanging: Boolean(tail && previous.querySelector('[data-hwp-column-tail]')) };
                  const cached = sourceColumns.get(node);
                  if (cached?.predecessor === predecessor) {
                    // Paper gaps and columns are laid out in separate passes.
                    // Keep positive geometry through subpixel rounding only
                    // while the prior measured line still fits the live body.
                    const rounding = active && cached.placement ? columnLayoutRounding : .5 / 64;
                    if (sameColumnBounds(cached.bounds, bounds, rounding) &&
                        (!cached.placement || cached.placement.height + .05 <= bounds.room))
                      placement = cached.placement;
                  }
                  // Re-evaluate the undecorated paragraph after a paper or font
                  // change so the temporary transform cannot qualify itself.
                  if (!placement && !active && (cached?.predecessor !== predecessor ||
                      !sameColumnBounds(cached.bounds, bounds, .5 / 64) || cached.placement)) {
                    const measured = measureSourceColumnEnd(dom, line, scale, bottom, ordinary);
                    placement = measured && measured.height + .05 <= bounds.room ? measured : null;
                    sourceColumns.set(node, { bounds, predecessor, placement });
                  }
                }
              }
              const height = placement?.height ?? ordinary;
              if (height !== null) {
                const paint = placement ? sourceColumnPaintStyle(placement) : '';
                decorations.push(Decoration.node(pos, pos + node.nodeSize, { 'data-hwp-column-end': 'true',
                  ...(placement ? { 'data-hwp-column-paint': 'true' } : {}),
                  ...(placement?.scale ? { 'data-hwp-column-scaled': 'true' } : {}),
                  style: `--hwp-column-em:${line.font}px;--hwp-column-height:${height}px${paint ? ';' + paint : ''}` }));
                values.push(`${pos}:column-end:${height}:${paint}`);
              }
              return;
            }
            if (!eligible(dom, line)) return;
            let style = line.align === 'top' ? `line-height:${line.base}px` : `line-height:0;--hwp-body-ratio:${line.percent / 100}`;
            if (line.align === 'mixed-top') style += `;--hwp-body-ink-offset:${(line.font - line.base) / 2}px`;
            const mixedPitch = storedPairHeight(dom, line, dom.dataset.hwpEditorMixedPitch === 'true');
            if (mixedPitch !== undefined) style += `;height:${mixedPitch}px`;
            decorations.push(Decoration.node(pos, pos + node.nodeSize,
              { 'data-hwp-body-layout': line.paperHeading ? 'paper-heading' : line.align ?? 'true', style,
                ...(mixedPitch !== undefined && { 'data-hwp-editor-mixed-pitch': 'true' }) },
              mixedPitch !== undefined ? { mixedPitchNode: node } : undefined));
            if (line.paperHeading) node.forEach((run, offset) => {
              if (run.isText) decorations.push(Decoration.inline(pos + 1 + offset, pos + 1 + offset + run.nodeSize, { 'data-hwp-body-run': 'true' }));
            });
            else decorations.push(Decoration.inline(pos + 1, pos + node.nodeSize - 1, { 'data-hwp-body-run': 'true' }));
            values.push(`${pos}:${encodeHwpBodyLine(line)}:${node.nodeSize}:${mixedPitch ?? ''}`);
          });
          const next = values.join(',');
          if (signature !== next || lastDocument !== view.state.doc) {
            signature = next; lastDocument = view.state.doc;
            view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, decorations)).setMeta('addToHistory', false));
          }
        });
      };
      const fontsChanged = () => { sourceColumns = new WeakMap(); schedule(); };
      document.fonts.addEventListener('loadingdone', fontsChanged);
      document.addEventListener(editorLayoutResume, schedule);
      schedule();
      return { update: schedule, destroy() {
        cancelAnimationFrame(frame);
        document.fonts.removeEventListener('loadingdone', fontsChanged);
        document.removeEventListener(editorLayoutResume, schedule);
      } };
    },
  })],
});
