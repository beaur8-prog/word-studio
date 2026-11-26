import { documentScale } from './document-scale';
import { hwpFootnoteLayoutAttribute, hwpFootnoteMarkAttribute, hwpFootnoteMark, parseHwpFootnoteLayout } from './hwp-footnote-layout';
import { measuredHwpLineInk } from './hwp-line-ink';
import { horizontalInkFits } from './hwp-source-cell-lines';
import { hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } from './hwp-text-advances';
import { paragraphLineOffsets, splitText } from './paragraph-fragments';

const noteSelector = '[data-note-reference]';
const unrelatedObjects = 'br,a,img,svg,table,mark,u,s,del,[id],[data-docx-tab],[data-docx-bookmark],[data-word-overlap],[data-page-number]';
const references = (element: HTMLElement) => [...element.querySelectorAll<HTMLElement>(noteSelector)];
const fingerprint = (text: string) => hwpWrapFingerprint(normalizeHwpWrapText(text));

function textRects(element: HTMLElement, range: Range) {
  range.selectNodeContents(element);
  return [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
}

/** Leave editable line flow untouched. For a source-proved terminal pair,
 * fit the existing two lines by their painted extents above the same notes. */
export function sourceFootnoteEndBottom(paragraph: HTMLElement, body: HTMLElement, notes: HTMLElement): number | undefined {
  if (!paragraph.getAttribute(hwpSourceWrapAttribute)?.includes('"noteTerminal"') ||
      !(paragraph instanceof HTMLParagraphElement) || paragraph.parentElement !== body ||
      paragraph.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') ||
      paragraph.querySelector(unrelatedObjects) || notes.dataset.hwpFootnoteLayoutActive !== 'true') return;
  const proof = parseHwpSourceWrap(paragraph.getAttribute(hwpSourceWrapAttribute)), source = proof?.noteTerminal;
  const nodes = references(paragraph), blank = paragraph.previousElementSibling, prior = blank?.previousElementSibling;
  if (!proof || !source || nodes.length !== 2 || notes.children.length !== 2 ||
      !(blank instanceof HTMLParagraphElement) || blank.textContent.trim() ||
      blank.querySelector('img,svg,table,' + noteSelector) || !(prior instanceof HTMLParagraphElement) ||
      prior.querySelector(unrelatedObjects + ',' + noteSelector) || fingerprint(prior.textContent) !== source.prior) return;
  for (const [i, ref] of nodes.entries()) {
    const story = notes.children[i] as HTMLElement, native = source.notes[i];
    const mark = hwpFootnoteMark(ref.getAttribute(hwpFootnoteMarkAttribute));
    const layout = parseHwpFootnoteLayout(story.getAttribute(hwpFootnoteLayoutAttribute));
    if (ref.dataset.noteReference !== story.dataset.documentNote || mark !== native.mark || layout?.mark !== native.mark ||
        fingerprint(story.textContent) !== native.fingerprint) return;
  }
  const literal = paragraph.cloneNode(true) as HTMLElement;
  references(literal).forEach(ref => ref.remove());
  const text = normalizeHwpWrapText(literal.textContent);
  if (text.length !== proof.ends[1] || hwpWrapFingerprint(text) !== proof.fingerprint) return;
  const offsets = paragraphLineOffsets(paragraph);
  if (offsets.length !== 1 || paragraphLineOffsets(prior).length) return;
  const style = getComputedStyle(paragraph), scale = documentScale(paragraph), bounds = body.getBoundingClientRect();
  const original = paragraph.getBoundingClientRect(), footnotes = notes.getBoundingClientRect();
  if (!(scale > 0) || Math.abs(Number(body.dataset.bodyHeight) - source.fullBodyHeight) > .3 ||
      Math.abs(bounds.width / scale - source.width) > .2 ||
      Math.abs(original.width / scale + source.hanging - source.width) > .2 ||
      Math.abs(parseFloat(style.textIndent) + source.hanging) > .12 ||
      Math.abs((original.top - bounds.top) / scale - source.lineTop) > source.font / 5 ||
      Math.abs(original.height / scale - source.pitch * 2) > .35 || bounds.bottom >= footnotes.top ||
      [style, getComputedStyle(prior)].some(css => css.transform !== 'none' || css.backgroundColor !== 'rgba(0, 0, 0, 0)' ||
        css.textShadow !== 'none' || parseFloat(css.borderTopWidth) || parseFloat(css.borderBottomWidth))) return;
  const [first, second] = splitText(paragraph, offsets[0]);
  const firstRefs = references(first);
  if (firstRefs.length !== nodes.length || firstRefs.some((ref, i) => ref.dataset.noteReference !== nodes[i].dataset.noteReference) ||
      references(second).length || first.textContent + second.textContent !== paragraph.textContent) return;
  const range = document.createRange(), priorRects = textRects(prior, range);
  if (!priorRects.length) return;
  const priorBottom = Math.max(bounds.top, ...priorRects.map(rect => rect.bottom));
  paragraph.replaceWith(first, second);
  try {
    const head = first.getBoundingClientRect(), tail = second.getBoundingClientRect();
    if (Math.abs(head.top - original.top) > .06 * scale || Math.abs(tail.bottom - original.bottom) > .06 * scale ||
        Math.abs((tail.top - head.top) / scale - source.pitch) > .2 ||
        !horizontalInkFits(first, range, bounds, scale) || !horizontalInkFits(second, range, bounds, scale)) return;
    const ink = measuredHwpLineInk(second, scale, source.font, true), headInk = textRects(first, range);
    if (!ink || !headInk.length) return;
    const delta = bounds.top + source.secondBaseline * scale - ink.baseline;
    const firstTop = Math.min(...headInk.map(rect => rect.top)) + delta;
    const firstBottom = Math.max(...headInk.map(rect => rect.bottom)) + delta;
    const bottom = ink.bottom + delta;
    if (!(delta < -.05 * scale && delta > -source.font * scale && firstTop >= priorBottom + .5 * scale &&
        firstBottom + .5 * scale <= ink.top + delta && bottom <= bounds.bottom)) return;
    paragraph.style.transform = `translateY(${delta / scale}px)`;
    paragraph.dataset.hwpProjectedBodyEnd = 'true';
    return bottom;
  } finally {
    first.replaceWith(paragraph);
    second.remove();
  }
}
