import { parseHwpBodyLine } from './hwp-body-line';
import { measuredHwpLineInk } from './hwp-line-ink';
import { hwpWrapFingerprint } from './hwp-text-advances';

const unsupported = 'a,br,img,svg,table,mark,u,s,del,sup,sub,[data-flow-widget],[data-note-reference],'
  + '[data-page-number],[data-docx-tab],[data-docx-bookmark],[data-word-overlap],[data-symbol-font]';
const inkOffset = '--hwp-editor-body-end-shift';

function ordinaryLine(paragraph: HTMLParagraphElement) {
  const css = getComputedStyle(paragraph);
  return css.writingMode === 'horizontal-tb' && css.direction === 'ltr' && css.transform === 'none' &&
    css.backgroundColor === 'rgba(0, 0, 0, 0)' && css.textShadow === 'none' &&
    !parseFloat(css.borderTopWidth) && !parseFloat(css.borderBottomWidth) &&
    !paragraph.closest('td,th,li,[data-column-count]:not([data-column-count="1"])') &&
    !paragraph.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') &&
    !paragraph.querySelector(unsupported);
}

function relativeRunBoxes(paragraph: HTMLParagraphElement, scale: number, shift = 0) {
  const frame = paragraph.getBoundingClientRect(), range = document.createRange();
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT), boxes: number[][] = [];
  while (walker.nextNode()) {
    const text = walker.currentNode as Text;
    if (!text.data.trim()) continue;
    range.selectNodeContents(text);
    const rects = [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
    if (rects.length !== 1) return;
    const [box] = rects;
    boxes.push([(box.left - frame.left) / scale, (box.top - frame.top) / scale - shift, box.width / scale, box.height / scale]);
  }
  return boxes.length ? boxes : undefined;
}

function cloneLine(paragraph: HTMLParagraphElement, host: HTMLElement) {
  const css = getComputedStyle(paragraph);
  const clone = paragraph.cloneNode(true) as HTMLParagraphElement;
  clone.style.removeProperty(inkOffset);
  clone.style.width = css.width;
  clone.style.margin = '0';
  for (const property of ['fontFamily', 'fontSize', 'fontStyle', 'fontWeight', 'letterSpacing', 'wordSpacing',
    'whiteSpace', 'textAlign', 'textIndent', 'lineHeight', 'minHeight', 'direction', 'writingMode'] as const)
    clone.style[property] = css[property];
  clone.lang = paragraph.closest<HTMLElement>('[lang]')?.lang || document.documentElement.lang;
  host.append(clone);
  return clone;
}

function provenInk(paragraph: HTMLParagraphElement, clone: HTMLParagraphElement, scale: number, maxFont: number) {
  const active = relativeRunBoxes(paragraph, scale, parseFloat(paragraph.style.getPropertyValue(inkOffset)) || 0);
  const sample = relativeRunBoxes(clone, 1);
  const box = clone.getBoundingClientRect(), actual = paragraph.getBoundingClientRect();
  if (!active || !sample || active.length !== sample.length || Math.abs(box.height - actual.height / scale) > 1 / 32 ||
      Math.abs(box.width - actual.width / scale) > 1 / 32 ||
      sample.some((run, i) => run.some((edge, j) => Math.abs(edge - active[i][j]) > 1 / 16))) return;
  const ink = measuredHwpLineInk(clone, 1, maxFont);
  return ink && { top: ink.top - box.top, bottom: ink.bottom - box.top, baseline: ink.baseline - box.top };
}

/** The pager provides the natural room after removing old view-only gaps.
 * Probe the same laid-out runs outside ProseMirror before admitting paint
 * on a sheet where the full paragraph box would not fit. */
export function editorSourceBodyEnd(
  paragraph: HTMLParagraphElement,
  previous: HTMLParagraphElement,
  scale: number,
  room: number,
) {
  const policy = parseHwpBodyLine(paragraph.dataset.hwpBodyLine), source = policy?.bodyEndFit;
  if (!policy || !source || paragraph.dataset.hwpBodyLayout !== 'mixed-top' || !(scale > 0) || previous === paragraph ||
      paragraph.parentElement !== previous.parentElement || !(room > 0 && room < policy.base) ||
      Math.abs(room - source.remainder) > .25 || !ordinaryLine(paragraph) || !ordinaryLine(previous) ||
      hwpWrapFingerprint(paragraph.textContent ?? '') !== source.fingerprint ||
      hwpWrapFingerprint(previous.textContent ?? '') !== source.preceding) return;
  let sibling = paragraph.previousElementSibling;
  if (sibling?.matches('.ProseMirror-widget.pagination-gap')) sibling = sibling.previousElementSibling;
  if (sibling !== previous) return;
  const box = paragraph.getBoundingClientRect(), prior = previous.getBoundingClientRect();
  const indent = parseFloat(getComputedStyle(paragraph).textIndent) || 0;
  if (Math.abs(box.width / scale - Math.min(0, indent) - source.width) > .15 ||
      Math.abs(box.height / scale - policy.base) > .1 || !paragraph.querySelector('[data-hwp-body-run]')) return;

  const host = document.createElement('div');
  host.className = 'tiptap';
  host.setAttribute('aria-hidden', 'true');
  host.style.cssText = 'position:fixed;top:0;left:-100000px;visibility:hidden;pointer-events:none;min-height:0;padding:0';
  const pre = cloneLine(previous, host), tail = cloneLine(paragraph, host);
  document.body.append(host);
  try {
    const limit = policy.maxFont ?? policy.font;
    const preceding = provenInk(previous, pre, scale, limit), ink = provenInk(paragraph, tail, scale, limit);
    if (!preceding || !ink) return;
    const shift = source.baseline - ink.baseline;
    const bottom = ink.bottom + shift, gap = prior.height / scale - preceding.bottom + ink.top + shift;
    if (shift < -.05 && shift >= -4 && bottom <= room && gap >= .5)
      return { bottom, style: `${inkOffset}:${shift}px` };
  } finally { host.remove(); }
}
