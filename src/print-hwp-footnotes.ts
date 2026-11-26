import { cssPixels } from './hwp-cell-line';
import { hwpFootnoteLayoutAttribute, parseHwpFootnoteLayout } from './hwp-footnote-layout';
import { measuredHwpLineInk } from './hwp-line-ink';
import { paragraphLineOffsets, splitText } from './paragraph-fragments';

function textBottom(paragraph: HTMLElement) {
  const range = document.createRange();
  range.selectNodeContents(paragraph);
  return Math.max(-Infinity, ...[...range.getClientRects()].filter(box => box.width > 0 && box.height > 0).map(box => box.bottom));
}

/** Preserve the source's terminal height when the visible last line fits it. */
function unpaintedFootnoteLeading(paragraph: HTMLParagraphElement, font: number, pitch: number) {
  if (pitch <= font || paragraph.textContent.length > 1500) return 0;
  const last = paragraphLineOffsets(paragraph).at(-1) ?? 0;
  const [, line] = splitText(paragraph, last);
  if (last) line.removeAttribute('data-hwp-footnote-label');
  Object.assign(line.style, { position: 'absolute', left: '-100000px', top: '0', width: 'max-content',
    visibility: 'hidden', whiteSpace: 'nowrap', textIndent: '0' });
  const lastBox = textBottom(paragraph);
  let lastInk = lastBox;
  paragraph.after(line);
  try {
    const reported = textBottom(line), ink = measuredHwpLineInk(line, 1, font);
    if (ink && Number.isFinite(reported)) lastInk = Math.min(lastBox, lastBox + ink.bottom - reported);
  } finally { line.remove(); }
  return Number.isFinite(lastInk) ? Math.max(0, Math.min(pitch - font, paragraph.getBoundingClientRect().bottom - lastInk - .25)) : 0;
}

/** Apply native note spacing to compatible plain footnotes on the print copy. */
export function layoutHwpFootnotes(box: HTMLElement) {
  delete box.dataset.hwpFootnoteLayoutActive;
  for (const name of ['--hwp-note-spacing-below', '--hwp-note-separator-width', '--hwp-note-separator-stroke'])
    box.style.removeProperty(name);
  if (!box.children.length) return;
  const notes = [...box.children].map(child => {
    if (!(child instanceof HTMLElement)) return;
    const layout = parseHwpFootnoteLayout(child.getAttribute(hwpFootnoteLayoutAttribute));
    const paragraph = child.firstElementChild;
    if (!layout || child.children.length !== 1 || !(paragraph instanceof HTMLParagraphElement) ||
        !paragraph.textContent.trim() || paragraph.querySelector('br,img,svg,table,sup,sub,a,[data-flow-widget],' +
          '[data-note-reference],[data-document-field],[data-docx-tab]') ||
        (cssPixels(paragraph.style.marginTop) ?? 0) !== 0 ||
        (cssPixels(paragraph.style.marginBottom) ?? 0) !== 0) return;
    const font = cssPixels(paragraph.style.fontSize), pitch = cssPixels(paragraph.style.lineHeight);
    if (!font || !pitch || pitch < font || !paragraph.style.fontFamily) return;
    return { child, paragraph, font, pitch, layout };
  });
  if (notes.some(note => !note)) return;
  const complete = notes.filter(note => note !== undefined);
  const format = complete[0].layout;
  if (complete.some(({ layout }) => layout.above !== format.above || layout.below !== format.below ||
      layout.between !== format.between || layout.separator !== format.separator || layout.stroke !== format.stroke)) return;
  box.dataset.hwpFootnoteLayoutActive = 'true';
  box.style.setProperty('--hwp-note-spacing-below', `${format.below}px`);
  box.style.setProperty('--hwp-note-separator-width', `${format.separator}px`);
  box.style.setProperty('--hwp-note-separator-stroke', `${format.stroke}px`);
  for (const [index, { child, paragraph, font, pitch, layout }] of complete.entries()) {
    child.style.font = `${font}px/${pitch}px ${paragraph.style.fontFamily}`;
    child.style.marginTop = index ? `${layout.between}px` : '0px';
    paragraph.dataset.hwpFootnoteLabel = layout.mark + ' ';
    const trim = unpaintedFootnoteLeading(paragraph, font, pitch);
    if (trim > 0) child.style.marginBottom = `-${trim}px`;
  }
  return format.above;
}
