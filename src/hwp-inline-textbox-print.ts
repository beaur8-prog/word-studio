import { horizontalInkFits, lineOpportunities, sourceLineBoundaries } from './hwp-source-cell-lines';
import { hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } from './hwp-text-advances';
import { paragraphLineOffsets } from './paragraph-fragments';
import { textBoxFrame } from './textbox-frame';

const precision = .05 + 1 / 15 + 2 / 64;

/** Project saved source wraps only within the current, unchanged text box.
 * The live model and standardized DOCX keep ordinary editable paragraphs. */
export function layoutHwpInlineTextBoxLines(box: HTMLElement) {
  const frame = textBoxFrame(box.dataset.documentTextbox);
  const content = box.querySelector<HTMLElement>(':scope > [data-textbox-content]');
  if (!frame || frame.position !== 'inline' || frame.shape !== 'rect' || frame.align !== 'center' ||
      !content || content.dataset.hwpPrintTextbox === 'true' || !content.querySelector('[data-hwp-source-wrap]')) return;
  const bounds = content.getBoundingClientRect(), style = getComputedStyle(content);
  const borderLeft = parseFloat(style.borderLeftWidth), borderRight = parseFloat(style.borderRightWidth);
  const horizontal = borderLeft + borderRight;
  const range = document.createRange();
  const edits: { paragraph: HTMLParagraphElement; original: Node[]; content: DocumentFragment; ends: number[]; bands: HTMLElement[] }[] = [];
  for (const paragraph of content.children) {
    if (!(paragraph instanceof HTMLParagraphElement) || paragraph.querySelector(':not(span),[id],[data-document-field]')) return;
    const text = paragraph.textContent ?? '';
    if (!text) continue;
    const proof = parseHwpSourceWrap(paragraph.dataset.hwpSourceWrap), source = proof?.inlineTextBox;
    if (!proof || !source || Math.abs(frame.width - source.width) > precision ||
        Math.abs(frame.height - source.height) > precision || hwpWrapFingerprint(text) !== proof.fingerprint ||
        normalizeHwpWrapText(text).length !== proof.ends.at(-1)) return;
    const css = getComputedStyle(paragraph), para = paragraph.getBoundingClientRect();
    if (css.transform !== 'none' || css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' || css.textAlign !== 'justify' ||
        Math.abs((parseFloat(css.textIndent) || 0) + source.hanging) > precision ||
        Math.abs(para.width + source.hanging + horizontal - source.textWidth) > .24 ||
        Math.abs(parseFloat(css.lineHeight) - source.pitch) > precision) return;
    if ([paragraph, ...paragraph.querySelectorAll('span')].some(run => {
      const typography = getComputedStyle(run);
      return typography.fontFamily.replace(/^["']|["']$/g, '') !== source.family ||
        Math.abs(parseFloat(typography.fontSize) - source.size) > .07 || typography.fontStyle !== 'normal' ||
        typography.fontWeight !== '400' || typography.textDecorationLine !== 'none' || typography.textShadow !== 'none';
    })) return;
    if (proof.ends.length === 1) {
      if (paragraphLineOffsets(paragraph).length) return;
      continue;
    }
    const cuts = sourceLineBoundaries(paragraph, proof.ends.slice(0, -1));
    if (!cuts) return;
    const lines = lineOpportunities(paragraph, cuts, range);
    for (const boundary of lines.querySelectorAll('wbr')) boundary.remove();
    const bands = [...lines.children] as HTMLSpanElement[];
    for (const [index, band] of bands.entries()) {
      band.style.display = 'block';
      band.style.textIndent = index ? '0px' : css.textIndent;
      band.style.textAlignLast = index === bands.length - 1 ? 'start' : 'justify';
      band.style.minHeight = source.pitch + 'px';
      band.style.setProperty('line-height', source.pitch + 'px', 'important');
      for (const run of [band, ...band.querySelectorAll('span')]) run.style.whiteSpace = 'pre-wrap';
    }
    edits.push({ paragraph, original: [...paragraph.childNodes], content: lines,
      ends: cuts.slice(1, -1).map(cut => cut.offset), bands });
  }
  if (!edits.length) return;
  for (const edit of edits) edit.paragraph.replaceChildren(edit.content);
  const fits = () => {
    const after = content.getBoundingClientRect();
    const inkBounds = { left: after.left + borderLeft, right: after.right - borderRight,
      top: after.top + parseFloat(style.borderTopWidth), bottom: after.bottom - parseFloat(style.borderBottomWidth) };
    return after.width === bounds.width && Math.abs(after.height - frame.height) <= precision &&
      edits.every(({ paragraph, ends, original }) => {
        const actual = paragraphLineOffsets(paragraph);
        if (actual.length !== ends.length || actual.some((end, index) => end !== ends[index]) ||
            paragraph.textContent !== original.map(node => node.textContent).join('')) return false;
        range.selectNodeContents(paragraph);
        const glyphs = [...range.getClientRects()].filter(rect => rect.width > 0 && rect.height > 0);
        return glyphs.length > 0 && glyphs.every(rect => rect.top >= inkBounds.top && rect.bottom <= inkBounds.bottom) &&
          horizontalInkFits(paragraph, range, inkBounds);
      });
  };
  if (!fits()) {
    // Each band already records an authored line. A substitute face can
    // trigger an extra wrap even when all of its ink fits within the outline.
    for (const { bands } of edits)
      for (const band of bands)
        for (const run of [band, ...band.querySelectorAll('span')]) run.style.whiteSpace = 'pre';
    if (!fits()) {
      for (const edit of edits) edit.paragraph.replaceChildren(...edit.original);
      return;
    }
  }
  content.dataset.hwpPrintTextbox = 'true';
}
