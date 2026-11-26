import { characterScaleProperty, parseCharacterScale } from './character-scale';
import { measuredHwpInkAtLineOffsets } from './hwp-line-ink';
import { hwpNestedCellFaceAttribute, parseHwpNestedCellFace } from './hwp-nested-table-layout';
import { horizontalInkFits } from './hwp-source-cell-lines';
import { hwpSingleBodyFace, hwpWrapFingerprint, normalizeHwpWrapText } from './hwp-text-advances';
import { paragraphLineOffsets } from './paragraph-fragments';

const nestedFaceFallback = ', "Noto Sans KR"';
const measureFaceAttribute = 'data-hwp-nested-face-for-measure';

const familyMatches = (value: string) => value.replace(/^['"]|['"]$/g, '') === hwpSingleBodyFace;

function readFace(value: string | null) {
  if (!value || value.length > 500) return;
  try { return parseHwpNestedCellFace(JSON.parse(value)); } catch { return; }
}

/** Alternate faces must recover the native natural breaks and visible ink
 * without changing the stored font, em size, scale or paragraph pitch. */
export function layoutHwpNestedCellFaces(root: HTMLTableElement, trackProjection = false) {
  const range = document.createRange();
  for (const paragraph of root.querySelectorAll<HTMLParagraphElement>(`p[${hwpNestedCellFaceAttribute}]`)) {
    const cell = paragraph.parentElement, row = cell?.parentElement;
    const proof = readFace(paragraph.getAttribute(hwpNestedCellFaceAttribute));
    if (!proof || !(cell instanceof HTMLTableCellElement) || !(row instanceof HTMLTableRowElement) ||
        cell.children.length !== 1 || cell.rowSpan !== 1 || cell.colSpan !== 1 || row.dataset.wordRowRule === 'exact' ||
        paragraph.querySelector(':not(span),[id],[data-docx-bookmark],[data-word-overlap]')) continue;
    const text = paragraph.textContent ?? '';
    if (!/^[가-힣 ]{8,40}$/u.test(text) || hwpWrapFingerprint(text) !== proof.fingerprint ||
        normalizeHwpWrapText(text).length !== proof.ends.at(-1) || paragraphLineOffsets(paragraph).length + 1 <= proof.ends.length) continue;
    const prior = paragraph.getBoundingClientRect(), frame = cell.getBoundingClientRect(), rowBefore = row.getBoundingClientRect();
    const css = getComputedStyle(paragraph), cellCss = getComputedStyle(cell);
    if (cellCss.verticalAlign !== 'middle' || css.textAlign !== 'center' || css.direction !== 'ltr' ||
        css.writingMode !== 'horizontal-tb' || Math.abs(frame.width - proof.width) > .5 ||
        Math.abs(parseFloat(css.lineHeight) - proof.pitch) > .12 || !familyMatches(css.fontFamily)) continue;
    const matrix = new DOMMatrixReadOnly(css.transform === 'none' ? undefined : css.transform);
    const scale = proof.scale / 100;
    if (Math.abs(matrix.a - scale) > .002 || matrix.d !== 1 || matrix.b || matrix.c || matrix.e || matrix.f) continue;
    const leaves = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    const styled = new Set<HTMLElement>();
    if (familyMatches(paragraph.style.fontFamily)) styled.add(paragraph);
    let valid = true;
    while (leaves.nextNode()) {
      const node = leaves.currentNode as Text;
      if (!node.data.trim()) continue;
      const inline = node.parentElement!, style = getComputedStyle(inline);
      if (!familyMatches(style.fontFamily) || Math.abs(parseFloat(style.fontSize) - proof.size) > .07 ||
          style.fontWeight !== '400' || style.fontStyle !== 'normal' ||
          parseCharacterScale(inline.style.getPropertyValue(characterScaleProperty)) !== proof.scale) valid = false;
      if (familyMatches(inline.style.fontFamily)) styled.add(inline);
    }
    if (!valid || !styled.size) continue;
    const previous = [...styled].map(node => [node, node.style.fontFamily] as const);
    for (const [node] of previous) node.style.fontFamily += nestedFaceFallback;
    const breaks = paragraphLineOffsets(paragraph), current = paragraph.getBoundingClientRect(), bounds = cell.getBoundingClientRect();
    const endsMatch = breaks.length + 1 === proof.ends.length &&
      breaks.every((offset, index) => normalizeHwpWrapText(text.slice(0, offset)).length === proof.ends[index]);
    const closeHeight = current.height < prior.height && row.getBoundingClientRect().height < rowBefore.height &&
      Math.abs(row.getBoundingClientRect().height - proof.height) <= .5;
    const ink = endsMatch && closeHeight && current.width === prior.width && bounds.width === frame.width &&
      horizontalInkFits(paragraph, range, bounds, scale) ? measuredHwpInkAtLineOffsets(paragraph, breaks, proof.size) : undefined;
    if (!ink || ink.length !== proof.ends.length || ink.some((line, i) => line.top < bounds.top || line.bottom > bounds.bottom ||
        (i > 0 && line.top < ink[i - 1].bottom))) {
      for (const [node, original] of previous) node.style.fontFamily = original;
    } else {
      paragraph.dataset.hwpNestedFaceApplied = 'true';
      if (trackProjection)
        for (const node of styled) node.setAttribute(measureFaceAttribute, '');
    }
  }
}

/** Remove the print measure's face before native editable output. */
export function clearHwpNestedCellFaceProjection(root: HTMLElement) {
  for (const node of root.querySelectorAll<HTMLElement>(`[${measureFaceAttribute}]`)) {
    node.removeAttribute(measureFaceAttribute);
    const family = node.style.fontFamily;
    if (!family.endsWith(nestedFaceFallback)) continue;
    const source = family.slice(0, -nestedFaceFallback.length);
    if (familyMatches(source)) node.style.fontFamily = source;
  }
  root.querySelectorAll('p[data-hwp-nested-face-applied]')
    .forEach(paragraph => paragraph.removeAttribute('data-hwp-nested-face-applied'));
}
