import { keepShortWordParagraphTogether, paragraphLineOffsets, paragraphPrefixHeight, paragraphTextOffset, splitText } from './paragraph-fragments';
import { parseTabStops } from './document-tabs';
import { imageFloating } from './image-floating';
import { tableAnchor } from './table-anchor-data';
import { textBoxFrame } from './textbox-frame';

function wordBodyTextParagraph(element: HTMLElement | null): element is HTMLElement {
  return !!element?.matches('p[data-word-paragraph-style],h1[data-word-paragraph-style],h2[data-word-paragraph-style],h3[data-word-paragraph-style],h4[data-word-paragraph-style],h5[data-word-paragraph-style],h6[data-word-paragraph-style]') &&
    !element.closest('li,td,th,[data-document-textbox],[data-textbox-content],[data-column-count]:not([data-column-count="1"])') &&
    !element.querySelector('img:not(.ProseMirror-separator),svg,table,[data-note-reference],[data-page-number],[data-flow-widget]');
}

/** The existing keep chain rules apply to body paragraphs without tabs. */
export function wordKeepParagraph(element: HTMLElement | null): element is HTMLElement {
  return wordBodyTextParagraph(element) && !element.textContent?.includes('\t') && !element.querySelector('[data-docx-tab]');
}

/** A left tab is independent of the width of the following field. A cut can
 * reuse its geometry after the first line has consumed both tab and text. */
function autoLeftTabCanSplit(paragraph: HTMLElement, firstLineEnd: () => number | undefined) {
  if (paragraph.dataset.wordLineRule !== 'auto' ||
      paragraph.matches('[data-word-outline-level]:not([data-word-outline-level="9"]),[data-hwp-body-line],[data-hwp-cell-line]') ||
      paragraph.closest('blockquote,.document-furniture,[data-document-note]') ||
      paragraph.querySelector('br:not(.ProseMirror-trailingBreak),[data-document-field],[data-symbol-font],[data-word-overlap],sup,sub')) return false;
  const tabs = paragraph.querySelectorAll<HTMLElement>('[data-docx-tab]');
  if (tabs.length !== 1 || tabs[0].textContent !== '\t' || paragraph.textContent?.split('\t').length !== 2 ||
      parseTabStops(paragraph.dataset.tabStops).some(stop => stop.type !== 'left')) return false;
  const style = getComputedStyle(paragraph);
  if (style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' || !['left', 'start'].includes(style.textAlign)) return false;
  const cut = firstLineEnd();
  const tab = tabs[0];
  if (!cut || paragraphTextOffset(paragraph, tab, tab.childNodes.length) >= cut) return false;
  const [first] = splitText(paragraph, cut);
  return !!first.textContent?.split('\t')[1]?.trim();
}

/** Later wrapping anchors can change the width of an earlier continuation
 * during pagination. Verify the complete story before enforcing line counts. */
export function wordFlowHasWrappingAnchors(root: HTMLElement) {
  const anchors = root.querySelectorAll<HTMLElement>('img[data-image-floating],[data-textbox-position="floating"],table[data-table-anchor]');
  return Array.from(anchors).some(anchor => {
    if (anchor.closest('.document-furniture,.section-papers,.paper-backgrounds,.pagination-headers,[data-document-note]')) return false;
    if (anchor instanceof HTMLImageElement) {
      const image = imageFloating(anchor.dataset.imageFloating);
      if (!image) return false;
      const { drawingGroup } = image;
      return drawingGroup ? !drawingGroup.behind && drawingGroup.wrap.type !== 0 : !image.behindDocument && !!image.wrap?.type;
    }
    if (anchor instanceof HTMLTableElement) return !!tableAnchor(anchor.dataset.tableAnchor)?.wrap;
    return textBoxFrame(anchor.dataset.documentTextbox)?.wrap === 'square';
  });
}

/** Legal lines at either side of a split in a fitting Word body paragraph.
 * Keep the existing fallback for paragraphs taller than a fresh page. */
export function wordWidowSplitBounds(paragraph: HTMLElement, lineCount: number, bodyHeight: number, scale = 1,
  firstLineEnd?: () => number | undefined) {
  if (paragraph.tagName !== 'P' || !wordBodyTextParagraph(paragraph) ||
      paragraph.getBoundingClientRect().height / scale > bodyHeight) return;
  const orphans = Math.max(1, Number(paragraph.style.orphans) || 1);
  const widows = Math.max(1, Number(paragraph.style.widows) || 1);
  if (orphans === 1 && widows === 1) return;
  if ((paragraph.textContent?.includes('\t') || paragraph.querySelector('[data-docx-tab]')) &&
      (!firstLineEnd || !autoLeftTabCanSplit(paragraph, firstLineEnd))) return;
  return { minimumHead: orphans, maximumHead: lineCount - widows };
}

/** Keep-with-next protects the first legal prefix, not the entire next long
 * paragraph. A fitting keep-lines paragraph is the exception. */
export function wordKeepPrefixHeight(paragraph: HTMLElement, bodyHeight: number, scale = 1) {
  const height = paragraph.getBoundingClientRect().height / scale;
  const offsets = paragraphLineOffsets(paragraph);
  if (!offsets.length || paragraph.dataset.wordKeepLines === 'true' ||
      keepShortWordParagraphTogether(paragraph, offsets.length + 1, bodyHeight, scale)) return height;
  const orphans = Math.max(1, Number(paragraph.style.orphans) || 1);
  const cut = offsets[orphans - 1];
  return cut === undefined ? height : paragraphPrefixHeight(paragraph, cut, scale);
}

/** Capture siblings before print moves any of them into a page. Hard breaks
 * and story boundaries terminate chains. Oversized chains are handled as one
 * decision, so their prefixes cannot repeatedly move to new blank pages. */
export function wordKeepChains(root: HTMLElement) {
  const chains = new WeakMap<HTMLElement, HTMLElement[]>();
  const containers = [root, ...root.querySelectorAll<HTMLElement>('[data-section-content],[data-column-count="1"]')];
  for (const container of containers) {
    const children = [...container.children] as HTMLElement[];
    for (let index = 0; index < children.length; index++) {
      const first = children[index];
      if (!wordKeepParagraph(first) || first.dataset.wordKeepNext !== 'true') continue;
      const group = [first]; let cursor = index;
      while (children[cursor].dataset.wordKeepNext === 'true' && wordKeepParagraph(children[cursor + 1])) group.push(children[++cursor]);
      if (group.length > 1) { chains.set(first, group); index = cursor - 1; }
    }
  }
  return chains;
}
