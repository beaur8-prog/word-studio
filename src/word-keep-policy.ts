import {keepShortWordParagraphTogether, paragraphLineOffsets, paragraphPrefixHeight} from './paragraph-fragments';
import { imageFloating } from './image-floating';
import { tableAnchor } from './table-anchor-data';
import { textBoxFrame } from './textbox-frame';

/** Ordinary Word story paragraphs only. Tables, columns, fields, notes and
 * tabbed stories retain their existing layout until they share this budget. */
export function wordKeepParagraph(element: HTMLElement | null): element is HTMLElement {
  return !!element?.matches('p[data-word-paragraph-style],h1[data-word-paragraph-style],h2[data-word-paragraph-style],h3[data-word-paragraph-style],h4[data-word-paragraph-style],h5[data-word-paragraph-style],h6[data-word-paragraph-style]') &&
    !element.closest('li,td,th,[data-document-textbox],[data-textbox-content],[data-column-count]:not([data-column-count="1"])') &&
    !element.textContent?.includes('\t') &&
    !element.querySelector('img:not(.ProseMirror-separator),svg,table,[data-note-reference],[data-page-number],[data-flow-widget],[data-docx-tab]');
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
export function wordWidowSplitBounds(paragraph: HTMLElement, lineCount: number, bodyHeight: number, scale = 1) {
  if (paragraph.tagName !== 'P' || !wordKeepParagraph(paragraph) ||
      paragraph.getBoundingClientRect().height / scale > bodyHeight) return;
  const orphans = Math.max(1, Number(paragraph.style.orphans) || 1);
  const widows = Math.max(1, Number(paragraph.style.widows) || 1);
  if (orphans > 1 || widows > 1) return { minimumHead: orphans, maximumHead: lineCount - widows };
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
