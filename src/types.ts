import type { JSONContent } from "@tiptap/react";
/** Global DOCX background, carried with each imported section's page settings.
 * Cached RGB and a bounded VML fill preserve appearance across different themes. */
export interface WordBackground { color: string; display: boolean; drawingXml?: string; }
export interface PageSettings {
  size: "A4" | "Letter" | "Custom";
  /** Unrotated physical paper dimensions in millimetres. */
  width?: number;
  height?: number;
  orientation: "portrait" | "landscape";
  margin: number;
  margins?: { top: number; right: number; bottom: number; left: number };
  header: string;
  footer: string;
  pageNumbers: boolean;
  differentFirstPage?: boolean;
  differentOddEven?: boolean;
  firstHeader?: string;
  firstFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstPageNumbers?: boolean;
  evenPageNumbers?: boolean;
  headerDistance?: number;
  footerDistance?: number;
  gutter?: number;
  gutterTop?: boolean;
  breakType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  columns?: { count: number; space: number; widths?: number[]; balance?: boolean };
  /** Imported Hancom section policy; preserve the paragraphs in the model. */
  hideEmptyLines?: boolean;
  /** OOXML document compatibility: allow the final exact line's empty leading
   * below the body edge. Unset retains native Studio/HWP layout. */
  suppressBottomSpacing?: boolean;
  wordBackground?: WordBackground;
  headerHtml?: string;
  footerHtml?: string;
  firstHeaderHtml?: string;
  firstFooterHtml?: string;
  evenHeaderHtml?: string;
  evenFooterHtml?: string;
}
export interface Comment {
  id: string;
  quote: string;
  text: string;
  createdAt: string;
  resolved: boolean;
  author?: string;
  initials?: string;
  html?: string;
}
export interface StudioDocument {
  id?: string;
  format: "word-studio";
  version: 1;
  title: string;
  content: JSONContent;
  page: PageSettings;
  /** Document-wide Word auto hyphenation flag. Undefined keeps the DOCX element absent. */
  wordAutoHyphenation?: boolean;
  /** Preserve the authored Word trailing-space wrap option, including explicit off. */
  wordWrapTrailSpaces?: boolean;
  /** The Word default layout applies when the source omitted the compatibility option. */
  wordDefaultHangingSpaces?: true;
  comments: Comment[];
  updatedAt: string;
}
export const basePageMargins = (page: PageSettings) => page.margins ?? {
  top: page.margin, right: page.margin, bottom: page.margin, left: page.margin,
};
export function paperSize(page: PageSettings): [number, number] {
  return page.size === 'Custom' && Number.isFinite(page.width) && Number.isFinite(page.height)
    ? [page.width!, page.height!] : page.size === 'Letter' ? [215.9, 279.4] : [210, 297];
}
export function paperDimensions(page: PageSettings): [number, number] {
  const [width, height] = paperSize(page);
  return page.orientation === 'landscape' ? [height, width] : [width, height];
}
export const pageMargins = (page: PageSettings) => {
  const margins = { ...basePageMargins(page) };
  margins[page.gutterTop ? 'top' : 'left'] += page.gutter ?? 0;
  return margins;
};
const nonLiteralStory = '[data-page-number],[data-document-field],[data-note-reference],[data-docx-tab],[data-symbol-font]';
const hasLooseText = (container: Node) => [...container.childNodes].some(node =>
  node.nodeType === Node.TEXT_NODE && node.textContent?.trim());

function isLiteralStoryParagraph(element: Element | null, allowLeadingBreak = false) {
  if (element?.tagName !== 'P') return false;
  const inline = allowLeadingBreak ? 'span,b,strong,i,em,u,s,br' : 'span,b,strong,i,em,u,s';
  if ([...element.querySelectorAll('*')].some(child => !child.matches(inline))) return false;
  const breaks = element.querySelectorAll('br');
  if (breaks.length === 0) return true;
  if (breaks.length !== 1) return false;
  const prefix = document.createRange();
  prefix.selectNodeContents(element);
  prefix.setEndBefore(breaks[0]);
  return !prefix.toString().trim();
}

// The width must come from the three explicit cell tracks. The furniture
// renderer builds the same fixed grid in editor and print before measuring it.
function isLiteralStoryTable(element: Element | null, paperWidth: number) {
  if (!(element instanceof HTMLTableElement) || element.dataset.wordTableLayout !== 'fixed' ||
    element.hasAttribute('data-table-anchor') || element.hasAttribute('data-word-table-width-auto') ||
    element.rows.length !== 1 || element.querySelector('caption,table,colgroup') ||
    (element.style.marginLeft && parseFloat(element.style.marginLeft) !== 0)) return false;
  const cells = [...element.rows[0].cells];
  if (cells.length !== 3 || !cells.every(cell => cell.tagName === 'TD' && cell.rowSpan === 1 && cell.colSpan === 1 &&
    cell.childElementCount === 1 && !hasLooseText(cell) && isLiteralStoryParagraph(cell.firstElementChild, true))) return false;
  const widths = cells.map(cell => cell.getAttribute('colwidth') || '');
  if (widths.some(width => !/^\d+(?:\.\d+)?$/.test(width) || Number(width) <= 0)) return false;
  const total = widths.reduce((sum, width) => sum + Number(width), 0);
  const authoredWidth = element.style.width;
  if (authoredWidth && (!/^\d+(?:\.\d+)?px$/.test(authoredWidth) || Math.abs(parseFloat(authoredWidth) - total) > .5))
    return false;
  return total <= paperWidth;
}

// The complete-row probe can reuse the authored body height if repeated
// literal stories have been measured and fit in the original margins.
export function hasUniformTablePageBody(page: PageSettings, occupied?: { top: number; bottom: number }) {
  if (page.differentFirstPage || page.differentOddEven || page.pageNumbers) return false;
  const margins = pageMargins(page), px = 96 / 25.4;
  const stories = [[page.header, page.headerHtml, page.headerDistance], [page.footer, page.footerHtml, page.footerDistance]] as const;
  let hasLiteral = false;
  for (const [i, [text, html, distance]] of stories.entries()) {
    const story = document.createElement('template');
    story.innerHTML = html || '';
    if (story.content.querySelector(nonLiteralStory)) return false;
    // Keep the blank-story handling used by empty DOCX export paragraphs.
    if (!text?.trim() && !story.content.textContent?.trim() &&
        !story.content.querySelector('img,svg,table,[data-page-number]')) continue;
    const first = story.content.firstElementChild;
    const extent = i === 0 ? occupied?.top : occupied?.bottom;
    if (!(extent && extent > 0) || distance === undefined || page.breakType === 'continuous') return false;
    if (html) {
      const width = (paperDimensions(page)[0] - margins.left - margins.right) * px;
      if (story.content.childElementCount !== 1 || hasLooseText(story.content) || !first?.textContent?.trim() ||
        (!isLiteralStoryParagraph(first) && !isLiteralStoryTable(first, width))) return false;
    } else if (/[\r\n]/.test(text)) return false;
    hasLiteral = true;
  }
  if (!hasLiteral) return true;
  return occupied!.top <= margins.top * px && occupied!.bottom <= margins.bottom * px;
}
export function pageFurniture(page: PageSettings, number: number, first = number === 1) {
  const prefix = first && page.differentFirstPage ? 'first' : number % 2 === 0 && page.differentOddEven ? 'even' : null;
  return prefix ? { header: page[`${prefix}Header`] ?? '', footer: page[`${prefix}Footer`] ?? '', pageNumbers: page[`${prefix}PageNumbers`] ?? false,
    headerHtml: page[`${prefix}HeaderHtml`], footerHtml: page[`${prefix}FooterHtml`] }
    : { header: page.header, footer: page.footer, pageNumbers: page.pageNumbers, headerHtml: page.headerHtml, footerHtml: page.footerHtml };
}
export const defaultPage: PageSettings = {
  size: "A4",
  orientation: "portrait",
  margin: 25.4,
  header: "",
  footer: "",
  pageNumbers: true,
};
