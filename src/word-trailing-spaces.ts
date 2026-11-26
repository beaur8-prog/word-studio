import { wordContentLanguage, type StudioDocument } from './types';
import { parseTabStops, parseWordNumberingTabs } from './document-tabs';
import { cssPixels } from './hwp-cell-line';
import { isWordGroupedRectTextBox, textBoxFrame } from './textbox-frame';

export const wordDefaultHangingSpacesProperty = 'WordStudio.DefaultTrailingSpaces';

export type WordLayoutOptions = {
  wordHangingSpaces?: boolean;
  wordBodyStrut?: boolean;
  language?: string;
};

export function wordLayoutOptions(doc: StudioDocument): WordLayoutOptions {
  return {
    wordHangingSpaces: usesWordHangingSpaces(doc),
    language: wordContentLanguage(doc.wordDefaultLanguage),
    // The explicit OOXML flag models one whitespace behavior. Either setting
    // still carries Word's layout policy, as does the recovered Word default.
    wordBodyStrut: doc.wordWrapTrailSpaces !== undefined || doc.wordDefaultHangingSpaces === true,
  };
}

function usesWordHangingSpaces(doc: StudioDocument): boolean {
  return doc.wordWrapTrailSpaces === false ||
    (doc.wordWrapTrailSpaces === undefined && doc.wordDefaultHangingSpaces === true);
}

export function isOrdinaryWordBodyParagraph(paragraph: HTMLElement, root: HTMLElement, allowWrapWidgets = false): boolean {
  const parent = paragraph.parentElement;
  return !!parent && (parent === root || parent.matches('.section-editable,[data-section-content]')) &&
    paragraph.matches('p[data-word-paragraph-style]:not([data-hwp-body-line],[data-hwp-cell-line])') &&
    !paragraph.closest('li,td,th,blockquote,[data-column-count]:not([data-column-count="1"]),[data-document-textbox],[data-textbox-content],[data-document-note],.document-furniture') &&
    !paragraph.querySelector('img:not(.ProseMirror-separator),svg,[data-note-reference],[data-page-number],[data-document-field]') &&
    (allowWrapWidgets || !paragraph.querySelector('[data-flow-widget]'));
}

/** The numbering tab is the leading list separator. It is independent of
 * spaces following words at the end of lines in the actual paragraph. */
export function isWordAlignedBulletParagraph(paragraph: HTMLElement): boolean {
  const item = paragraph.parentElement;
  if (!item?.matches('li[data-list-id][data-list-level="0"][data-list-format="bullet"]') ||
    item.dataset.listFont?.toLowerCase() !== 'symbol' || item.dataset.listPattern !== '\uf0b7' ||
    (item.dataset.listSuffix && item.dataset.listSuffix !== 'tab') ||
    item.closest('li li,table,blockquote,[data-column-count],[data-document-textbox],[data-document-note],.document-furniture') ||
    item.children.length !== 1 ||
    !paragraph.matches('p[data-word-paragraph-style]:not([data-hwp-body-line],[data-hwp-cell-line])') ||
    paragraph.querySelector('img:not(.ProseMirror-separator),svg,[data-docx-tab],[data-note-reference],[data-page-number],[data-document-field],[data-flow-widget]') ||
    parseTabStops(paragraph.dataset.tabStops).length) return false;

  const left = cssPixels(paragraph.style.marginLeft);
  const stops = parseWordNumberingTabs(paragraph.dataset.wordNumberingTabs);
  return left !== null && left > 0 && paragraph.style.textIndent.startsWith('-') &&
    (stops.length === 0 || (stops.length === 1 && Math.round(left * 15) === stops[0]));
}

export function isWordGroupedTextBoxParagraph(paragraph: HTMLElement): boolean {
  const content = paragraph.parentElement;
  const box = content?.parentElement;
  if (!content?.hasAttribute('data-textbox-content') || !box?.dataset.documentTextbox ||
      !paragraph.matches('p[data-word-paragraph-style]') ||
      box.closest('.document-furniture,[data-document-note],td,th,[data-column-count]') ||
      paragraph.querySelector('img:not(.ProseMirror-separator),svg,[data-docx-tab],[data-note-reference],[data-page-number],[data-flow-widget],[data-document-field]'))
    return false;
  const frame = textBoxFrame(box.dataset.documentTextbox);
  return !!frame && isWordGroupedRectTextBox(frame);
}
