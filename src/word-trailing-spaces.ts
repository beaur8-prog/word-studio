import type { StudioDocument } from './types';

export const wordDefaultHangingSpacesProperty = 'WordStudio.DefaultTrailingSpaces';

export function usesWordHangingSpaces(doc: StudioDocument): boolean {
  return doc.wordWrapTrailSpaces === false ||
    (doc.wordWrapTrailSpaces === undefined && doc.wordDefaultHangingSpaces === true);
}

export function isWordHangingSpaceParagraph(paragraph: HTMLElement, root: HTMLElement): boolean {
  const parent = paragraph.parentElement;
  return !!parent && (parent === root || parent.matches('.section-editable,[data-section-content]')) &&
    paragraph.matches('p[data-word-paragraph-style]:not([data-hwp-body-line],[data-hwp-cell-line])') &&
    !paragraph.closest('li,td,th,blockquote,[data-column-count]:not([data-column-count="1"]),[data-document-textbox],[data-textbox-content],[data-document-note],.document-furniture') &&
    !paragraph.querySelector('img:not(.ProseMirror-separator),svg,[data-note-reference],[data-page-number],[data-flow-widget],[data-document-field]');
}
