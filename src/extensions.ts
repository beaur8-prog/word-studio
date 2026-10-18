import { Node, Extension, Mark } from "@tiptap/core";
import StarterKit from "@tiptap/starter-kit";
import { TextStyleKit } from "@tiptap/extension-text-style";
import TextAlign from "@tiptap/extension-text-align";
import { TableKit } from "@tiptap/extension-table";
import Image from "@tiptap/extension-image";
import Highlight from "@tiptap/extension-highlight";
import Subscript from "@tiptap/extension-subscript";
import Superscript from "@tiptap/extension-superscript";
import Placeholder from "@tiptap/extension-placeholder";
import { CellAppearance } from './table-format';
import { TableRowResize } from './table-row-resize';
import { Pagination } from './pagination';
import { ImportedFormat, StudioTableView } from './imported-format';

export const PageBreak = Node.create({
  name: "pageBreak",
  group: "block",
  atom: true,
  selectable: true,
  parseHTML: () => [{ tag: "div[data-page-break]" }],
  renderHTML: () => [
    "div",
    { "data-page-break": "true", class: "page-break" },
    ["span", {}, "페이지 나누기"],
  ],
});
export const CommentAnchor = Mark.create({
  name: 'commentAnchor', inclusive: false,
  addAttributes: () => ({ ids: { default: '', parseHTML: el => el.getAttribute('data-comment-ids'), renderHTML: a => ({ 'data-comment-ids': a.ids }) } }),
  parseHTML: () => [{ tag: 'span[data-comment-ids]' }],
  renderHTML: ({ HTMLAttributes }) => ['span', { ...HTMLAttributes, class: 'comment-anchor' }, 0],
});
export const SectionBreak = Node.create({
  name: 'sectionBreak', group: 'block', atom: true,
  addAttributes: () => ({ page: { default: '', parseHTML: el => el.getAttribute('data-section-page'), renderHTML: a => ({ 'data-section-page': a.page }) } }),
  parseHTML: () => [{ tag: 'div[data-section-page]' }],
  renderHTML: ({ HTMLAttributes }) => ['div', { ...HTMLAttributes, class: 'section-break' }, '구역 나누기'],
});
const ParagraphFormat = Extension.create({
  name: "paragraphFormat",
  addGlobalAttributes() {
    return [
      {
        types: ["paragraph", "heading"],
        attributes: {
          indent: {
            default: 0,
            parseHTML: (e) => parseInt(e.style.marginLeft) || 0,
            renderHTML: (a) =>
              a.indent ? { style: `margin-left:${a.indent}px` } : {},
          },
          lineSpacing: {
            default: null,
            parseHTML: (e) => e.style.lineHeight || null,
            renderHTML: (a) =>
              a.lineSpacing ? { style: `line-height:${a.lineSpacing}` } : {},
          },
        },
      },
    ];
  },
});
export const editorExtensions = [
  StarterKit.configure({
    link: { openOnClick: false, protocols: ["https", "http", "mailto"] },
  }),
  TextStyleKit,
  TextAlign.configure({ types: ["heading", "paragraph"] }),
  TableKit.configure({
    table: { resizable: true, allowTableNodeSelection: true, View: StudioTableView },
  }),
  CellAppearance,
  TableRowResize,
  Image.configure({ allowBase64: true }),
  Highlight.configure({ multicolor: true }),
  Subscript,
  Superscript,
  Placeholder.configure({ placeholder: "여기에 문서를 작성하세요…" }),
  PageBreak,
  SectionBreak,
  CommentAnchor,
  ParagraphFormat,
  ImportedFormat,
  Pagination,
];
