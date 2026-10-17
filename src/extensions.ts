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
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";

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
const paginationKey = new PluginKey("studioPagination");
export const Pagination = Extension.create({
  name: "pagination",
  addProseMirrorPlugins() {
    let scheduled = 0;
    let lastSignature = "";
    return [
      new Plugin({
        key: paginationKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            const next = tr.getMeta(paginationKey);
            return next ?? old.map(tr.mapping, tr.doc);
          },
        },
        props: { decorations: (state) => paginationKey.getState(state) },
        view(view) {
          const calculate = () => {
            cancelAnimationFrame(scheduled);
            scheduled = requestAnimationFrame(() => {
              if (!view.dom.isConnected) return;
              if (view.composing) return;
              const sheet = view.dom.closest(
                ".document-stack",
              ) as HTMLElement | null;
              if (!sheet) return;
              const h = Number(sheet.dataset.pageHeight) || 1122.52,
                m = Number(sheet.dataset.margin ?? 96),
                bottom = Number(sheet.dataset.bottomMargin ?? m),
                g = 24;
              const dec: Decoration[] = [];
              let y = m;
              const sizes: string[] = [];
              view.state.doc.forEach((node, pos) => {
                const dom = view.nodeDOM(pos) as HTMLElement | null;
                if (!dom) return;
                const size =
                  node.type.name === "pageBreak"
                    ? 0
                    : dom.offsetHeight +
                      parseFloat(getComputedStyle(dom).marginBottom || "0");
                let p = Math.floor(y / (h + g));
                let end = p * (h + g) + h - bottom;
                if (node.type.name === "pageBreak") {
                  const gap = Math.max(0, end - y) + m + bottom + g;
                  dec.push(
                    Decoration.node(pos, pos + node.nodeSize, {
                      style: `height:${gap}px`,
                    }),
                  );
                  y += gap;
                  sizes.push(`${pos}:b:${gap}`);
                  return;
                }
                if (y + size > end + 1 && size <= h - m - bottom) {
                  const gap = Math.max(0, end - y) + m + bottom + g;
                  dec.push(
                    Decoration.widget(
                      pos,
                      () => {
                        const el = document.createElement("div");
                        el.className = "pagination-gap";
                        el.style.height = `${gap}px`;
                        el.contentEditable = "false";
                        return el;
                      },
                      { side: -1, key: `gap-${pos}-${gap}` },
                    ),
                  );
                  y += gap;
                  sizes.push(`${pos}:${gap}`);
                }
                y += size;
              });
              const pages = Math.max(1, Math.ceil((y + bottom) / (h + g)));
              sheet.dataset.pages = String(pages);
              sheet.dispatchEvent(
                new CustomEvent("pagecount", { bubbles: true, detail: pages }),
              );
              const signature = `${h}:${m}:${bottom}:${view.state.doc.content.size}:${sizes.join(",")}`;
              if (signature !== lastSignature) {
                lastSignature = signature;
                view.dispatch(
                  view.state.tr
                    .setMeta(
                      paginationKey,
                      DecorationSet.create(view.state.doc, dec),
                    )
                    .setMeta("addToHistory", false),
                );
              }
            });
          };
          calculate();
          const ro = new ResizeObserver(calculate);
          ro.observe(view.dom);
          document.fonts.addEventListener('loadingdone', calculate);
          view.dom.addEventListener('compositionend', calculate);
          return {
            update: calculate,
            destroy() {
              cancelAnimationFrame(scheduled);
              ro.disconnect();
              document.fonts.removeEventListener('loadingdone', calculate);
              view.dom.removeEventListener('compositionend', calculate);
            },
          };
        },
      }),
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
    table: { resizable: true, allowTableNodeSelection: true },
  }),
  Image.configure({ allowBase64: true }),
  Highlight.configure({ multicolor: true }),
  Subscript,
  Superscript,
  Placeholder.configure({ placeholder: "여기에 문서를 작성하세요…" }),
  PageBreak,
  SectionBreak,
  CommentAnchor,
  ParagraphFormat,
  Pagination,
];
