import DOMPurify from "dompurify";
import documentLimits from '../electron/document-limits.json';
import { encodeHwpCellLine, parseHwpCellLine } from './hwp-cell-line';
import { encodeHwpBodyLine, parseHwpBodyLine } from './hwp-body-line';
// Base64 expansion of one supported embedded raster, plus its data-URI header.
const MAX_IMAGE_SOURCE_LENGTH = Math.ceil(documentLimits.embeddedImageBytes / 3) * 4 + 64;
export const RASTER_DATA =
  /^data:image\/(png|jpe?g|gif|webp|bmp);base64,([A-Za-z0-9+/=\s]+)$/i;
const SAFE_LINK = /^(https?:\/\/|mailto:|tel:|#)/i;
export function validImage(src: unknown): src is string {
  return (
    typeof src === "string" &&
    src.length <= MAX_IMAGE_SOURCE_LENGTH &&
    RASTER_DATA.test(src)
  );
}
export function safeLink(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= 8192 &&
    SAFE_LINK.test(value) &&
    !/[\u0000-\u001f]/.test(value)
  );
}

export function sanitizeImportedHtml(source: string): {
  html: string;
  warnings: string[];
} {
  const purified = DOMPurify.sanitize(source, {
    ALLOWED_TAGS: [
      "p",
      "br",
      "div",
      "span",
      "h1",
      "h2",
      "h3",
      "h4",
      "h5",
      "h6",
      "strong",
      "b",
      "em",
      "i",
      "u",
      "s",
      "strike",
      "del",
      "sub",
      "sup",
      "mark",
      "a",
      "ul",
      "ol",
      "li",
      "blockquote",
      "pre",
      "code",
      "hr",
      "table",
      "colgroup", "col",
      "thead",
      "tbody",
      "tfoot",
      "tr",
      "th",
      "td",
      "img",
    ],
    ALLOWED_ATTR: [
      "style",
      "href",
      "src",
      "alt",
      "title",
      "width",
      "height",
      "colspan",
      "colwidth",
      "rowspan",
      "start",
      "data-type",
      "data-color",
      "data-colwidth",
      "data-page-break",
      "data-column-break",
      "data-section-page",
      "data-comment-ids",
      "data-inline-image",
      "data-image-floating",
      "data-table-anchor",
      "data-docx-tab", "data-tab-stops",
      "data-word-line-rule", "data-word-line",
      "data-exact-line-height",
      "data-hwp-cell-line",
      "data-hwp-body-line",
      "data-word-row-rule",
      "data-row-split", "data-cell-split",
      "data-word-overlap",
      "data-docx-bookmark", "id",
      "data-list-id", "data-list-level", "data-list-format", "data-list-pattern", "data-list-label",
      "data-document-note", "data-note-reference", "data-note-kind", "data-note-label", "data-page-number",
      "data-column-count", "data-column-space", "data-column-balance", "data-section-content",
      "data-symbol-font", "data-symbol-code", "data-symbol-glyph",
    ],
    ALLOW_DATA_ATTR: false,
  });
  const template = document.createElement("template");
  template.innerHTML = purified;
  const warnings: string[] = [];
  let removedImages = 0;
  for (const img of template.content.querySelectorAll("img")) {
    if (!validImage(img.getAttribute("src"))) {
      img.replaceWith(
        document.createTextNode(img.getAttribute("alt") || "[외부 그림]"),
      );
      removedImages++;
    }
  }
  if (removedImages)
    warnings.push(
      `외부 주소 또는 지원하지 않는 그림 ${removedImages}개를 제외했습니다.`,
    );
  for (const anchor of template.content.querySelectorAll("a")) {
    if (!safeLink(anchor.getAttribute("href"))) anchor.removeAttribute("href");
    anchor.setAttribute("rel", "noopener noreferrer");
    anchor.setAttribute("target", "_blank");
  }
  const safeStyles = new Set([
    "font-family",
    "font-size",
    "font-weight",
    "font-style",
    "text-decoration",
    "text-decoration-line",
    "text-decoration-style", "font-variant-caps", "text-transform",
    "color",
    "background-color",
    "text-align",
    "line-height",
    "margin-left",
    "margin-right", "margin-top", "margin-bottom", "text-indent", "letter-spacing",
    "vertical-align",
    "white-space",
    "word-break",
    "widows", "orphans",
    "width",
    "height",
    "border", "border-top", "border-right", "border-bottom", "border-left",
    "border-style", "border-width", "border-color",
    ...['top', 'right', 'bottom', 'left'].flatMap(side => [`border-${side}-style`, `border-${side}-width`, `border-${side}-color`]),
    "padding", "padding-top", "padding-right", "padding-bottom", "padding-left",
  ]);
  for (const element of template.content.querySelectorAll<HTMLElement>(
    "[style]",
  )) {
    const oldStyle = element.style;
    const cellLine = parseHwpCellLine(oldStyle.getPropertyValue('--hwp-cell-line') || element.dataset.hwpCellLine);
    if (cellLine) element.dataset.hwpCellLine = encodeHwpCellLine(cellLine);
    else element.removeAttribute('data-hwp-cell-line');
    const bodyLine = parseHwpBodyLine(oldStyle.getPropertyValue('--hwp-body-line') || element.dataset.hwpBodyLine);
    if (bodyLine) element.dataset.hwpBodyLine = encodeHwpBodyLine(bodyLine);
    else element.removeAttribute('data-hwp-body-line');
    const output: string[] = [];
    // New Chromium versions enumerate white-space's longhands instead of the
    // authored shorthand. Retain the validated shorthand for import parsers.
    if (oldStyle.whiteSpace && !/url\s*\(|var\s*\(|\\/i.test(oldStyle.whiteSpace)) output.push(`white-space:${oldStyle.whiteSpace}`);
    for (const name of Array.from(oldStyle)) {
      const value = oldStyle.getPropertyValue(name);
      if (
        safeStyles.has(name) &&
        !/url\s*\(|expression\s*\(|@import|var\s*\(|\\/i.test(value) &&
        value.length < 300
      )
        output.push(`${name}:${value}`);
    }
    element.removeAttribute("style");
    if (output.length) element.setAttribute("style", output.join(";"));
  }
  return { html: template.innerHTML || "<p></p>", warnings };
}

