import { Extension } from "@tiptap/core";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import { imageFloating, type ImageFloating } from "./image-floating";

type Boundary = [globalThis.Node, number];
type Unit = {
  from: number;
  to: number;
  start: Boundary;
  end: Boundary;
  text: string;
  hard?: boolean;
  floating?: boolean;
};
type Widget = {
  offset: number;
  width?: number;
  break?: boolean;
  height?: number;
};
type Picture = {
  index: number;
  left: number;
  top: number;
  width: number;
  height: number;
  cssLeft: number;
  cssTop: number;
  geometry: ImageFloating;
};
type ParagraphPlan = { index: number; widgets: Widget[]; pictures: Picture[] };
const atomic =
  "img,[data-symbol-font],[data-word-overlap],[data-docx-tab],[data-docx-bookmark],[data-note-reference],[data-page-number]";
const floatingImages = (root: HTMLElement) =>
  [...root.querySelectorAll<HTMLImageElement>("img[data-image-floating]")].filter(
    (img) => !img.closest(".document-furniture,.paper-backgrounds,.section-papers"),
  );
const paragraphs = (root: HTMLElement) =>
  [...root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")].filter(
    (p) => !p.closest(".document-furniture,.paper-backgrounds,.section-papers"),
  );

function units(p: HTMLElement) {
  const result: Unit[] = [];
  let offset = 0;
  const visit = (node: globalThis.Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      for (const part of new Intl.Segmenter(undefined, {
        granularity: "grapheme",
      }).segment(node.textContent || "")) {
        result.push({
          from: offset,
          to: offset + part.segment.length,
          start: [node, part.index],
          end: [node, part.index + part.segment.length],
          text: part.segment,
        });
        offset += part.segment.length;
      }
    } else if (node instanceof HTMLElement) {
      if (
        node.hasAttribute("data-flow-widget") ||
        node.classList.contains("ProseMirror-separator") ||
        node.classList.contains("ProseMirror-trailingBreak")
      )
        return;
      if (node.matches(atomic) || node.tagName === "BR") {
        const index = [...node.parentNode!.childNodes].indexOf(node);
        result.push({
          from: offset,
          to: offset + 1,
          start: [node.parentNode!, index],
          end: [node.parentNode!, index + 1],
          text: node.tagName === "BR" ? "\n" : "\ufffc",
          hard: node.tagName === "BR",
          floating: node.matches("img[data-image-floating]"),
        });
        offset++;
      } else node.childNodes.forEach(visit);
    }
  };
  p.childNodes.forEach(visit);
  return result;
}
function widgetDOM(widget: Widget) {
  const el = document.createElement(widget.break ? "br" : "span");
  el.dataset.flowWidget = "true";
  el.contentEditable = "false";
  el.setAttribute("aria-hidden", "true");
  if (!widget.break)
    Object.assign(
      el.style,
      widget.height !== undefined
        ? { display: "block", height: `${widget.height}px`, lineHeight: "0" }
        : {
            display: "inline-block",
            width: `${Math.max(0, widget.width || 0)}px`,
            height: "0",
            verticalAlign: "baseline",
            pointerEvents: "none",
          },
    );
  return el;
}
function insertWidgets(p: HTMLElement, widgets: Widget[]) {
  const index = units(p);
  for (const { widget } of widgets
    .map((widget, order) => ({ widget, order }))
    .sort((a, b) => b.widget.offset - a.widget.offset || b.order - a.order)) {
    const unit = index.find((u) => u.from === widget.offset);
    const boundary = unit?.start || index.at(-1)?.end;
    if (!boundary) continue;
    const range = document.createRange();
    range.setStart(...boundary);
    range.collapse(true);
    range.insertNode(widgetDOM(widget));
  }
}

/** Return the part of a polygon intersected by an entire text line's ink band.
 * Sampling its vertices as well as both band edges prevents sloped contours
 * from cutting through letters between the sampled baselines. */
function exclusion(picture: Picture, y: number, height: number): [number, number][] {
  const { geometry: g, left, top, width, height: ph } = picture,
    m = g.margins;
  const y1 = y - (m?.bottom || 0) / 9525,
    y2 = y + height + (m?.top || 0) / 9525;
  if (g.wrap?.type === 0 || y2 <= top + 0.1 || y1 >= top + ph - 0.1) return [];
  if (g.wrap?.type === 3) return [[-1e7, 1e7]];
  let a = left,
    b = left + width;
  if ((g.wrap?.type === 2 || g.wrap?.type === 4) && g.polygon) {
    const low = Math.max(0, (y1 - top) / ph),
      high = Math.min(1, (y2 - top) / ph),
      values: number[] = [];
    const ys = [low, high, ...g.polygon.map((p) => p[1]).filter((v) => v > low && v < high)];
    for (const scan of ys)
      for (let i = 0; i < g.polygon.length; i++) {
        const p = g.polygon[i],
          q = g.polygon[(i + 1) % g.polygon.length];
        if (Math.abs(p[1] - q[1]) < 1e-8) {
          if (Math.abs(scan - p[1]) < 1e-6) values.push(p[0], q[0]);
        } else if (scan >= Math.min(p[1], q[1]) - 1e-7 && scan <= Math.max(p[1], q[1]) + 1e-7)
          values.push(p[0] + ((q[0] - p[0]) * (scan - p[1])) / (q[1] - p[1]));
      }
    if (!values.length) return [];
    a = left + Math.min(...values) * width;
    b = left + Math.max(...values) * width;
  }
  return [[a - (m?.left || 0) / 9525, b + (m?.right || 0) / 9525]];
}
function slots(pictures: Picture[], x: number, y: number, width: number, height: number) {
  let spans: [number, number][] = [[x, x + width]];
  for (const picture of pictures)
    for (let [a, b] of exclusion(picture, y, height)) {
      const side = picture.geometry.wrap?.side;
      if (side === "left") b = 1e7;
      if (side === "right") a = -1e7;
      if (side === "largest") {
        if (a - x >= x + width - b) b = 1e7;
        else a = -1e7;
      }
      spans = spans.flatMap(([l, r]) =>
        b <= l || a >= r
          ? [[l, r] as [number, number]]
          : ([
              [l, Math.min(r, a)],
              [Math.max(l, b), r],
            ].filter(([l, r]) => r - l >= 4) as [number, number][]),
      );
    }
  return spans;
}
function imagePosition(
  img: HTMLImageElement,
  index: number,
  root: HTMLElement,
): Picture | undefined {
  const g = imageFloating(img.dataset.imageFloating);
  if (!g) return;
  const p = img.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6");
  if (!p) return;
  p.style.position = "relative";
  const rect = p.getClientRects()[0] || p.getBoundingClientRect(),
    group = p.closest<HTMLElement>("[data-column-count]");
  const sheet = p.closest<HTMLElement>("[data-page-height],.print-page") || root;
  const sr = sheet.getBoundingClientRect(),
    width = Number(img.getAttribute("width")) || img.width,
    height = Number(img.getAttribute("height")) || img.height;
  const ph = Number(sheet.dataset.pageHeight) || parseFloat(sheet.style.height) || 1122.52;
  const textRoot = p.closest<HTMLElement>(".section-editable") || root;
  const rootStyle = getComputedStyle(textRoot),
    top =
      Number(sheet.dataset.margin) ||
      parseFloat(rootStyle.paddingTop) ||
      parseFloat(root.style.top) ||
      96;
  const bottom = Number(sheet.dataset.bottomMargin) || parseFloat(rootStyle.paddingBottom) || 96;
  const page = Math.max(0, Math.floor((rect.top - sr.top - top + 0.5) / (ph + 24))),
    py = sr.top + page * (ph + 24);
  const leftMargin = parseFloat(rootStyle.paddingLeft) || parseFloat(root.style.left) || 0;
  const rightMargin = parseFloat(rootStyle.paddingRight) || leftMargin;
  const cell = p.closest<HTMLElement>("td,th");
  const cellStyle = cell ? getComputedStyle(cell) : undefined;
  const column = cell ? cell.getBoundingClientRect().left + (parseFloat(cellStyle!.paddingLeft) || 0) : group
    ? group.getBoundingClientRect().left +
      Math.floor(
        (rect.left - group.getBoundingClientRect().left + 0.5) /
          ((group.clientWidth + Number(group.dataset.columnSpace)) /
            Number(group.dataset.columnCount)),
      ) *
        ((group.clientWidth + Number(group.dataset.columnSpace)) /
          Number(group.dataset.columnCount))
    : sr.left + leftMargin;
  const horizontal = g.horizontalPosition,
    vertical = g.verticalPosition;
  const hbase =
    horizontal.relative === "page"
      ? sr.left
      : horizontal.relative === "column"
        ? column
        : sr.left + leftMargin;
  const hwidth =
    horizontal.relative === "page"
      ? sheet.clientWidth
      : horizontal.relative === "column" && cell
        ? cell.getBoundingClientRect().width - (parseFloat(cellStyle!.paddingLeft) || 0) - (parseFloat(cellStyle!.paddingRight) || 0)
      : horizontal.relative === "column" && group
        ? (group.clientWidth -
            (Number(group.dataset.columnCount) - 1) * Number(group.dataset.columnSpace)) /
          Number(group.dataset.columnCount)
        : sheet.clientWidth - leftMargin - rightMargin;
  const vbase =
    vertical.relative === "page" ? py : vertical.relative === "margin" ? py + top : rect.top;
  const vh =
    vertical.relative === "page"
      ? ph
      : vertical.relative === "margin"
        ? ph - top - bottom
        : rect.height;
  const left =
    hbase +
    (horizontal.align === "center"
      ? (hwidth - width) / 2
      : horizontal.align === "right"
        ? hwidth - width
        : (horizontal.offset || 0) / 9525);
  let y =
    vbase +
    (vertical.align === "center"
      ? (vh - height) / 2
      : vertical.align === "bottom"
        ? vh - height
        : (vertical.offset || 0) / 9525);
  if (
    g.wrap?.type !== 0 &&
    height <= ph - top - bottom &&
    !["paragraph", "line"].includes(vertical.relative || "paragraph")
  )
    y = Math.max(py + top, Math.min(y, py + ph - bottom - height));
  Object.assign(img.style, {
    position: "absolute",
    float: "none",
    margin: "0",
    left: `${left - rect.left}px`,
    top: `${y - rect.top}px`,
    zIndex: g.behindDocument ? "-1" : "1",
  });
  return {
    index,
    left,
    top: y,
    width,
    height,
    cssLeft: left - rect.left,
    cssTop: y - rect.top,
    geometry: g,
  };
}

/** Plan in a detached copy of the live geometry. The editable document receives
 * only ProseMirror decorations: no split paragraphs, copied text or DOM moves. */
export function planFloatingLayout(root: HTMLElement): ParagraphPlan[] {
  if (!root.querySelector("img[data-image-floating]")) return [];
  const owner = root.closest<HTMLElement>(".document-stack,.print-page");
  const host = (owner?.cloneNode(false) || document.createElement("div")) as HTMLElement;
  Object.assign(host.style, {
    position: "fixed",
    left: "-100000px",
    top: "0",
    transform: "none",
    zoom: "1",
    visibility: "hidden",
    pointerEvents: "none",
    width: `${parseFloat(getComputedStyle(owner || root).width)}px`,
  });
  const copy = root.cloneNode(true) as HTMLElement;
  copy.querySelectorAll("[data-flow-widget]").forEach((el) => el.remove());
  copy.removeAttribute("contenteditable");
  host.append(copy);
  document.body.append(host);
  const measure = document.createElement("span");
  Object.assign(measure.style, {
    position: "fixed",
    left: "-100000px",
    top: "0",
    visibility: "hidden",
    display: "inline-block",
    whiteSpace: "pre",
    width: "max-content",
    padding: "0",
    margin: "0",
    textIndent: "0",
  });
  measure.className = "flow-text-measure";
  host.append(measure);
  try {
    const ps = paragraphs(copy),
      images = floatingImages(copy);
    const pictures = images
      .map((img, index) => imagePosition(img, index, copy))
      .filter((p): p is Picture => !!p);
    const result: ParagraphPlan[] = [];
    for (let pi = 0; pi < ps.length; pi++) {
      const p = ps[pi],
        own = images
          .flatMap((img, index) =>
            img.closest("p,h1,h2,h3,h4,h5,h6") === p ? [imagePosition(img, index, copy)!] : [],
          )
          .filter(Boolean);
      for (const picture of own) {
        const i = pictures.findIndex((v) => v.index === picture.index);
        if (i >= 0) pictures[i] = picture;
      }
      // Behind/in-front artwork needs positioning but does not exclude text.
      // Its often page-sized bounds must not trigger a manual re-layout of
      // unrelated paragraphs, which changes justification and mixed-font
      // line endings even though every line still has its full width.
      const flowing = pictures.filter(pic => pic.geometry.wrap?.type !== 0);
      const style = getComputedStyle(p),
        box = p.getClientRects()[0] || p.getBoundingClientRect();
      const lh = parseFloat(style.lineHeight) || parseFloat(style.fontSize) * 1.3,
        width = parseFloat(style.width) || box.width;
      const group = p.closest<HTMLElement>("[data-column-count]");
      const groupBox = group?.getBoundingClientRect(),
        colHeight = group ? parseFloat(group.style.getPropertyValue("column-height")) : 0;
      const colGap = Number(group?.dataset.columnSpace) || 0,
        colWidth = group
          ? (group.clientWidth - (Number(group.dataset.columnCount) - 1) * colGap) /
            Number(group.dataset.columnCount)
          : width;
      const rowGap = group ? parseFloat(group.style.rowGap) || 0 : 0;
      let x = box.left,
        y = box.top;
      const relevant = () =>
        flowing.some((pic) => exclusion(pic, y, lh).some(([a, b]) => b > x && a < x + width));
      if (
        !flowing.some((pic) => pic.top + pic.height >= y - 0.5 && pic.top < box.bottom + 0.5)
      ) {
        if (own.length) result.push({ index: pi, widgets: [], pictures: own });
        continue;
      }
      const tokens = units(p),
        widgets: Widget[] = [],
        cache = new Map<string, { width: number; height: number }>();
      for (const property of [
        "fontFamily",
        "fontSize",
        "fontWeight",
        "fontStyle",
        "fontStretch",
        "fontVariant",
        "letterSpacing",
        "wordSpacing",
        "textTransform",
        "lineHeight",
      ] as const)
        measure.style[property] = style[property];
      // Keep unitless leading unitless: a larger inline font must enlarge
      // its actual line box just as it does in the editable paragraph.
      measure.style.lineHeight = p.style.lineHeight || String(lh / parseFloat(style.fontSize));
      const metrics = (start: number, end: number) => {
        if (start === end) return { width: 0, height: lh };
        const key = `${start}:${end}`;
        if (cache.has(key)) return cache.get(key)!;
        const range = document.createRange();
        range.setStart(...tokens[start].start);
        range.setEnd(...tokens[end - 1].end);
        let fragment: globalThis.Node = range.cloneContents();
        let ancestor =
          range.commonAncestorContainer.nodeType === Node.TEXT_NODE
            ? range.commonAncestorContainer.parentElement
            : (range.commonAncestorContainer as HTMLElement);
        // A range entirely inside a bold/link/style run does not clone that
        // common ancestor. Restore its inherited styles for faithful widths.
        while (ancestor && ancestor !== p) {
          const wrapper = ancestor.cloneNode(false);
          wrapper.appendChild(fragment);
          fragment = wrapper;
          ancestor = ancestor.parentElement;
        }
        measure.replaceChildren(fragment);
        measure
          .querySelectorAll("img[data-image-floating],[data-flow-widget]")
          .forEach((el) => el.remove());
        const rect = measure.getBoundingClientRect();
        const value = { width: rect.width, height: Math.max(lh, rect.height) };
        cache.set(key, value);
        return value;
      };
      const size = (start: number, end: number) => metrics(start, end).width;
      const fit = (start: number, end: number, available: number) => {
        let low = start,
          high = end,
          best = start;
        while (low <= high) {
          const mid = (low + high) >> 1;
          if (size(start, mid) <= available + 0.15) {
            best = mid;
            low = mid + 1;
          } else high = mid - 1;
        }
        if (
          best < end &&
          best > start &&
          /[\p{Script=Latin}\p{N}]/u.test(tokens[best - 1].text) &&
          /[\p{Script=Latin}\p{N}]/u.test(tokens[best].text)
        ) {
          let boundary = best;
          while (boundary > start && /[\p{Script=Latin}\p{N}]/u.test(tokens[boundary - 1].text))
            boundary--;
          if (boundary > start) best = boundary;
        }
        return best;
      };
      let start = 0,
        line = 0;
      while (start < tokens.length && line++ < 10000) {
        if (
          !relevant() &&
          y >
            Math.max(
              ...flowing.map(
                (pic) => pic.top + pic.height + (pic.geometry.margins?.bottom || 0) / 9525,
              ),
            )
        )
          break;
        const indent = line === 1 ? parseFloat(style.textIndent) || 0 : 0;
        const hard = tokens.findIndex((token, i) => i >= start && token.hard);
        const end = hard < 0 ? tokens.length : hard,
          lineStart = start;
        let lineHeight = metrics(start, fit(start, end, width - indent)).height;
        const arrange = (height: number) => {
          const lineWidgets: Widget[] = [];
          const spans = slots(flowing, x + indent, y, width - indent, height);
          let nextStart = lineStart,
            cursor = x + indent;
          for (const [left, right] of spans) {
            const next = fit(nextStart, end, right - left);
            if (
              next === nextStart &&
              right - left < Math.min(width, size(nextStart, Math.min(nextStart + 1, end)))
            )
              continue;
            const last = next === nextStart && nextStart < end ? nextStart + 1 : next;
            const textWidth = size(nextStart, last);
            const alignment =
              style.textAlign === "right" || style.textAlign === "end"
                ? 1
                : style.textAlign === "center"
                  ? 0.5
                  : 0;
            const placedLeft = left + Math.max(0, right - left - textWidth) * alignment;
            if (placedLeft > cursor + 0.2)
              lineWidgets.push({
                offset: tokens[nextStart]?.from ?? tokens.at(-1)!.to,
                width: placedLeft - cursor,
              });
            cursor = placedLeft + textWidth;
            nextStart = last;
            if (nextStart >= end) break;
          }
          return { lineWidgets, nextStart, cursor, blocked: !spans.length };
        };
        let arranged = arrange(lineHeight);
        // Exclusions depend on line height, and the selected words determine
        // that height. Re-evaluate when a font-size boundary changes the fit.
        for (let attempt = 0; attempt < 3 && !arranged.blocked; attempt++) {
          const actual = metrics(lineStart, arranged.nextStart).height;
          if (Math.abs(actual - lineHeight) < 0.1) break;
          lineHeight = actual;
          arranged = arrange(lineHeight);
        }
        if (arranged.blocked) {
          const jump = Math.max(
            lineHeight,
            Math.min(
              ...flowing
                .filter((pic) => pic.top + pic.height > y)
                .map(
                  (pic) => pic.top + pic.height - y + (pic.geometry.margins?.bottom || 0) / 9525,
                ),
            ),
          );
          widgets.push({ offset: tokens[start].from, height: jump });
          y += jump;
          continue;
        }
        widgets.push(...arranged.lineWidgets);
        start = arranged.nextStart;
        // Reserve unused width so native center/right alignment cannot shift
        // a measured gap over a picture. The unplanned tail keeps its alignment.
        const remaining = x + width - arranged.cursor - 0.4;
        if (remaining > 0.2)
          widgets.push({
            offset: tokens[start]?.from ?? tokens.at(-1)!.to,
            width: remaining,
          });
        if (start === end && hard >= 0) {
          start++;
          y += lineHeight;
        } else if (start < tokens.length) {
          widgets.push({ offset: tokens[start].from, break: true });
          y += lineHeight;
        } else break;
        if (group && groupBox && colHeight) {
          const row = Math.floor((y - groupBox.top) / (colHeight + rowGap)),
            rowTop = groupBox.top + row * (colHeight + rowGap);
          if (y + lh > rowTop + colHeight + 0.5) {
            const col = Math.round((x - groupBox.left) / (colWidth + colGap));
            if (col + 1 < Number(group.dataset.columnCount)) {
              x += colWidth + colGap;
              y = rowTop;
            } else {
              x = groupBox.left;
              y = rowTop + colHeight + rowGap;
            }
          }
        }
      }
      insertWidgets(p, widgets);
      result.push({ index: pi, widgets, pictures: own });
    }
    // Flow widgets can change a vertically aligned cell's content height and
    // move its anchor paragraph. Resolve absolute picture offsets against
    // that final geometry, not the paragraph's pre-wrap position.
    for (const paragraph of result) paragraph.pictures = paragraph.pictures.map(picture =>
      imagePosition(images[picture.index], picture.index, copy) || picture);
    return result;
  } finally {
    host.remove();
  }
}

export function applyFloatingLayout(root: HTMLElement) {
  root.querySelectorAll("[data-flow-widget]").forEach((el) => el.remove());
  const plan = planFloatingLayout(root),
    ps = paragraphs(root),
    images = floatingImages(root);
  for (const paragraph of plan) {
    const p = ps[paragraph.index];
    if (paragraph.pictures.length) p.style.position = "relative";
    for (const picture of paragraph.pictures) {
      const img = images[picture.index];
      const own = imagePosition(img, picture.index, root);
      if (own) img.style.zIndex = own.geometry.behindDocument ? "-1" : "1";
    }
    insertWidgets(p, paragraph.widgets);
  }
  images.forEach((img, index) => imagePosition(img, index, root));
}

const key = new PluginKey("floatingLayout");
export const FloatingLayout = Extension.create({
  name: "floatingLayout",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply: (tr, old) => tr.getMeta(key) ?? old.map(tr.mapping, tr.doc),
        },
        props: { decorations: (state) => key.getState(state) },
        view(view) {
          let frame = 0,
            signature = "",
            lastDocument = view.state.doc;
          const calculate = () => {
            cancelAnimationFrame(frame);
            frame = requestAnimationFrame(() => {
              if (
                !view.dom.isConnected ||
                view.composing ||
                view.dom.classList.contains("table-row-resizing")
              )
                return;
              const ps = paragraphs(view.dom),
                pos = new Map<HTMLElement, number>();
              view.state.doc.descendants((node, position) => {
                if (node.isTextblock) {
                  const dom = view.nodeDOM(position);
                  if (dom instanceof HTMLElement) pos.set(dom, position);
                }
              });
              const plan = planFloatingLayout(view.dom),
                decorations: Decoration[] = [],
                values: string[] = [];
              for (const paragraph of plan) {
                const p = ps[paragraph.index],
                  position = pos.get(p);
                if (position === undefined) continue;
                if (paragraph.pictures.length)
                  decorations.push(
                    Decoration.node(
                      position,
                      position + view.state.doc.nodeAt(position)!.nodeSize,
                      { style: "position:relative" },
                    ),
                  );
                const index = units(p);
                for (const picture of paragraph.pictures) {
                  const token = index.find(
                    (u) =>
                      u.floating &&
                      u.start[0].childNodes[u.start[1]] === floatingImages(view.dom)[picture.index],
                  );
                  if (!token) continue;
                  const g = picture.geometry;
                  const style = `position:absolute;float:none;margin:0;left:${picture.cssLeft}px;top:${picture.cssTop}px;z-index:${g.behindDocument ? -1 : 1}`;
                  decorations.push(
                    Decoration.node(position + 1 + token.from, position + 1 + token.to, { style }),
                  );
                  values.push(`image:${position + token.from}:${style}`);
                }
                for (const [order, widget] of paragraph.widgets.entries()) {
                  const at = position + 1 + widget.offset;
                  decorations.push(
                    Decoration.widget(at, () => widgetDOM(widget), {
                      side: -10000 + order,
                      key: `${position}:${order}:${JSON.stringify(widget)}`,
                    }),
                  );
                  values.push(`${at}:${JSON.stringify(widget)}`);
                }
              }
              const next = values.join("|");
              if (next !== signature || lastDocument !== view.state.doc) {
                signature = next;
                lastDocument = view.state.doc;
                view.dispatch(
                  view.state.tr
                    .setMeta(key, DecorationSet.create(view.state.doc, decorations))
                    .setMeta("addToHistory", false),
                );
              }
            });
          };
          const observer = new ResizeObserver(calculate);
          observer.observe(view.dom);
          view.dom.addEventListener("load", calculate, true);
          view.dom.addEventListener("compositionend", calculate);
          document.fonts.addEventListener("loadingdone", calculate);
          calculate();
          return {
            update: calculate,
            destroy() {
              cancelAnimationFrame(frame);
              observer.disconnect();
              view.dom.removeEventListener("load", calculate, true);
              view.dom.removeEventListener("compositionend", calculate);
              document.fonts.removeEventListener("loadingdone", calculate);
            },
          };
        },
      }),
    ];
  },
});
