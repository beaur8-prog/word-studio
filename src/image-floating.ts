import type { IFloating } from "docx";
import type JSZip from "jszip";
export type ImageFloating = IFloating & { polygon?: [number, number][] };
export type ImageGeometryRecord = { name: string; geometry: ImageFloating };

/** docx emits a rectangular wrapPolygon. Replace it with the document's
 * contour in every story and remove the temporary matching name. */
export async function restoreImageGeometry(zip: JSZip, records: Map<string, ImageGeometryRecord>) {
  const ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  for (const file of Object.keys(zip.files).filter((name) => /^word\/.*\.xml$/.test(name))) {
    const source = await zip.file(file)!.async("string");
    if (![...records.keys()].some((key) => source.includes(key))) continue;
    const xml = new DOMParser().parseFromString(source, "application/xml");
    for (const properties of xml.getElementsByTagNameNS(ns, "docPr")) {
      const record = records.get(properties.getAttribute("name") || "");
      if (!record) continue;
      properties.setAttribute("name", record.name);
      const anchor = properties.parentElement,
        wrap =
          anchor &&
          [...anchor.children].find(
            (el) => el.localName === "wrapTight" || el.localName === "wrapThrough",
          );
      if (!wrap || !record.geometry.polygon) continue;
      const polygon = xml.createElementNS(ns, "wp:wrapPolygon");
      polygon.setAttribute("edited", "0");
      record.geometry.polygon.forEach(([x, y], index) => {
        const point = xml.createElementNS(ns, index ? "wp:lineTo" : "wp:start");
        point.setAttribute("x", String(Math.round(x * 21600)));
        point.setAttribute("y", String(Math.round(y * 21600)));
        polygon.append(point);
      });
      const previous = wrap.getElementsByTagNameNS(ns, "wrapPolygon")[0];
      if (previous) previous.replaceWith(polygon);
      else wrap.append(polygon);
    }
    zip.file(file, new XMLSerializer().serializeToString(xml));
  }
}

export function imageFloating(value: unknown): ImageFloating | undefined {
  if (typeof value !== "string" || value.length > 10000) return;
  try {
    const data = JSON.parse(value);
    const position = (p: any, vertical: boolean) =>
      p &&
      typeof p === "object" &&
      (p.offset === undefined ||
        (typeof p.offset === "number" &&
          Number.isFinite(p.offset) &&
          Math.abs(p.offset) <= 100000000)) &&
      (p.align === undefined ||
        (vertical
          ? ["top", "center", "bottom", "inside", "outside"]
          : ["left", "center", "right", "inside", "outside"]
        ).includes(p.align)) &&
      (p.relative === undefined ||
        (vertical
          ? [
              "page",
              "margin",
              "paragraph",
              "line",
              "topMargin",
              "bottomMargin",
              "insideMargin",
              "outsideMargin",
            ]
          : [
              "page",
              "margin",
              "column",
              "character",
              "leftMargin",
              "rightMargin",
              "insideMargin",
              "outsideMargin",
            ]
        ).includes(p.relative));
    if (
      !position(data.horizontalPosition, false) ||
      !position(data.verticalPosition, true) ||
      ![0, 1, 2, 3, 4].includes(data.wrap?.type)
    )
      return;
    const polygon =
      Array.isArray(data.polygon) &&
      data.polygon.length >= 3 &&
      data.polygon.length <= 256 &&
      data.polygon.every(
        (p: unknown) =>
          Array.isArray(p) &&
          p.length === 2 &&
          p.every((n) => typeof n === "number" && Number.isFinite(n) && Math.abs(n) <= 4),
      )
        ? (data.polygon as [number, number][])
        : undefined;
    if (data.polygon !== undefined && !polygon) return;
    return {
      horizontalPosition: data.horizontalPosition,
      verticalPosition: data.verticalPosition,
      wrap: {
        type: data.wrap.type,
        ...(["bothSides", "left", "right", "largest"].includes(data.wrap.side)
          ? { side: data.wrap.side }
          : {}),
      },
      ...(polygon ? { polygon } : {}),
      ...(data.margins && typeof data.margins === "object"
        ? {
            margins: Object.fromEntries(
              ["top", "right", "bottom", "left"].map((k) => [
                k,
                Math.max(0, Math.min(10000000, Number(data.margins[k]) || 0)),
              ]),
            ),
          }
        : {}),
      behindDocument: data.behindDocument === true,
      allowOverlap: true,
      layoutInCell: true,
    };
  } catch {
    return;
  }
}

export function floatingCss(value: unknown, width: number) {
  const data = imageFloating(value);
  if (!data) return "";
  const horizontal = data.horizontalPosition,
    vertical = data.verticalPosition;
  const left =
    horizontal.align === "center"
      ? `calc(50% - ${width / 2}px)`
      : horizontal.align === "right"
        ? `calc(100% - ${width}px)`
        : `${(horizontal.offset || 0) / 9525}px`;
  return `position:absolute;float:none;left:${left};top:${vertical.relative === "paragraph" ? (vertical.offset || 0) / 9525 : 0}px;margin:0;z-index:${data.behindDocument ? -1 : 1}`;
}

export function drawingFloating(drawing: Element) {
  const normalized = drawing.getAttribute("data-ws-floating");
  if (normalized) return imageFloating(normalized);
  const anchor = [...drawing.children].find((el) => el.localName === "anchor");
  if (!anchor) return;
  const read = (kind: string) => {
    const p = [...anchor.children].find((el) => el.localName === kind);
    const align = [...(p?.children || [])].find((el) => el.localName === "align")?.textContent;
    const offset = Number(
      [...(p?.children || [])].find((el) => el.localName === "posOffset")?.textContent || 0,
    );
    return {
      relative: p?.getAttribute("relativeFrom") || (kind === "positionH" ? "column" : "paragraph"),
      ...(align ? { align } : { offset }),
    };
  };
  const type = [...anchor.children].find((el) => /^wrap/.test(el.localName))?.localName;
  const wrapping = [...anchor.children].find((el) => el.localName === type);
  const points = [...(wrapping?.getElementsByTagNameNS("*", "wrapPolygon")[0]?.children || [])].map(
    (p) => [Number(p.getAttribute("x")) / 21600, Number(p.getAttribute("y")) / 21600],
  );
  return imageFloating(
    JSON.stringify({
      horizontalPosition: read("positionH"),
      verticalPosition: read("positionV"),
      wrap: {
        type:
          (
            {
              wrapNone: 0,
              wrapSquare: 1,
              wrapTight: 2,
              wrapTopAndBottom: 3,
              wrapThrough: 4,
            } as Record<string, number>
          )[type || ""] ?? 1,
        side: wrapping?.getAttribute("wrapText"),
      },
      ...(points.length >= 3 ? { polygon: points } : {}),
      margins: Object.fromEntries(
        ["top", "right", "bottom", "left"].map((k) => [
          k,
          Number(anchor.getAttribute(`dist${k[0].toUpperCase()}`)) || 0,
        ]),
      ),
      behindDocument: anchor.getAttribute("behindDoc") === "1",
    }),
  );
}
