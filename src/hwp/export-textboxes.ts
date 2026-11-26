import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { nativeBoxLine, type NativeBoxLine } from '../hwp-floating-textbox';
import { elementRanges } from './markup-ranges';

export type Alignment = 'left' | 'center' | 'right' | 'justify';
export type NativeRun = { text: string; face: string; size: number; color: string; bold: boolean; italic: boolean; scale?: number; tracking?: number };
export type NativeParagraph = {
  runs: NativeRun[]; stored?: NativeBoxLine[]; align: Alignment; indent: number; left: number; right: number; before: number; after: number; pitch: number;
};
export type HwpExportTextBox = {
  marker: string; width: number; height: number; padding: [number, number, number, number];
  border?: string; stroke?: number; paragraphs: NativeParagraph[];
  fill?: string; cornerRadius?: number; before?: number;
  anchor?: { x: number; y: number; centered: boolean };
  nested?: { slot: number; child: HwpExportTextBox };
};

export type NativeBoxPaint = Map<string, { paragraphs: NativeParagraph[]; child: NonNullable<HwpExportTextBox['nested']> }>;
const toUnits = (value: number) => Math.round(value * 75);

function result(raw: string): {ok: boolean; controlIdx: number} {
  const value = JSON.parse(raw);
  if (!value.ok) throw new Error('글상자 또는 글상자 글을 한글 형식으로 만들지 못했습니다.');
  return value;
}

function nativeBorder(hex: string) {
  return Number.parseInt(hex.slice(5, 7) + hex.slice(3, 5) + hex.slice(1, 3), 16);
}

function restoreHwpExportTextBox(doc: HwpDocument, p: number, box: HwpExportTextBox) {
  const anchor = box.anchor, x = anchor?.centered ? 0 : anchor?.x ?? 0, y = anchor?.centered ? 0 : anchor?.y ?? 0;
  const { controlIdx: control } = result(doc.createShapeControl(JSON.stringify({sectionIdx:0, paraIdx:p,
    charOffset:0, width:toUnits(box.width), height:toUnits(box.height), horzOffset:toUnits(x),
    vertOffset:toUnits(y), treatAsChar:!box.anchor, textWrap:box.anchor ? 'InFrontOfText' : 'TopAndBottom'})));
  result(doc.setTextBoxAt(p, control, true));
  const [top, right, bottom, left] = box.padding.map(toUnits);
  const decoration = box.border ? { borderColor:nativeBorder(box.border), borderWidth:toUnits(box.stroke!) } :
    { borderWidth:33, lineType:0 };
  const fill = box.fill ? { fillType:'solid', fillBgColor:nativeBorder(box.fill) } : { fillType:'none' };
  result(doc.setShapeProperties(0, p, control, JSON.stringify({
    ...decoration, ...fill, tbMarginLeft:left, tbMarginRight:right, tbMarginTop:top, tbMarginBottom:bottom, tbVerticalAlign:'Center',
    ...(anchor ? { horzRelTo:'Paper', vertRelTo:'Paper', horzAlign:anchor.centered ? 'Center' : 'Left',
      vertAlign:anchor.centered ? 'Center' : 'Top', horzOffset:toUnits(x), vertOffset:toUnits(y), allowOverlap:true } : {}),
    ...(box.before !== undefined ? { outerMarginTop:toUnits(box.before) } : {}),
    ...(box.cornerRadius !== undefined ? { roundRate:Math.round(box.cornerRadius * 200 / Math.min(box.width, box.height)) } : {}),
  })));
  const faces = new Map<string, number[]>();
  for (const [index, paragraph] of box.paragraphs.entries()) {
    const literal = paragraph.runs.map(run => run.text).join('');
    if (literal) result(doc.insertTextInCell(0, p, control, 0, index, 0, literal));
    let offset = 0;
    for (const run of paragraph.runs) {
      let fontIds = faces.get(run.face);
      if (!fontIds) {
        fontIds = Array.from({length:7}, (_, lang) => doc.findOrCreateFontIdForLang(lang, run.face));
        faces.set(run.face, fontIds);
      }
      result(doc.applyCharFormatInCell(0, p, control, 0, index, offset, offset + run.text.length,
        JSON.stringify({fontIds, fontSize:toUnits(run.size), textColor:run.color, bold:run.bold, italic:run.italic,
          ...(run.scale !== undefined ? { ratios:Array(7).fill(run.scale) } : {}),
          ...(run.tracking !== undefined ? { spacings:Array(7).fill(run.tracking) } : {}),
        })));
      offset += run.text.length;
    }
    const baseSize = Math.max(...paragraph.runs.map(run => run.size));
    result(doc.applyParaFormatInCell(0, p, control, 0, index, JSON.stringify({alignment:paragraph.align,
      lineSpacingType:'Percent', lineSpacing:Math.round(paragraph.pitch * 100 / baseSize),
      marginLeft:toUnits(paragraph.left * 2), marginRight:toUnits(paragraph.right * 2), indent:toUnits(paragraph.indent * 2),
      spacingBefore:toUnits(paragraph.before * 2), spacingAfter:toUnits(paragraph.after * 2)})));
    if (index + 1 < box.paragraphs.length) result(doc.splitParagraphInCell(0, p, control, 0, index, offset));
    if (doc.getTextInCell(0, p, control, 0, index, 0, 0xffffffff) !== literal)
      throw new Error('한글 글상자의 글 또는 빈 문단을 확인하지 못했습니다.');
  }
  if (doc.getCellParagraphCount(0, p, control, 0) !== box.paragraphs.length)
    throw new Error('한글 글상자의 문단 수를 확인하지 못했습니다.');
}

function storyText(paragraph: NativeParagraph) {
  return paragraph.runs.map(run => run.text).join('');
}

/** The engine can create both editable controls, but exposes no setter for
 * nesting a shape. Move the generated child between generated HWPX stories
 * before section splitting and reopen the package in the native engine. */
function nativeBoxFiles(doc: HwpDocument) {
  let inflated = 0;
  return unzipSync(doc.exportHwpx(), { filter(file) {
    if ((inflated += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 글상자 변환 데이터가 너무 큽니다.');
    return true;
  } });
}

function nestExportTextBoxes(Core: typeof HwpDocument, doc: HwpDocument, boxes: readonly HwpExportTextBox[],
  locations: ReadonlyMap<string, number>, paint: NativeBoxPaint) {
  const files = nativeBoxFiles(doc);
  const section = 'Contents/section0.xml';
  const xml = strFromU8(files[section]);
  const paragraphs = elementRanges(xml, 'hp:p');
  const changes: { start: number; end: number; value: string }[] = [];
  const removed: number[] = [];
  for (const box of boxes) {
    if (!box.nested) continue;
    const { child, slot } = box.nested;
    const outerIndex = locations.get(box.marker)!, childIndex = locations.get(child.marker)!;
    const outer = paragraphs[outerIndex], inner = paragraphs[childIndex];
    const outerShapes = elementRanges(outer?.inner ?? '', 'hp:rect');
    const childShapes = elementRanges(inner?.inner ?? '', 'hp:rect');
    if (childIndex !== outerIndex + 1 || outerShapes.length !== 1 || childShapes.length !== 1 ||
        /<hp:(?:secPr|ctrl|pic|tbl)\b/.test(inner.inner) || /<hp:t>[^<]+<\/hp:t>/.test(inner.inner.slice(0, childShapes[0].start)))
      throw Error('중첩 글상자의 한글 위치가 예상과 다릅니다.');
    const outerShape = outerShapes[0], childShape = childShapes[0];
    const id = /\bid="(\d+)"/.exec(outerShape.open)?.[1];
    if (!id || paint.has(id)) throw Error('중첩 글상자의 한글 개체 ID를 확인하지 못했습니다.');
    paint.set(id, { paragraphs:box.paragraphs, child:box.nested });
    let outerXml = outer.inner.slice(outerShape.start, outerShape.end);
    const childXml = inner.inner.slice(childShape.start, childShape.end);
    // The engine's no-fill property still serializes a white brush.
    const fill = elementRanges(outerXml, 'hc:fillBrush'), draws = elementRanges(outerXml, 'hp:drawText');
    if (draws.length !== 1 || fill.length > 1 || fill.length && fill[0].end > draws[0].start && fill[0].start < draws[0].end)
      throw Error('중첩 글상자의 한글 채우기 또는 본문을 확인하지 못했습니다.');
    if (fill[0]) outerXml = outerXml.slice(0, fill[0].start) + outerXml.slice(fill[0].end);
    const draw = elementRanges(outerXml, 'hp:drawText')[0];
    const slots = elementRanges(draw.inner, 'hp:p');
    const host = slots[slot], runs = elementRanges(host?.inner ?? '', 'hp:run');
    if (slots.length !== box.paragraphs.length || !host || runs.length !== 1 || /<hp:t>[^<]+<\/hp:t>/.test(host.inner))
      throw Error('중첩 글상자의 빈 칸을 확인하지 못했습니다.');
    const at = draw.start + draw.open.length + host.start + host.open.length + runs[0].start + runs[0].open.length;
    const start = outer.start + outer.open.length + outerShape.start;
    changes.push({ start, end:start + outerShape.end - outerShape.start, value:outerXml.slice(0, at) + childXml + outerXml.slice(at) });
    changes.push({ start:inner.start, end:inner.end, value:'' });
    removed.push(childIndex);
  }
  let updated = xml;
  for (const { start, end, value } of changes.sort((a, b) => b.start - a.start))
    updated = updated.slice(0, start) + value + updated.slice(end);
  files[section] = strToU8(updated);
  const output = new Core(zipSync(files, { level:1 }));
  try {
    const controls = JSON.parse(output.getControls()) as
      { ctrlId:string; list:number; para:number; controlIndex:number; props?: { Width:number; Height:number } }[];
    for (const box of boxes) {
      if (!box.nested) continue;
      const p = locations.get(box.marker)!;
      const paragraph = p - removed.filter(index => index < p).length;
      // A reload inserts the section and column controls ahead of the
      // first paragraph's rectangle, changing its control index.
      const matches = controls.filter(shape => shape.ctrlId === 'gso' && shape.list === 0 && shape.para === paragraph &&
        shape.props?.Width === toUnits(box.width) && shape.props?.Height === toUnits(box.height));
      if (matches.length !== 1) throw Error('저장한 중첩 글상자의 위치를 다시 찾지 못했습니다.');
      const control = matches[0].controlIndex;
      const { child, slot } = box.nested;
      if (output.getCellParagraphCount(0, paragraph, control, 0) !== box.paragraphs.length ||
          box.paragraphs.some((story, index) => index !== slot &&
            output.getTextInCell(0, paragraph, control, 0, index, 0, 0xffffffff) !== storyText(story)) ||
          child.paragraphs.some((story, index) => output.getTextInCellByPath(0, paragraph, JSON.stringify([
            { controlIndex:control, cellIndex:0, cellParaIndex:slot }, { controlIndex:0, cellIndex:0, cellParaIndex:index },
          ]), 0, 0xffffffff) !== storyText(story))) throw Error('중첩 글상자의 한글 내용을 재열기에서 확인하지 못했습니다.');
    }
    return output;
  } catch (error) {
    output.free();
    throw error;
  }
}

export function restoreHwpExportTextBoxes(Core: typeof HwpDocument, doc: HwpDocument,
  boxes: readonly HwpExportTextBox[], paint: NativeBoxPaint) {
  if (!boxes.length) return doc;
  const all = boxes.flatMap(box => box.nested ? [box, box.nested.child] : [box]);
  const pending = new Map(all.map(box => [box.marker, box]));
  if (pending.size !== all.length) throw new Error('한글 글상자 위치가 중복되었습니다.');
  const locations = new Map<string, number>();
  const markerLength = all[0].marker.length;
  for (let p = 0, count = doc.getParagraphCount(0); p < count && pending.size; p++) {
    const marker = doc.getTextRange(0, p, 0, markerLength + 1);
    const box = pending.get(marker);
    if (!box) continue;
    doc.deleteText(0, p, 0, marker.length);
    restoreHwpExportTextBox(doc, p, box);
    locations.set(marker, p);
    pending.delete(marker);
  }
  if (pending.size) throw new Error('한글 글상자의 본문 위치를 확인하지 못했습니다.');
  return boxes.some(box => box.nested) ? nestExportTextBoxes(Core, doc, boxes, locations, paint) : doc;
}

function nativeBoxLines(story: string, lines: NativeBoxLine[] | undefined) {
  if (!lines?.length || lines.length > 12 || !lines.every(nativeBoxLine))
    throw Error('한글 글상자의 원래 줄 위치를 확인하지 못했습니다.');
  const xml = lines.map(([start, top, height, baseline, gap, width, flags]) =>
    `<hp:lineseg textpos="${start}" vertpos="${top}" vertsize="${height}" textheight="${height}" baseline="${baseline}" spacing="${gap}" horzpos="0" horzsize="${width}" flags="${flags}"/>`);
  return story + '<hp:linesegarray>' + xml.join('') + '</hp:linesegarray>';
}

function renderBoxStories(shape: string, paragraphs: NativeParagraph[], nested?: NonNullable<HwpExportTextBox['nested']>) {
  const draws = elementRanges(shape, 'hp:drawText');
  if (draws.length !== 1) throw Error('한글 글상자의 이야기 범위를 확인하지 못했습니다.');
  const draw = draws[0], slots = elementRanges(draw.inner, 'hp:p');
  if (slots.length !== paragraphs.length) throw Error('한글 글상자의 원래 문단 개수를 확인하지 못했습니다.');
  let content = draw.inner;
  for (let i = slots.length - 1; i >= 0; i--) {
    const slot = slots[i];
    let story = slot.inner;
    if (elementRanges(story, 'hp:linesegarray').length) throw Error('한글 글상자에 새로 계산된 줄 위치가 있습니다.');
    if (nested?.slot === i) {
      const rects = elementRanges(story, 'hp:rect');
      if (rects.length !== 1) throw Error('한글의 안쪽 사각형을 확인하지 못했습니다.');
      const rect = rects[0];
      story = story.slice(0, rect.start) + renderBoxStories(story.slice(rect.start, rect.end), nested.child.paragraphs) +
        story.slice(rect.end);
    }
    content = content.slice(0, slot.start) + slot.open + nativeBoxLines(story, paragraphs[i].stored) + '</hp:p>' + content.slice(slot.end);
  }
  return shape.slice(0, draw.start) + draw.open + content + '</hp:drawText>' + shape.slice(draw.end);
}

/** Other writer stages recalculate page controls and clear stale line boxes.
 * Reapply only the exact line boxes bound to the generated nested objects. */
export function restoreNativeBoxPaint(Core: typeof HwpDocument, doc: HwpDocument, paint: NativeBoxPaint) {
  if (!paint.size) return doc;
  const files = nativeBoxFiles(doc);
  const matched = new Set<string>();
  for (const key of Object.keys(files).filter(name => /^Contents\/section\d+\.xml$/.test(name))) {
    const original = strFromU8(files[key]);
    let xml = original;
    for (const shape of elementRanges(original, 'hp:rect').reverse()) {
      const id = /\bid="(\d+)"/.exec(shape.open)?.[1], box = id && paint.get(id);
      if (!box || !id) continue;
      if (matched.has(id)) throw Error('한글 글상자의 저장 ID가 중복되었습니다.');
      matched.add(id);
      xml = xml.slice(0, shape.start) + renderBoxStories(xml.slice(shape.start, shape.end), box.paragraphs, box.child) +
        xml.slice(shape.end);
    }
    if (xml !== original) files[key] = strToU8(xml);
  }
  if (matched.size !== paint.size) throw Error('한글 글상자의 원래 줄 위치를 다시 찾지 못했습니다.');
  return new Core(zipSync(files, { level:1 }));
}
