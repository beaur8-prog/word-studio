import { characterScaleProperty } from '../character-scale';
import { floatingTextBox, type HwpFloatingTextBox, type NativeBoxLine } from '../hwp-floating-textbox';
import type { HwpDocument } from '@rhwp/core';
import { hwpSourceWrapAttribute, hwpWrapFingerprint, normalizeHwpWrapText } from '../hwp-text-advances';
import type { TextBoxFrame } from '../textbox-frame';
import { paragraphAutoSpace, paragraphCss } from './paragraph-layout';

type Bounds = { x: number; y: number; width: number; height: number };
type NativeShape = {
  width: number; height: number; treatAsChar: boolean; rotationAngle: number; horzFlip: boolean; vertFlip: boolean;
  fillType: string; shadowType: number; roundRate: number; hasCaption: boolean; textWrap: string;
  vertRelTo?: string; vertAlign?: string; vertOffset?: number; horzRelTo?: string; horzAlign?: string; horzOffset?: number;
  outerMarginLeft: number; outerMarginRight: number; outerMarginTop: number; outerMarginBottom: number;
  tbMarginLeft: number; tbMarginRight: number; tbMarginTop: number; tbMarginBottom: number; tbVerticalAlign: string;
};
type Op = {
  type: string; bbox?: Bounds; text?: string; displayText?: string;
  style?: { fillColor: string | null; strokeColor: string; strokeWidth: number; strokeDash: string; opacity: number };
  transform?: { rotation?: number; horzFlip?: boolean; vertFlip?: boolean };
  rotation?: number; isVertical?: boolean; cornerRadius?: number;
};
type Layer = { kind: string; bounds: Bounds; ops?: Op[]; children?: Layer[]; child?: Layer };
type TextLine = { textStart: number; verticalPosition: number; width: number; height: number; spacing: number; baseline: number; textHeight?: number; flags?: number };
type NativeBand = { x: number; y: number; text: string };
type CharacterShape = {
  fontFamily: string; fontFamilies: string[]; fontSize: number; textColor: string; shadeColor: string;
  ratios: number[]; spacings: number[]; relativeSizes: number[]; charOffsets: number[];
  bold: boolean; italic: boolean; underline: boolean; strikethrough: boolean; outlineType: number; shadowType: number;
  superscript: boolean; subscript: boolean; emboss: boolean; engrave: boolean; emphasisDot: number;
  fillType: string; kerning: boolean; borderLeft: { type: number }; borderRight: { type: number };
  borderTop: { type: number }; borderBottom: { type: number };
};
const close = (a: number, b: number, precision = .16) => Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) <= precision;
const escapeHtml = (text: string) => text.replace(/[&<>"']/g, ch =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[ch]!);
// Keep the semicolons used by the outer HTML bridge as CSS delimiters.
const cssFamily = (face: string) => "'" + face.replace(/[\x00-\x1f\x7f\\'"&<>;]/gu, char =>
  '\\' + char.codePointAt(0)!.toString(16) + ' ') + "'";
const rgb = /^#[0-9a-f]{6}$/i;
const ordinaryText = /[\u0000-\u001f\u0590-\u08ff\u202a-\u202e\u2066-\u2069\uD800-\uDFFF\ufffc]/u;

function plainCharacter(shape: CharacterShape, allowAuthoredTracking = false) {
  const uniform = <T>(values: T[], one: T) => Array.isArray(values) && values.length === 7 && values.every(value => value === one);
  return shape.fontFamily && shape.fontFamily.length <= 100 && !/[\t\r\n\0]/.test(shape.fontFamily) &&
    Number.isSafeInteger(shape.fontSize) && shape.fontSize >= 400 && shape.fontSize <= 7200 &&
    rgb.test(shape.textColor) && shape.shadeColor.toLowerCase() === '#ffffff' &&
    uniform(shape.fontFamilies, shape.fontFamily) &&
    (allowAuthoredTracking ? uniform(shape.ratios, shape.ratios[0]) && shape.ratios[0] >= 70 && shape.ratios[0] <= 100 &&
      uniform(shape.spacings, shape.spacings[0]) && shape.spacings[0] >= -30 && shape.spacings[0] <= 40
      : uniform(shape.ratios, 100) && uniform(shape.spacings, 0)) &&
    uniform(shape.relativeSizes, 100) && uniform(shape.charOffsets, 0) &&
    !((!allowAuthoredTracking && (shape.bold || shape.italic)) || shape.underline || shape.strikethrough || shape.outlineType || shape.shadowType ||
      shape.superscript || shape.subscript || shape.emboss || shape.engrave || shape.emphasisDot || shape.kerning) &&
    shape.fillType === 'none' && ![shape.borderLeft, shape.borderRight, shape.borderTop, shape.borderBottom].some(b => b.type);
}

function shapeOps(layer: Layer): Op[] | undefined {
  const ops: Op[] = [];
  const visit = (part: Layer, depth: number): boolean => {
    if (depth > 16 || ops.length > 256) return false;
    if (part.kind === 'group') return !!part.children?.length && part.children.every(child => visit(child, depth + 1));
    if (part.kind === 'clipRect' && part.child) return visit(part.child, depth + 1);
    if (part.kind !== 'leaf' || !part.ops?.length || ops.length + part.ops.length > 256) return false;
    ops.push(...part.ops);
    return true;
  };
  return visit(layer, 0) ? ops : undefined;
}

function singleOutline(layer: Layer): { outline: Op; runs: Op[] } | undefined {
  const ops = shapeOps(layer);
  if (!ops) return;
  const outlines = ops.filter(op => op.type !== 'textRun');
  const [outline] = outlines, b = outline?.bbox, style = outline?.style;
  if (outlines.length !== 1 || outline.type !== 'rectangle' || !b || !style ||
      style.fillColor !== null || !rgb.test(style.strokeColor) || !(style.strokeWidth > 0 && style.strokeWidth <= 3) ||
      style.strokeDash !== 'solid' || style.opacity !== 1 || outline.transform?.rotation ||
      outline.transform?.horzFlip || outline.transform?.vertFlip ||
      !close(layer.bounds.x, b.x) || !close(layer.bounds.y, b.y) ||
      !close(layer.bounds.width, b.width) || !close(layer.bounds.height, b.height)) return;
  const runs = ops.filter(op => op.type === 'textRun');
  if (!runs.length || runs.some(run => !run.bbox || run.rotation || run.isVertical)) return;
  return { outline, runs };
}

export function sourceInlineTextBox(doc: HwpDocument, section: number, paragraph: number, control: number,
  page: number, layer: Layer, shape: NativeShape): string | undefined {
  if (!shape.treatAsChar || shape.rotationAngle || shape.horzFlip || shape.vertFlip || shape.fillType !== 'none' ||
      shape.shadowType || shape.roundRate || shape.hasCaption || shape.textWrap !== 'TopAndBottom' ||
      shape.outerMarginTop || shape.outerMarginRight || shape.outerMarginBottom || shape.outerMarginLeft ||
      shape.tbVerticalAlign !== 'Center') return;
  const drawn = singleOutline(layer);
  if (!drawn) return;
  const bounds = layer.bounds, outline = drawn.outline.style!;
  const padding: TextBoxFrame['padding'] = [shape.tbMarginTop / 75, shape.tbMarginRight / 75,
    shape.tbMarginBottom / 75, shape.tbMarginLeft / 75];
  const paper = JSON.parse(doc.getPageInfo(page));
  const bodyWidth = paper.width - paper.marginLeft - paper.marginRight;
  const bodyTop = paper.marginTop + paper.marginHeader, bodyBottom = paper.footerArea?.y;
  const innerWidth = bounds.width - padding[1] - padding[3], innerHeight = bounds.height - padding[0] - padding[2];
  if (paper.columns?.length !== 1 || !(bounds.width >= 100 && bounds.width <= 2000 && bounds.height >= 80 && bounds.height <= 1800) ||
      !close(bounds.x, paper.marginLeft) || !close(bounds.width, bodyWidth) || !close(bounds.width, shape.width / 75) ||
      !close(bounds.height, shape.height / 75) || bounds.y < bodyTop - .16 || bounds.y + bounds.height > bodyBottom + .16 ||
      !padding.every(value => Number.isFinite(value) && value >= 0) || innerWidth <= 0 || innerHeight <= 0) return;
  const count = doc.getCellParagraphCount(section, paragraph, control, 0);
  if (count < 2 || count > 16) return;
  const painted = new Map<number, NativeBand[]>();
  for (const run of JSON.parse(doc.getPageTextLayout(page)).runs ?? []) {
    const address = run.cellPath?.[0];
    if (run.secIdx !== section || run.parentParaIdx !== paragraph || run.cellPath?.length !== 1 ||
        address.controlIndex !== control || address.cellIndex !== 0) continue;
    const text = run.displayText ?? run.text ?? '';
    if (run.y < bounds.y - .16 || run.y + run.h > bounds.y + bounds.height + .16) return;
    const row = painted.get(address.cellParaIndex) ?? [];
    row.push({ x: run.x, y: run.y, text });
    painted.set(address.cellParaIndex, row);
  }
  const html: string[] = [];
  let height = 0, pitch = 0, lastY: number | undefined;
  let actualText = '';
  for (let index = 0; index < count; index++) {
    const address = JSON.stringify([{ controlIndex: control, cellIndex: 0, cellParaIndex: index }]);
    const text = doc.getTextInCellByPath(section, paragraph, address, 0, 0xffffffff);
    if (text.length > 1000 || ordinaryText.test(text)) return;
    actualText += text;
    const properties = JSON.parse(doc.getCellParaPropertiesAt(section, paragraph, control, 0, index));
    const char: CharacterShape = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, address, 0));
    if (!plainCharacter(char) || properties.headType !== 'None' || properties.borderConnect || properties.fontLineHeight ||
        properties.keepWithNext || properties.keepLines || properties.pageBreakBefore || properties.spacingBefore ||
        properties.spacingAfter || properties.alignment !== 'justify' || properties.lineSpacingType !== 'Percent' ||
        properties.koreanBreakUnit !== 1 || properties.indent > 0) return;
    if ([properties.borderLeft, properties.borderRight, properties.borderTop, properties.borderBottom].some(b => b.type)) return;
    const metric = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, address));
    const lines: TextLine[] = metric.lines;
    if (!metric.stored || metric.controlCount || !lines?.length || lines.length > 12 ||
        lines[0].textStart !== 0 || !text && lines.length !== 1) return;
    const em = char.fontSize / 75, first = lines[0];
    const lineWidth = innerWidth - properties.marginLeft - properties.marginRight;
    if (!(lineWidth > 0) || !close(first.width, lineWidth)) return;
    pitch ||= first.height + first.spacing;
    if (!(pitch >= em && pitch <= em * 3) ||
       lines.some((line, i) => !close(line.height, em) || !close(line.width, lineWidth) ||
         !close(line.height + line.spacing, pitch) || i && (line.textStart <= lines[i - 1].textStart ||
           line.textStart >= text.length || !close(line.verticalPosition - lines[i - 1].verticalPosition, pitch)))) return;
    const literal = normalizeHwpWrapText(text);
    const ends = [...lines.slice(1).map(line => normalizeHwpWrapText(text.slice(0, line.textStart)).length), literal.length];
    const bands = new Map<number, NativeBand[]>();
    for (const run of painted.get(index) ?? []) {
      const list = bands.get(run.y) ?? [];
      list.push(run); bands.set(run.y, list);
    }
    if (bands.size !== lines.length) return;
    let cumulative = 0;
    const recorded = [...bands].sort(([a], [b]) => a - b).map(([y, group]) => {
      if (lastY !== undefined && !close(y - lastY, pitch, .2)) return '';
      lastY = y;
      return group.sort((a, b) => a.x - b.x).map(run => run.text).join('');
    });
    if (normalizeHwpWrapText(recorded.join('')) !== literal ||
        recorded.some((line, i) => (cumulative += normalizeHwpWrapText(line).length) !== ends[i])) return;
    let spans = '', offset = 0;
    const runs = JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, address, 0, text.length));
    for (const run of runs) {
      if (run.startOffset !== offset || run.endOffset <= offset || run.endOffset > text.length) return;
      const style: CharacterShape = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, address, offset));
      if (!plainCharacter(style) || style.fontFamily !== char.fontFamily || style.fontSize !== char.fontSize) return;
      const font = cssFamily(style.fontFamily);
      spans += `<span style="font-family:${font};font-size:${style.fontSize / 100}pt;color:${style.textColor}">` +
        escapeHtml(text.slice(offset, run.endOffset)) + '</span>';
      offset = run.endOffset;
    }
    if (offset !== text.length) return;
    height += lines.length * pitch;
    const proof = literal ? ' ' + hwpSourceWrapAttribute + '="' + escapeHtml(JSON.stringify({
      fingerprint: hwpWrapFingerprint(text), ends, inlineTextBox: {
        width: bounds.width, height: bounds.height, textWidth: lineWidth, pitch,
        hanging: -properties.indent, size: em, family: char.fontFamily,
      },
    })) + '"' : '';
    html.push(paragraphAutoSpace('<p' + proof + ' style="' + paragraphCss(properties) +
      'font-size:' + char.fontSize / 100 + 'pt;font-family:' + cssFamily(char.fontFamily) +
      ';line-height:' + pitch.toFixed(4) + 'px;">', properties) +
      spans + '</p>');
  }
  if (!actualText.trim() || height > innerHeight ||
      normalizeHwpWrapText(drawn.runs.map(run => run.displayText ?? run.text ?? '').join('')) !== normalizeHwpWrapText(actualText)) return;
  const frame: TextBoxFrame = { width: bounds.width, height: bounds.height, x: 0, y: 0,
    position: 'inline', vertical: 'paragraph', shape: 'rect', padding, align: 'center',
    border: outline.strokeColor, borderWidth: outline.strokeWidth };
  return '<div data-document-textbox="' + escapeHtml(JSON.stringify(frame)) + '"><div data-textbox-content="true">' + html.join('') + '</div></div>';
}

type BoxAddress = { controlIndex: number; cellIndex: number; cellParaIndex: number };
type BoxBand = NativeBand & { h: number; w: number };
function sameBounds(a: Bounds, b: Bounds) {
  return close(a.x, b.x) && close(a.y, b.y) && close(a.width, b.width) && close(a.height, b.height);
}
function withinBox(box: Bounds, frame: Bounds) {
  return box.x >= frame.x - .16 && box.y >= frame.y - .16 &&
    box.x + box.width <= frame.x + frame.width + .16 && box.y + box.height <= frame.y + frame.height + .16;
}
function boxHtml(frame: TextBoxFrame, content: string) {
  return '<div data-document-textbox="' + escapeHtml(JSON.stringify(frame)) + '"><div data-textbox-content="true">' + content + '</div></div>';
}

/** An ordinary paper-centred rectangle may contain one nested inline
 * rectangle. The actual drawing and native paths must account for all paint. */
export function sourceFloatingNestedTextBox(doc: HwpDocument, section: number, paragraph: number,
  control: number, page: number, layer: Layer, shape: NativeShape): string | undefined {
  if (shape.treatAsChar || shape.rotationAngle || shape.horzFlip || shape.vertFlip || shape.fillType !== 'none' ||
      shape.shadowType || shape.roundRate || shape.hasCaption || shape.textWrap !== 'InFrontOfText' ||
      shape.vertRelTo !== 'Paper' || shape.horzRelTo !== 'Paper' || shape.vertAlign !== 'Center' ||
      shape.horzAlign !== 'Center' || shape.vertOffset || shape.horzOffset ||
      shape.outerMarginTop || shape.outerMarginRight || shape.outerMarginBottom || shape.outerMarginLeft ||
      shape.tbVerticalAlign !== 'Center') return;
  const bounds = layer.bounds, ops = shapeOps(layer), paper = JSON.parse(doc.getPageInfo(page));
  if (!ops || paper.columns?.length !== 1 || bounds.width < 100 || bounds.width > 2000 ||
      bounds.height < 100 || bounds.height > 1600 || !close(bounds.x * 2 + bounds.width, paper.width) ||
      !close(bounds.y * 2 + bounds.height, paper.height) ||
      !close(bounds.width, shape.width / 75) || !close(bounds.height, shape.height / 75)) return;
  const shapes = ops.filter(op => op.type !== 'textRun');
  if (shapes.length !== 2 || shapes.some(op => op.type !== 'rectangle' || !op.bbox || !op.style ||
      op.style.strokeColor || op.style.strokeWidth || op.style.opacity !== 1 || op.transform?.rotation ||
      op.transform?.horzFlip || op.transform?.vertFlip)) return;
  const outer = shapes.find(op => sameBounds(op.bbox!, bounds)), inner = shapes.find(op => op !== outer);
  const child = inner?.bbox, fill = inner?.style?.fillColor, radius = inner?.cornerRadius ?? 0;
  if (!outer || outer.cornerRadius || outer.style!.fillColor !== null || !child || !fill || !rgb.test(fill) ||
      !Number.isFinite(radius) || radius < 0 || radius > Math.min(child.width, child.height) / 2 ||
      !withinBox(child, bounds) || !close((bounds.width - child.width) / 2, child.x - bounds.x) ||
      child.height < 36 || child.height >= bounds.height / 3 || child.width < 100) return;
  const count = doc.getCellParagraphCount(section, paragraph, control, 0);
  if (count < 3 || count > 16) return;
  const slot = (index: number): BoxAddress => ({ controlIndex: control, cellIndex: 0, cellParaIndex: index });
  const padding: TextBoxFrame['padding'] = [shape.tbMarginTop / 75, shape.tbMarginRight / 75,
    shape.tbMarginBottom / 75, shape.tbMarginLeft / 75];
  if (!padding.every(n => Number.isFinite(n) && n >= 0 && n <= 100)) return;
  const painted = new Map<string, BoxBand[]>();
  for (const run of JSON.parse(doc.getPageTextLayout(page)).runs ?? []) {
    if (run.secIdx !== section || run.parentParaIdx !== paragraph || run.cellPath?.[0]?.controlIndex !== control) continue;
    const path = run.cellPath as BoxAddress[];
    if (![1, 2].includes(path.length) || path[0].cellIndex !== 0 || path[0].cellParaIndex < 0 ||
        path[0].cellParaIndex >= count || path.length === 2 &&
        (path[1].controlIndex !== 0 || path[1].cellIndex !== 0 || path[1].cellParaIndex !== 0)) return;
    const key = JSON.stringify(path), lines = painted.get(key) ?? [];
    const text = run.displayText ?? run.text ?? '';
    if (text.trim() && !withinBox({ x: run.x, y: run.y, width: run.w, height: run.h }, path.length === 2 ? child : bounds)) return;
    lines.push({ x: run.x, y: run.y, w: run.w, h: run.h, text });
    painted.set(key, lines);
  }
  const literal = (path: string) => doc.getTextInCellByPath(section, paragraph, path, 0, 0xffffffff);
  const metrics = (path: string): { stored: boolean; controlCount: number; lines: TextLine[] } =>
    JSON.parse(doc.getParagraphLineMetrics(section, paragraph, path));
  const properties = (path: string) => JSON.parse(doc.getCellParaPropertiesAtByPath(section, paragraph, path));
  const character = (path: string, offset: number): CharacterShape =>
    JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, path, offset));
  const stories: string[] = [];
  const nativeLine = (line: TextLine): NativeBoxLine => {
    const units = (n: number) => Math.round(n * 75);
    return [line.textStart, units(line.verticalPosition), units(line.height), units(line.baseline),
      units(line.spacing), units(line.width), line.flags!];
  };
  let nested = false;
  function story(path: string, props: Record<string, any>, info: ReturnType<typeof metrics>,
    frame: Bounds, order: number, placement?: HwpFloatingTextBox['child']) {
    const text = literal(path), asci = normalizeHwpWrapText(text), char = character(path, 0);
    const lines = info.lines;
    if (!asci || text.length > 1000 || ordinaryText.test(text) || !plainCharacter(char, true) || !info.stored ||
        info.controlCount || !lines?.length || lines.length > 12 || lines[0].textStart ||
        props.headType !== 'None' || props.borderConnect || props.fontLineHeight || props.keepWithNext || props.keepLines ||
        props.pageBreakBefore || props.marginLeft || props.marginRight || props.spacingAfter || ![0, 1].includes(props.koreanBreakUnit) ||
        props.indent > 0 || props.lineSpacingType !== 'Percent' ||
        [props.borderLeft, props.borderRight, props.borderTop, props.borderBottom].some(border => border?.type)) return;
    const first = lines[0], pitch = first.height + first.spacing, size = char.fontSize / 75, scale = char.ratios[0];
    if (!close(first.height, size) || !close(pitch, size * props.lineSpacing / 100, .08) || lines.some((line, i) =>
        !close(line.height, size) || !close(line.width, first.width) || !close(line.height + line.spacing, pitch) ||
        i && (line.textStart <= lines[i - 1].textStart || line.textStart >= text.length ||
          !close(line.verticalPosition - lines[i - 1].verticalPosition, pitch)))) return;
    const bands = new Map<number, BoxBand[]>();
    for (const run of painted.get(path) ?? []) {
      const group = bands.get(run.y) ?? [];
      group.push(run);
      bands.set(run.y, group);
    }
    if (bands.size !== lines.length) return;
    const ends = [...lines.slice(1).map(line => normalizeHwpWrapText(text.slice(0, line.textStart)).length), asci.length];
    let cumulative = 0;
    const positioned = [...bands].sort(([a], [b]) => a - b);
    const matched = positioned.map(([y, group], i) => {
      if (!close(y - positioned[0][0], lines[i].verticalPosition - first.verticalPosition) ||
          !group.every(run => close(run.h, size))) return '';
      return group.sort((a, b) => a.x - b.x).map(run => run.text).join('');
    });
    if (normalizeHwpWrapText(matched.join('')) !== asci ||
        matched.some((line, i) => (cumulative += normalizeHwpWrapText(line).length) !== ends[i])) return;
    let offset = 0, spans = '';
    const faces = new Set([char.fontFamily]);
    for (const run of JSON.parse(doc.getCharShapeRunsInCellByPath(section, paragraph, path, 0, text.length))) {
      const style = character(path, offset);
      if (run.startOffset !== offset || run.endOffset <= offset || run.endOffset > text.length ||
          !plainCharacter(style, true) || style.fontSize !== char.fontSize || style.ratios[0] !== scale) return;
      faces.add(style.fontFamily);
      spans += '<span style="font-family:' + cssFamily(style.fontFamily) + ';font-size:' + style.fontSize / 100 +
        'pt;color:' + style.textColor + ';font-weight:' + (style.bold ? 700 : 400) + ';font-style:' +
        (style.italic ? 'italic' : 'normal') + ';letter-spacing:' + (style.spacings[0] * size / 100) + 'px;' +
        characterScaleProperty + ':' + scale + ';">' + escapeHtml(text.slice(offset, run.endOffset)) + '</span>';
      offset = run.endOffset;
    }
    if (offset !== text.length) return;
    if (lines.some(line => !close(line.height, line.textHeight ?? NaN))) return;
    const proof = floatingTextBox({ width: frame.width, height: frame.height, order, count, pitch, size, scale,
      hanging: -props.indent, alignment: props.alignment, faces: [...faces], stored: lines.map(nativeLine),
      baselines: positioned.map(([y], i) => y + lines[i].baseline - frame.y), child: placement }, lines.length);
    if (!proof) return;
    const tag = '<p ' + hwpSourceWrapAttribute + '="' + escapeHtml(JSON.stringify({
      fingerprint: hwpWrapFingerprint(text), ends, floatingTextBox: proof,
    })) + '" style="' + paragraphCss(props) + 'font-size:' + char.fontSize / 100 + 'pt;font-family:' +
      cssFamily(char.fontFamily) + ';line-height:' + pitch.toFixed(4) + 'px;">';
    return paragraphAutoSpace(tag, props) + spans + '</p>';
  }
  for (let order = 0; order < count; order++) {
    const address = slot(order), path = JSON.stringify([address]), props = properties(path), info = metrics(path);
    if (info.controlCount === 0) {
      if (!close(info.lines?.[0]?.width, bounds.width - padding[1] - padding[3], .25)) return;
      const html = story(path, props, info, bounds, order);
      if (!html) return;
      stories.push(html);
      continue;
    }
    if (nested || info.controlCount !== 1 || !info.stored || info.lines?.length !== 1 || literal(path).trim() ||
        props.alignment !== 'center' || props.indent || props.marginLeft || props.marginRight ||
        props.spacingAfter || props.spacingBefore < 0 || props.spacingBefore > 200) return;
    const childPath = JSON.stringify([address, { controlIndex: 0, cellIndex: 0, cellParaIndex: 0 }]);
    if (doc.getCellParagraphCountByPath(section, paragraph, childPath) !== 1) return;
    const sub = metrics(childPath), innerProps = properties(childPath);
    const band = info.lines[0].height + info.lines[0].spacing;
    const side = (child.width - (sub.lines?.[0]?.width ?? NaN)) / 2;
    if (!(side >= 0 && side < 30) || sub.lines?.length !== 1 || innerProps.alignment !== 'center' ||
        innerProps.indent || innerProps.spacingBefore || innerProps.spacingAfter) return;
    const html = story(childPath, innerProps, sub, child, order,
      { x: child.x - bounds.x, y: child.y - bounds.y, before: props.spacingBefore / 2, band, host: nativeLine(info.lines[0]) });
    if (!html) return;
    stories.push(boxHtml({ x: 0, y: 0, width: child.width, height: child.height, position: 'inline',
      vertical: 'paragraph', shape: radius ? 'roundRect' : 'rect', ...(radius ? { cornerRadius: radius } : {}),
      padding: [0, side, 0, side], align: 'center', fill }, html));
    nested = true;
  }
  const layerText = ops.filter(op => op.type === 'textRun'), nativeRuns = [...painted.values()].flat();
  // A child shape is painted after the outer story even when its anchor is
  // in the middle. Match the independent ordered paint runs before using
  // the stored paragraph order to construct the editable stories above.
  if (!nested || layerText.length !== nativeRuns.length || layerText.some((op, i) => {
    const run = nativeRuns[i], box = op.bbox;
    return !box || op.rotation || op.isVertical ||
      !sameBounds(box, { x: run.x, y: run.y, width: run.w, height: run.h }) ||
      normalizeHwpWrapText(op.displayText ?? op.text ?? '') !== normalizeHwpWrapText(run.text);
  })) return;
  return boxHtml({ ...bounds, position: 'floating', vertical: 'page', shape: 'rect', align: 'center', padding }, stories.join(''));
}
