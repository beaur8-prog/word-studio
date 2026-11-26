import { paragraphLineOffsets } from './paragraph-fragments';
import { wordFontMetricsChanged } from './word-font-fallbacks';

const detail = 16;
const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
let canvas: (CanvasRenderingContext2D & { lang?: string }) | null | undefined;

type Contour = { ascent: number; descent: number };
const contours = new Map<string, Contour>();
const maxContours = 8192;

if (typeof document !== 'undefined') {
  const clear = () => contours.clear();
  document.fonts.addEventListener('loading', clear);
  document.fonts.addEventListener('loadingdone', clear);
  document.fonts.addEventListener('loadingerror', clear);
  document.addEventListener(wordFontMetricsChanged, clear);
}

/** The contour is independent of the DOM line's position. Its baseline
 * below is still measured on every rendered line. */
function measureContour(context: CanvasRenderingContext2D, profile: string | undefined, text: string): Contour {
  const key = profile && profile + '\0' + text;
  const cached = key && contours.get(key);
  if (cached) return cached;
  const shape = context.measureText(text);
  const result = { ascent: shape.actualBoundingBoxAscent, descent: shape.actualBoundingBoxDescent };
  if (key && Number.isFinite(result.ascent) && Number.isFinite(result.descent)) {
    if (contours.size >= maxContours) contours.delete(contours.keys().next().value!);
    contours.set(key, result);
  }
  return result;
}

function textBoxes(range: Range, leaf: Text, ignoreEdgeSpaces: boolean) {
  if (ignoreEdgeSpaces) {
    range.setStart(leaf, leaf.length - leaf.data.trimStart().length);
    range.setEnd(leaf, leaf.data.trimEnd().length);
  } else range.selectNodeContents(leaf);
  return [...range.getClientRects()].filter(box => box.width > 0 && box.height > 0);
}

const boundsKey = (box: DOMRect) => [box.left, box.top, box.width, box.height].join(':');

/** Use the laid-out baseline and the language of each DOM run. Canvas's
 * inherited language otherwise shapes some CJK punctuation differently. */
export function measuredHwpLineInk(line: HTMLElement, scale: number, maxFont: number, ignoreEdgeSpaces = false) {
  return inkLines(line, scale, maxFont, ignoreEdgeSpaces, true)?.[0] ?? null;
}

/** Each text node must already fit one visual line. Callers can temporarily
 * split nodes at natural boundaries without changing the authored paragraph. */
export function measuredHwpInkLines(paragraph: HTMLElement, scale: number, maxFont: number) {
  return inkLines(paragraph, scale, maxFont, false, false);
}

/** Measure the visible contours of natural lines without changing their layout. */
export function measuredHwpInkAtLineOffsets(paragraph: HTMLElement, offsets: number[], size: number,
  calibrateBaseline = false) {
  const restore: (() => void)[] = [];
  try {
    const content = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    const splits: { node: Text; positions: number[] }[] = [];
    let start = 0;
    while (content.nextNode()) {
      const node = content.currentNode as Text, end = start + node.length;
      const positions = offsets.filter(offset => offset > start && offset < end).map(offset => offset - start);
      if (positions.length) splits.push({ node, positions });
      start = end;
    }
    for (const { node, positions } of splits) {
      const data = node.data;
      const pieces = positions.reverse().map(position => node.splitText(position));
      restore.push(() => { pieces.forEach(piece => piece.remove()); node.data = data; });
    }
    const after = paragraphLineOffsets(paragraph);
    if (offsets.length !== after.length || offsets.some((offset, i) => offset !== after[i])) return;
    return inkLines(paragraph, 1, size, false, false, calibrateBaseline);
  } finally { restore.forEach(undo => undo()); }
}

const baselineFontProperties = [
  'direction', 'writing-mode', 'font-family', 'font-size', 'font-style', 'font-weight', 'font-stretch',
  'font-variant-caps', 'font-feature-settings', 'font-kerning', 'font-optical-sizing', 'font-variation-settings',
  'font-synthesis', 'letter-spacing', 'word-spacing', 'line-height',
];

/** Relate a rendered text rectangle to its baseline without adding an
 * inline box to the live line, where even a zero-width box may wrap. */
function isolatedBaseline(leaf: Text, original: DOMRect, probe: HTMLElement, ignoreEdgeSpaces: boolean) {
  const host = document.createElement('div'), mark = leaf.parentElement!;
  host.style.cssText = 'position:fixed;left:-100000px;top:0;visibility:hidden;width:max-content;white-space:pre;line-height:normal;margin:0;padding:0;border:0';
  host.lang = mark.closest<HTMLElement>('[lang]')?.lang || document.documentElement.lang;
  const copy = document.createElement('span'), css = getComputedStyle(mark);
  for (const property of baselineFontProperties) copy.style.setProperty(property, css.getPropertyValue(property));
  const text = document.createTextNode(leaf.data), baseline = probe.cloneNode(false) as HTMLElement;
  copy.append(text, baseline);
  host.append(copy);
  document.body.append(host);
  try {
    const [sample, extra] = textBoxes(document.createRange(), text, ignoreEdgeSpaces);
    if (sample && !extra && sample.height === original.height)
      return original.top + baseline.getBoundingClientRect().bottom - sample.top;
  } finally { host.remove(); }
}

function inkLines(line: HTMLElement, scale: number, maxFont: number, ignoreEdgeSpaces: boolean, single: boolean,
  calibrateBaseline = false) {
  if (!(scale > 0) || !line.textContent?.trim() || line.textContent.length > 250) return null;
  canvas ??= document.createElement('canvas').getContext('2d');
  if (!canvas || typeof canvas.lang !== 'string') return null;
  const range = document.createRange(), walker = document.createTreeWalker(line, NodeFilter.SHOW_TEXT);
  const runs: { leaf: Text; box: string; probe: HTMLElement; ascent: number; descent: number }[] = [];
  while (walker.nextNode()) {
    const leaf = walker.currentNode as Text;
    if (!leaf.data.trim()) continue;
    const element = leaf.parentElement!, css = getComputedStyle(element), size = parseFloat(css.fontSize);
    if (!(size > 0 && size <= maxFont + .06) || css.verticalAlign !== 'baseline' || css.direction !== 'ltr' ||
        css.writingMode !== 'horizontal-tb' || css.transform !== 'none' || css.textShadow !== 'none' ||
        css.textDecorationLine !== 'none' || css.textTransform !== 'none' || css.fontVariantCaps !== 'normal' ||
        css.fontVariationSettings !== 'normal' || css.fontStretch !== '100%' ||
        parseFloat(css.getPropertyValue('-webkit-text-stroke-width')) > 0) return null;
    const boxes = textBoxes(range, leaf, ignoreEdgeSpaces);
    if (boxes.length !== 1) return null;
    const lang = element.closest<HTMLElement>('[lang]')?.lang || document.documentElement.lang;
    if (!lang) return null;
    canvas.lang = lang;
    canvas.font = `${css.fontStyle} ${css.fontWeight} ${size * detail}px ${css.fontFamily}`;
    const profile = document.fonts.status === 'loaded' ? canvas.font + '\0' + canvas.lang : undefined;
    let { ascent, descent } = measureContour(canvas, profile, leaf.data);
    for (const { segment } of graphemes.segment(leaf.data)) {
      if (!segment.trim()) continue;
      const shape = measureContour(canvas, profile, segment);
      ascent = Math.max(ascent, shape.ascent);
      descent = Math.max(descent, shape.descent);
    }
    if (!Number.isFinite(ascent) || !Number.isFinite(descent)) return null;
    const probe = document.createElement('span');
    probe.style.cssText = 'display:inline-block;vertical-align:baseline;width:0;height:0;font-size:0;padding:0;border:0;margin:0';
    probe.style.setProperty('line-height', '0px', 'important');
    runs.push({ leaf, box: boundsKey(boxes[0]), probe,
      ascent: (ascent + 1) * scale / detail, descent: (descent + 1) * scale / detail });
  }
  if (!runs.length) return null;
  const frame = boundsKey(line.getBoundingClientRect());
  const directBaselines = () => {
    try {
      for (const { leaf, probe } of runs) leaf.after(probe);
      if (boundsKey(line.getBoundingClientRect()) !== frame) return;
      const positions: number[] = [];
      for (const { leaf, box, probe } of runs) {
        const rectangles = textBoxes(range, leaf, ignoreEdgeSpaces);
        if (rectangles.length !== 1 || boundsKey(rectangles[0]) !== box) return;
        positions.push(probe.getBoundingClientRect().bottom);
      }
      return positions;
    } finally {
      for (const { probe } of runs) probe.remove();
    }
  };

  let positions = directBaselines();
  if (!positions && calibrateBaseline && boundsKey(line.getBoundingClientRect()) === frame) {
    const calibrated = runs.map(({ leaf, box, probe }) => {
      const [original, extra] = textBoxes(range, leaf, ignoreEdgeSpaces);
      if (original && !extra && boundsKey(original) === box)
        return isolatedBaseline(leaf, original, probe, ignoreEdgeSpaces);
    });
    if (calibrated.every(position => position !== undefined && Number.isFinite(position)))
      positions = calibrated as number[];
  }
  if (!positions) return null;
  const result: { top: number; bottom: number; baseline: number }[] = [];
  for (const [index, run] of runs.entries()) {
    const y = positions[index], prior = result.at(-1);
    if (prior && Math.abs(y - prior.baseline) <= .05 * scale) {
      prior.baseline = y;
      prior.top = Math.min(prior.top, y - run.ascent);
      prior.bottom = Math.max(prior.bottom, y + run.descent);
    } else {
      if (prior && (single || y < prior.baseline)) return null;
      result.push({ top: y - run.ascent, bottom: y + run.descent, baseline: y });
    }
  }
  return result;
}

/** Font boxes retained by a print-only wrap at the same physical height. */
export function retainedTextBands(paragraph: HTMLElement, range: Range, bounds: Pick<DOMRect, 'top' | 'bottom'>) {
  const bands = new Set<string>();
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    if (!walker.currentNode.textContent?.trim()) continue;
    range.selectNodeContents(walker.currentNode);
    for (const rect of range.getClientRects()) {
      if (!(rect.width > 0 && rect.height > 0)) continue;
      if (rect.top < bounds.top || rect.bottom > bounds.bottom) return;
      bands.add(rect.top + ':' + rect.bottom);
    }
  }
  if (bands.size) return bands;
}
