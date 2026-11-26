import { characterScaleProperty, parseCharacterScale } from './character-scale';
import { type HwpBodyLine, parseHwpBodyLine } from './hwp-body-line';
import { hwpWrapFingerprint } from './hwp-text-advances';

const contentElements = ':scope > span';
const forbidden = 'a,br,img,svg,sup,sub,mark,u,s,del,[data-flow-widget],[data-page-number],[data-docx-tab],[data-note-reference],[data-document-field]';
const unsupportedText = /[\u0000-\u001f\u0590-\u08ff\u202a-\u202e\u2066-\u2069]/u;
let canvas: CanvasRenderingContext2D | null | undefined;
const pixelGuard = .5;
const lineTolerance = .08;
const reported = new Set<string>();

export type SourceColumnPlacement = { height: number; offset: number; width?: string; indent?: string; scale?: number };
type Paint = { ascent: number; descent: number };

const provenText = (p: Element, fingerprint: string) =>
  p.tagName === 'P' && !p.matches('[data-print-paragraph-fragment],[data-paragraph-border-box],[data-word-paragraph-bottom-border]') &&
  !p.querySelector(forbidden) && !unsupportedText.test(p.textContent ?? '') &&
  hwpWrapFingerprint(p.textContent ?? '') === fingerprint;

export function sourceColumnTail(p: HTMLElement, line: HwpBodyLine): HTMLElement | null {
  const fit = line.columnFit, previous = p.previousElementSibling;
  if (!fit?.hang || !(previous instanceof HTMLElement) || !provenText(p, fit.fingerprint) ||
      !provenText(previous, fit.preceding) || !/[^\s] $/u.test(previous.textContent ?? '')) return null;
  return previous;
}

export function layoutSourceColumnTail(p: HTMLElement, line: HwpBodyLine) {
  const previous = sourceColumnTail(p, line);
  if (!previous || previous.querySelector('[data-hwp-column-tail]')) return;
  const walker = document.createTreeWalker(previous, NodeFilter.SHOW_TEXT);
  let last: Text | undefined;
  while (walker.nextNode()) if (walker.currentNode.textContent?.length) last = walker.currentNode as Text;
  if (!last?.data.endsWith(' ')) return;
  const whitespace = last.splitText(last.length - 1), span = document.createElement('span');
  span.dataset.hwpColumnTail = 'true';
  whitespace.replaceWith(span);
  span.append(whitespace);
}

function makeMeasureBlock(p: HTMLElement, style: CSSStyleDeclaration) {
  const line = p.cloneNode(true) as HTMLElement;
  const marker = document.createElement('span');
  marker.style.cssText = 'display:inline-block;vertical-align:baseline;height:0;width:0;line-height:0;font-size:0;padding:0;border:0;margin:0';
  Object.assign(line.style, { margin: '0', whiteSpace: style.whiteSpace, fontFamily: style.fontFamily, fontSize: style.fontSize,
    fontWeight: style.fontWeight, fontStyle: style.fontStyle, letterSpacing: style.letterSpacing, wordSpacing: style.wordSpacing });
  delete line.dataset.hwpColumnEnd;
  restoreSourceColumnEnd(line);
  line.style.removeProperty('--hwp-column-em');
  line.style.removeProperty('--hwp-column-height');
  return { line, marker };
}

function projectedInk(line: HTMLElement, maxSize: number): Paint | null {
  canvas ??= document.createElement('canvas').getContext('2d');
  if (!canvas) return null;
  let ascent = -Infinity, descent = -Infinity;
  const walker = document.createTreeWalker(line, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const leaf = walker.currentNode as Text;
    if (!leaf.data.trim()) continue;
    const css = getComputedStyle(leaf.parentElement!);
    const size = parseFloat(css.fontSize);
    if (!(size > 0 && size <= maxSize + .1) || css.textDecorationLine !== 'none' ||
      css.textShadow !== 'none' || css.verticalAlign !== 'baseline' || css.textTransform !== 'none') return null;
    // Scale the canvas measurement only. Chromium's paint and box estimates
    // approach subpixel precision without changing the actual visible run.
    canvas.font = `${css.fontStyle} ${css.fontWeight} ${size * 16}px ${css.fontFamily}`;
    const ink = canvas.measureText(leaf.data);
    ascent = Math.max(ascent, ink.actualBoundingBoxAscent / 16);
    descent = Math.max(descent, ink.actualBoundingBoxDescent / 16);
  }
  return Number.isFinite(ascent) && Number.isFinite(descent) ? { ascent, descent } : null;
}

function horizontalScale(p: HTMLElement): number | null {
  const children = [...p.querySelectorAll<HTMLElement>(contentElements)];
  const values = children.map(run => parseCharacterScale(run.style.getPropertyValue(characterScaleProperty)));
  return values.length && values[0] !== null && values[0] >= 50 && values[0] < 100 &&
    values.every(value => value === values[0]) ? values[0] / 100 : null;
}

/** Pick the original baseline only after confirming the same text, measure,
 * column and predecessor. Ordinary reflow handles changed content or fonts. */
export function measureSourceColumnEnd(p: HTMLElement, source: HwpBodyLine, zoom: number, bottom: number, ordinary: number | null): SourceColumnPlacement | null {
  const fit = source.columnFit, previous = p.previousElementSibling;
  const trace = (stage: string, ...numbers: number[]) => {
    if (!fit || reported.size >= 100) return;
    const entry = JSON.stringify({ remainder:fit.remainder, baseline:fit.baseline, width:fit.width, stage, numbers:numbers.map(n=>Math.round(n*10000)/10000) });
    if (!reported.has(entry)) {
      reported.add(entry);
      console.info('WINDOWS_HWP_COLUMN_STAGES ' + entry);
    }
  };
  if (!fit || !provenText(p, fit.fingerprint) || !(previous instanceof HTMLElement) ||
      !provenText(previous, fit.preceding) || !p.parentElement?.matches('[data-column-count="2"]') ||
      [...p.childNodes].some(node => !(node instanceof HTMLSpanElement)) || !(zoom > 0)) return null;
  const css = getComputedStyle(p), before = getComputedStyle(previous);
  if (css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' || css.transform !== 'none' ||
      before.transform !== 'none' || (parseFloat(css.marginTop) || 0) || (parseFloat(before.marginBottom) || 0)) return null;
  const previousBox = previous.getBoundingClientRect();
  const room = (bottom - previousBox.bottom) / zoom;
  trace('room', room, ordinary ?? 0, source.base);
  if (!(room > 0 && room < source.base - .1) || ordinary !== null && ordinary <= room) return null;
  const width = p.getBoundingClientRect().width / zoom, indent = parseFloat(css.textIndent) || 0;
  if (Math.abs(width - Math.min(0, indent) - fit.width) > .24) return null;
  const host = document.createElement('div');
  host.className = 'tiptap'; host.setAttribute('aria-hidden', 'true');
  host.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;min-height:0;padding:0;width:${width}px`;
  const { line, marker } = makeMeasureBlock(p, css);
  const { line: preceding, marker: previousBaseline } = makeMeasureBlock(previous, before);
  // Keep the ordinary measured line when this engine already fits it.
  host.append(line); document.body.append(host);
  try {
    let scaled: SourceColumnPlacement | null = null;
    if (line.getBoundingClientRect().height > source.base + .08) {
      const factor = horizontalScale(p), left = parseFloat(css.marginLeft) || 0, right = parseFloat(css.marginRight) || 0;
      if (factor === null || left < 0 || right < 0 || indent + left < -.05) return null;
      line.style.width = `${width / factor}px`;
      line.style.textIndent = `${indent / factor}px`;
      scaled = { height: 0, offset: 0, scale: factor, width: `calc(${100 / factor}% - ${(left + right) / factor}px)`,
        indent: `${indent / factor}px` };
    }
    if (Math.abs(line.getBoundingClientRect().height - source.base) > .08) return null;
    line.dataset.hwpColumnEnd = 'true';
    line.style.setProperty('--hwp-column-em', `${source.font}px`);
    line.append(marker);
    const maximum = source.maxFont ?? source.font;
    const ink = projectedInk(line, maximum);
    if (!ink) return null;
    const baseline = marker.getBoundingClientRect().bottom - line.getBoundingClientRect().top;
    if (fit.baseline + ink.descent > fit.remainder + lineTolerance) return null;
    const heightAtBaseline = (paintBaseline: number) => Math.ceil((paintBaseline + ink.descent + pixelGuard) * 64) / 64;
    const ordinaryHeight = heightAtBaseline(fit.baseline);
    const missingRoom = Math.max(0, ordinaryHeight - Math.floor(room * 64) / 64);
    // The authored baseline is also a measured distance from the footer.
    // If the previous box consumes slightly more room, keep that distance
    // within the source measure's precision and recheck the previous ink.
    trace('fit-room', room, ordinaryHeight, missingRoom, fit.remainder - room, Number(scaled !== null), baseline, ink.ascent, ink.descent);
    if (missingRoom && (scaled || fit.remainder <= room ||
        Math.abs(missingRoom - (fit.remainder - room)) > lineTolerance)) return null;
    const paintBaseline = fit.baseline - missingRoom;
    const offset = paintBaseline - baseline;
    if (!(offset < -.1 && offset >= -4)) return null;

    preceding.style.width = `${previousBox.width / zoom}px`;
    host.append(preceding);
    const previousRect = preceding.getBoundingClientRect();
    const hanging = preceding.querySelector('[data-hwp-column-tail]');
    if (hanging) hanging.before(previousBaseline);
    else preceding.append(previousBaseline);
    const priorInk = projectedInk(preceding, maximum);
    if (!priorInk || preceding.getBoundingClientRect().height !== previousRect.height ||
        Math.abs(previousRect.height - previousBox.height / zoom) > .12) return null;
    const upperInk = paintBaseline - ink.ascent, previousEnd = previousBaseline.getBoundingClientRect().bottom - previousRect.top + priorInk.descent;
    trace('relative-ink', room, previousRect.height + upperInk - previousEnd, fit.baseline - paintBaseline);
    if (previousRect.height + upperInk - previousEnd < pixelGuard) return null;
    const height = heightAtBaseline(paintBaseline);
    return height > 0 && height <= room ? { ...scaled, height, offset } : null;
  } finally { host.remove(); }
}

function sourceColumnPaint(placement: SourceColumnPlacement) {
  const inline: Record<string, string> = { '--hwp-column-paint-offset': `${placement.offset}px` };
  if (placement.width && placement.indent && placement.scale) {
    inline['--hwp-column-scaled-width'] = placement.width;
    inline['--hwp-column-scaled-indent'] = placement.indent;
    inline['--hwp-column-scale'] = String(placement.scale);
  }
  return inline;
}

export function sourceColumnPaintStyle(placement: SourceColumnPlacement) {
  return Object.entries(sourceColumnPaint(placement)).map(([name, value]) => `${name}:${value}`).join(';');
}

export function restoreSourceColumnEnd(paragraph: HTMLElement) {
  if (paragraph.dataset.hwpColumnPaint !== 'true') return;
  delete paragraph.dataset.hwpColumnPaint;
  delete paragraph.dataset.hwpColumnScaled;
  for (const property of ['--hwp-column-paint-offset', '--hwp-column-scaled-width', '--hwp-column-scaled-indent', '--hwp-column-scale'])
    paragraph.style.removeProperty(property);
}

export function applySourceColumnEnd(paragraph: HTMLElement, placement: SourceColumnPlacement) {
  paragraph.dataset.hwpColumnPaint = 'true';
  if (placement.scale) paragraph.dataset.hwpColumnScaled = 'true';
  for (const [name, value] of Object.entries(sourceColumnPaint(placement))) paragraph.style.setProperty(name, value);
}

export function markSourceColumnTails(root: HTMLElement) {
  for (const p of root.querySelectorAll<HTMLElement>('p[data-hwp-column-end-source]')) {
    const source = parseHwpBodyLine(p.dataset.hwpBodyLine);
    if (source?.columnFit) layoutSourceColumnTail(p, source);
  }
}
