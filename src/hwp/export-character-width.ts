import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { characterScaleProperty, parseCharacterScale } from '../character-scale';
import { entities } from './cell-content';
import { elementRanges } from './markup-ranges';

const colorDeclaration = /(?:^|;)\s*color\s*:\s*([^;]+)(?=;|$)/gi;
const widthDeclaration = new RegExp(`(?:^|;)\\s*${characterScaleProperty}\\s*:\\s*([^;]+)(?=;|$)`, 'gi');
const languageWidths = ['hangul', 'latin', 'hanja', 'japanese', 'other', 'symbol', 'user'];

function opaqueColor(value: string): string | undefined {
  const text = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(text)) return text;
  if (/^#[0-9a-f]{3}$/.test(text)) return '#' + [...text.slice(1)].map(digit => digit + digit).join('');
  const channels = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/.exec(text)?.slice(1).map(Number);
  if (channels && channels.every(value => value <= 255))
    return '#' + channels.map(value => value.toString(16).padStart(2, '0')).join('');
}

type WidthStyle = { color: string; ratio: number; text: string[] };
const literal = (markup: string) => entities(markup.replace(/&nbsp;/gi, ' ')).replace(/\s/gu, '');

function pasteTexts(files: Record<string, Uint8Array>, refs: ReadonlyMap<string, string>) {
  const text = new Map<string, string[]>();
  for (const [name, file] of Object.entries(files)) {
    if (!/^Contents\/section\d+\.xml$/.test(name)) continue;
    const runs: (string | undefined)[] = [];
    let inText = false;
    for (const [tag] of strFromU8(file).matchAll(/<[^>]+>|[^<]+/g)) {
      if (tag.startsWith('<hp:run ')) {
        if (!tag.endsWith('/>')) runs.push(refs.get(tag.match(/\bcharPrIDRef="(\d+)"/)?.[1] ?? ''));
      } else if (tag === '</hp:run>') runs.pop();
      else if (/^<hp:t(?:\s|>)/.test(tag)) inText = !tag.endsWith('/>');
      else if (tag === '</hp:t>') inText = false;
      else if (inText && !tag.startsWith('<')) {
        const marker = runs.at(-1);
        if (!marker) continue;
        let parts = text.get(marker);
        if (!parts) text.set(marker, parts = []);
        parts.push(literal(tag));
      }
    }
  }
  return text;
}

/** The native HTML paste carries opaque text colors but ignores glyph-width
 * ratios. Carry a temporary, unoccupied color through its run formatting and
 * restore the original color and authored ratio in the native char shape. */
export function stageHwpCharacterWidths(html: string) {
  const styles = new Map<string, WidthStyle>(), palettes = new Map<string, string>();
  const taken = new Set(['#000000', '#ffffff']);
  let staged = html;
  if (html.includes(characterScaleProperty)) {
    for (const match of html.matchAll(/#[0-9a-f]{3,8}\b|rgb\([^)]*\)/gi)) {
      const color = opaqueColor(match[0]);
      if (color) taken.add(color);
    }
    let nextColor = crypto.getRandomValues(new Uint32Array(1))[0] % 0x1000000;
    // The print tree can wrap the terminal spaces of an otherwise uniform
    // run in a plain span. It inherits the same color and glyph width.
    staged = html.replace(/<span\b([^>]*)>([^<>]+)(?:<span>(?:\s|&nbsp;)+<\/span>)?<\/span>/gi, (whole, attributes: string, content: string) => {
      const css = /(?:^|\s)style="([^"]*)"/i.exec(attributes);
      if (!css) return whole;
      const text = literal(content);
      if (!text) return whole;
      const width = [...css[1].matchAll(widthDeclaration)].at(-1);
      const ratio = parseCharacterScale(width?.[1]);
      if (ratio === null || ratio === 100 || ratio < 50 || ratio > 200) return whole;
      const last = [...css[1].matchAll(colorDeclaration)].at(-1);
      const color = last ? opaqueColor(last[1]) : undefined;
      if (!last || !color) return whole;
      const key = color + ':' + ratio;
      let marker = palettes.get(key);
      if (!marker) {
        do {
          nextColor = (nextColor + 1) % 0x1000000;
          marker = '#' + nextColor.toString(16).padStart(6, '0');
        } while (taken.has(marker));
        taken.add(marker);
        palettes.set(key, marker);
        styles.set(marker, { color, ratio, text: [] });
      }
      styles.get(marker)!.text.push(text);
      const at = last.index! + last[0].lastIndexOf(last[1]);
      const replacement = css[1].slice(0, at) + marker + css[1].slice(at + last[1].length);
      return whole.replace(css[0], () => css[0].replace(css[1], () => replacement));
    });
  }
  return { html: staged, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!styles.size) return doc;
    let size = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 장평 변환 데이터가 너무 큽니다.');
      return true;
    } });
    const name = 'Contents/header.xml', original = strFromU8(files[name]);
    let xml = original;
    const refs = new Map<string, string>();
    for (const shape of elementRanges(original, 'hh:charPr').reverse()) {
      const marker = shape.open.match(/\btextColor="(#[0-9a-f]{6})"/i)?.[1].toLowerCase();
      const source = marker && styles.get(marker);
      if (!marker || !source) continue;
      const id = shape.open.match(/\bid="(\d+)"/)?.[1];
      if (!id || !/<hh:ratio\b[^>]*\/>/.test(shape.inner)) return null;
      refs.set(id, marker);
      const ratio = '<hh:ratio ' + languageWidths.map(name => `${name}="${source.ratio}"`).join(' ') + '/>';
      const originalShape = original.slice(shape.start, shape.end);
      const replacement = originalShape.replace(/"(#[0-9a-f]{6})"/gi,
        (value, color) => color.toLowerCase() === marker ? `"${source.color}"` : value)
        .replace(/<hh:ratio\b[^>]*\/>/, ratio);
      xml = xml.slice(0, shape.start) + replacement + xml.slice(shape.end);
    }
    const text = pasteTexts(files, refs);
    for (const [marker, style] of styles) {
      if (text.get(marker)?.join('') !== style.text.join('')) return null;
    }
    files[name] = strToU8(xml);
    return new Core(zipSync(files, { level: 1 }));
  } };
}
