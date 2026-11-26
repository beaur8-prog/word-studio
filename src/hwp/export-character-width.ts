import { activeHwpLatinFont, hwpLatinFontAttribute } from '../hwp-latin-font';
import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { characterScaleProperty, parseCharacterScale } from '../character-scale';
import { entities } from './cell-content';
import { elementRanges } from './markup-ranges';

const colorDeclaration = /(?:^|;)\s*color\s*:\s*([^;]+)(?=;|$)/gi;
const widthDeclaration = new RegExp(`(?:^|;)\\s*${characterScaleProperty}\\s*:\\s*([^;]+)(?=;|$)`, 'gi');
const latinAttribute = new RegExp(`(?:^|\\s)${hwpLatinFontAttribute}="([^"]*)"`, 'i');
const fontFamilyDeclaration = /(?:^|;)\s*font-family\s*:\s*([^;]+)(?=;|$)/gi;
const languageWidths = ['hangul', 'latin', 'hanja', 'japanese', 'other', 'symbol', 'user'];

function opaqueColor(value: string): string | undefined {
  const text = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(text)) return text;
  if (/^#[0-9a-f]{3}$/.test(text)) return '#' + [...text.slice(1)].map(digit => digit + digit).join('');
  const channels = /^rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)$/.exec(text)?.slice(1).map(Number);
  if (channels && channels.every(value => value <= 255))
    return '#' + channels.map(value => value.toString(16).padStart(2, '0')).join('');
}

type CharacterStyle = { color: string; ratio?: number; latin?: string; text: string[] };
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

/** The native HTML paste carries opaque text colors but flattens the width
 * and language slots. Restore witnessed attributes on each exact native run
 * using a temporary, otherwise unused color. */
export function stageHwpCharacterStyles(html: string, includeLatin = true) {
  const styles = new Map<string, CharacterStyle>(), palettes = new Map<string, string>();
  const taken = new Set(['#000000', '#ffffff']);
  const canPreserveLatin = includeLatin && html.includes(hwpLatinFontAttribute);
  let hasSourceFonts = false;
  let staged = html;
  if (html.includes(characterScaleProperty) || canPreserveLatin) {
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
      const declaredRatio = parseCharacterScale(width?.[1]);
      const ratio = declaredRatio !== null && declaredRatio !== 100 && declaredRatio >= 50 && declaredRatio <= 200
        ? declaredRatio : undefined;
      const declaredLatin = canPreserveLatin ? latinAttribute.exec(attributes)?.[1] : undefined;
      const latin = declaredLatin && activeHwpLatinFont(entities(declaredLatin),
        [...css[1].matchAll(fontFamilyDeclaration)].at(-1)?.[1]);
      if (ratio === undefined && !latin) return whole;
      const last = [...css[1].matchAll(colorDeclaration)].at(-1);
      const color = last ? opaqueColor(last[1]) : undefined;
      if (!last || !color) return whole;
      hasSourceFonts ||= !!latin;
      const key = JSON.stringify([color, ratio, latin]);
      let marker = palettes.get(key);
      if (!marker) {
        do {
          nextColor = (nextColor + 1) % 0x1000000;
          marker = '#' + nextColor.toString(16).padStart(6, '0');
        } while (taken.has(marker));
        taken.add(marker);
        palettes.set(key, marker);
        styles.set(marker, { color, ratio, ...(latin ? { latin } : {}), text: [] });
      }
      styles.get(marker)!.text.push(text);
      const at = last.index! + last[0].lastIndexOf(last[1]);
      const replacement = css[1].slice(0, at) + marker + css[1].slice(at + last[1].length);
      return whole.replace(css[0], () => css[0].replace(css[1], () => replacement));
    });
  }
  return { html: staged, hasSourceFonts, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!styles.size) return doc;
    const latinIds = new Map<string, number>();
    for (const { latin } of styles.values()) {
      if (!latin || latinIds.has(latin)) continue;
      // Register a face name only. No font file is opened or embedded.
      const id = doc.findOrCreateFontIdForLang(1, latin);
      if (!Number.isInteger(id) || id < 0 || id > 65535) return null;
      latinIds.set(latin, id);
    }
    let size = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 장평 변환 데이터가 너무 큽니다.');
      return true;
    } });
    const name = 'Contents/header.xml', original = strFromU8(files[name]);
    const shapes = elementRanges(original, 'hh:charPr').flatMap(shape => {
      const marker = shape.open.match(/\btextColor="(#[0-9a-f]{6})"/i)?.[1].toLowerCase();
      const source = marker && styles.get(marker);
      return marker && source ? [{ shape, marker, source }] : [];
    });
    const refs = new Map<string, string>();
    for (const { shape, marker } of shapes) {
      const id = shape.open.match(/\bid="(\d+)"/)?.[1];
      if (!id) return null;
      refs.set(id, marker);
    }
    const text = pasteTexts(files, refs);
    for (const [marker, style] of styles) {
      if (text.get(marker)?.join('') !== style.text.join('')) return null;
    }
    let xml = original;
    for (const { shape, marker, source } of shapes.reverse()) {
      let replacement = original.slice(shape.start, shape.end).replace(/"(#[0-9a-f]{6})"/gi,
        (value, color) => color.toLowerCase() === marker ? `"${source.color}"` : value);
      if (source.ratio !== undefined) {
        if (!/<hh:ratio\b[^>]*\/>/.test(shape.inner)) return null;
        const ratio = '<hh:ratio ' + languageWidths.map(name => `${name}="${source.ratio}"`).join(' ') + '/>';
        replacement = replacement.replace(/<hh:ratio\b[^>]*\/>/, ratio);
      }
      if (source.latin) {
        const font = latinIds.get(source.latin);
        if (font === undefined || !/<hh:fontRef\b[^>]*\blatin="\d+"[^>]*\/>/.test(shape.inner)) return null;
        replacement = replacement.replace(/(<hh:fontRef\b[^>]*\blatin=")\d+"/,
          (_, prefix: string) => prefix + font + '"');
      }
      xml = xml.slice(0, shape.start) + replacement + xml.slice(shape.end);
    }
    files[name] = strToU8(xml);
    return new Core(zipSync(files, { level: 1 }));
  } };
}
