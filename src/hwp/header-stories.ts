import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, zipSync, unzipSync } from 'fflate';
import { elementRanges } from './markup-ranges';
import { entities } from './cell-content';
import documentLimits from '../../electron/document-limits.json';

export type HeaderDefinition = { section: number; para: number; kind: 'header' | 'footer'; apply: number; xml: string };

/** Text, tables, inline pictures and fields can stay editable. Arbitrary drawing
 * controls retain the existing isolated drawing-layer rendering path. */
export function editableHeaderStories(doc: HwpDocument, definitions: HeaderDefinition[], render: (doc: HwpDocument) => string) {
  const plain = definitions.filter(d => !/<hp:(?:rect|line|ellipse|polygon|container|equation|ole|footNote|endNote)\b/.test(d.xml));
  const result = new Map<HeaderDefinition, string>();
  if (!plain.length) return result;
  const prefix = `WORDSTUDIO-HEADER-${crypto.randomUUID()}-`;
  const marker = (text: string) => `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run></hp:p>`;
  let size = 0;
  const copy = unzipSync(doc.exportHwpx(), { filter: file => {
    if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 머리말 변환 데이터가 너무 큽니다.');
    return true;
  } });
  for (const name of Object.keys(copy).filter(n => /^Contents\/section\d+\.xml$/.test(n))) {
    const xml = strFromU8(copy[name]);
    const body = name === 'Contents/section0.xml' ? plain.map((definition, i) => {
      const story = definition.xml.match(/<hp:subList\b[^>]*>([\s\S]*)<\/hp:subList>/)?.[1];
      if (!story) throw Error('머리말·꼬리말의 본문을 찾지 못했습니다.');
      return marker(`${prefix}${i}-START`) + story + marker(`${prefix}${i}-END`);
    }).join('') : marker('');
    copy[name] = strToU8(xml.replace(/(<hs:sec\b[^>]*>)[\s\S]*?(<\/hs:sec>)/, '$1' + body + '$2'));
  }
  const Core = doc.constructor as typeof HwpDocument, isolated = new Core(zipSync(copy, { level: 1 }));
  try {
    const html = render(isolated); let active = -1, from = 0;
    for (const p of elementRanges(html, 'p')) {
      const text = entities(p.inner.replace(/<[^>]*>/g, '')).trim();
      const start = plain.findIndex((_, i) => text === `${prefix}${i}-START`);
      if (start >= 0) { active = start; from = p.end; }
      else if (active >= 0 && text === `${prefix}${active}-END`) {
        const definition = plain[active];
        let content = html.slice(from, p.start).trim();
        const tables = elementRanges(content, 'table');
        // Native tables used as inline header/footer content inherit their
        // owning paragraph's alignment. A block TABLE needs explicit margins.
        if (tables.length === 1 && !entities((content.slice(0, tables[0].start) + content.slice(tables[0].end)).replace(/<[^>]*>/g, '')).trim()) {
          const id = definition.xml.match(/<hp:p\b[^>]*paraPrIDRef="(\d+)"/)?.[1];
          const info = strFromU8(copy['Contents/header.xml']);
          const props = info.match(new RegExp(`<hh:paraPr\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/hh:paraPr>`))?.[0];
          const align = props?.match(/<hh:align\b[^>]*horizontal="([^"]+)"/)?.[1];
          if (align === 'CENTER' || align === 'RIGHT') content = content.replace(/<table\b([^>]*?)style="/, `<table$1style="margin-left:auto;margin-right:${align === 'CENTER' ? 'auto' : '0'};`);
        }
        result.set(definition, content); active = -1;
      }
    }
    if (result.size !== plain.length) throw Error('일부 머리말·꼬리말 본문을 변환하지 못했습니다.');
    return result;
  } finally { isolated.free(); }
}

export function headerDefinitions(files: Record<string, Uint8Array>): HeaderDefinition[] {
  const result: HeaderDefinition[] = [];
  for (const [name, bytes] of Object.entries(files)) {
    const match = name.match(/^Contents\/section(\d+)\.xml$/); if (!match) continue;
    const section = Number(match[1]), xml = strFromU8(bytes);
    let depth = 0, para = -1, start = -1, kind: 'header' | 'footer' = 'header', apply = 0;
    for (const token of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\/?(?:[\w.-]+:)?(?:p|header|footer)\b[^>]*>/g)) {
      const tag = token[0]; if (tag.startsWith('<!')) continue;
      if (/^<\/(?:[\w.-]+:)?p\b/.test(tag)) depth--;
      else if (/^<(?:[\w.-]+:)?p\b/.test(tag)) { if (depth++ === 0) para++; if (tag.endsWith('/>')) depth--; }
      else if (/^<\//.test(tag)) {
        if (start >= 0) result.push({ section, para, kind, apply, xml: xml.slice(start, token.index! + tag.length) });
        start = -1;
      } else {
        start = token.index!; kind = /:footer\b/.test(tag) ? 'footer' : 'header';
        const type = tag.match(/applyPageType="([^"]+)"/)?.[1]; apply = type === 'EVEN' ? 1 : type === 'ODD' ? 2 : 0;
      }
    }
  }
  return result;
}

/** Match a story by its owning source paragraph and parity. Header control
 * indices may collide with body tables or controls nested inside cells. */
export function headerPreviewPage(doc: HwpDocument, definition: HeaderDefinition): number | undefined {
  for (let page = 0; page < doc.pageCount(); page++) {
    if (definition.apply && (page + 1) % 2 !== (definition.apply === 1 ? 0 : 1)) continue;
    const controls = JSON.parse(doc.getPageControlLayout(page)).controls ?? [];
    if (controls.some((c: any) => c.secIdx === definition.section && c.headerFooter?.kind === definition.kind && c.headerFooter.outerParaIdx === definition.para)) return page;
  }
}

export function headerLayer(root: any, kind: string): any | undefined {
  if (root?.groupKind?.kind === kind) return root;
  for (const child of root?.children ?? []) { const match = headerLayer(child, kind); if (match) return match; }
  return root?.child ? headerLayer(root.child, kind) : undefined;
}

export function hiddenStories(files: Record<string, Uint8Array>) {
  const result: { section: number; para: number; header: boolean; footer: boolean }[] = [];
  for (const [name, bytes] of Object.entries(files)) {
    const match = name.match(/^Contents\/section(\d+)\.xml$/); if (!match) continue;
    let depth = 0, para = -1;
    for (const token of strFromU8(bytes).matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\/?(?:[\w.-]+:)?(?:p|pageHiding)\b[^>]*>/g)) {
      const tag = token[0]; if (tag.startsWith('<!')) continue;
      if (/^<\/(?:[\w.-]+:)?p\b/.test(tag)) depth--;
      else if (/^<(?:[\w.-]+:)?p\b/.test(tag)) { if (depth++ === 0) para++; if (tag.endsWith('/>')) depth--; }
      else if (!tag.startsWith('</')) {
        const header = /hideHeader="1"/.test(tag), footer = /hideFooter="1"/.test(tag);
        if (header || footer) result.push({ section: Number(match[1]), para, header, footer });
      }
    }
  }
  return result;
}
