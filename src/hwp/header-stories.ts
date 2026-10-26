import type { HwpDocument } from '@rhwp/core';
import { strFromU8 } from 'fflate';

export type HeaderDefinition = { section: number; para: number; kind: 'header' | 'footer'; apply: number; xml: string };
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
