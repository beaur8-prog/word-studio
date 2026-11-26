import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { hwpInlineTableTop, hwpInlineTableTopAttribute, tableAnchor, type TableAnchor } from '../table-anchor-data';
import { elementRanges, topLevelHwpPasteBlocks } from './markup-ranges';
import { entities } from './cell-content';
import documentLimits from '../../electron/document-limits.json';

const set = (tag: string, values: Record<string, string | number>) => {
  for (const [key, value] of Object.entries(values)) {
    const pattern = new RegExp(`\\b${key}="[^"]*"`), text = `${key}="${value}"`;
    tag = pattern.test(tag) ? tag.replace(pattern, text) : tag.replace(/\s*\/?>$/, end => ' ' + text + end);
  }
  return tag;
};
const pasteOuterMargin = 283;
const pasteTableIndent = 20;
/** The source table follows a one-row label and an intact blank on the
 * same printed sheet. An isolated continuation needs the ordinary paste
 * placement to avoid creating another sheet for its empty native carrier. */
function markupAnchor(table: string) {
  const value = table.match(/\bdata-table-anchor=(["'])([\s\S]*?)\1/)?.[2];
  return value ? tableAnchor(entities(value)) : undefined;
}

function nativeParagraphFlowTables(html: string): Map<number, 'flow' | 'label'> {
  const result = new Map<number, 'flow' | 'label'>();
  if (!html.includes('data-hwp-centered-borders=')) return result;
  let start = 0;
  const scan = (end: number) => {
    const page = html.slice(start, end);
    if (!page.includes('data-hwp-centered-borders=')) return;
    const blocks = topLevelHwpPasteBlocks(page);
    if (blocks?.length !== 3 || !blocks[0].trimStart().startsWith('<table') || !blocks[2].trimStart().startsWith('<table') ||
        elementRanges(blocks[0], 'tr').length !== 1) return;
    const [space] = elementRanges(blocks[1], 'p'), [table] = elementRanges(blocks[2], 'table');
    const [label] = elementRanges(blocks[0], 'table');
    if (!table || !space || !label || entities(space.inner.replace(/<\/?span\b[^>]*>/g, '')).trim() ||
        /<table\b|<(?:img|svg)\b|\bdata-table-anchor=/.test(label.inner)) return;
    const anchor = markupAnchor(table.open);
    if (!anchor || !nativeParagraphFlow(table.open, anchor)) return;
    result.set(end - blocks[2].length + table.start, 'flow');
    if (!markupAnchor(label.open)) result.set(start + label.start, 'label');
  };
  for (const marker of html.matchAll(/<div\b[^>]*\bdata-(?:page-break|section-page)=["'][^>]*>\s*<\/div>/gi)) {
    scan(marker.index!);
    start = marker.index! + marker[0].length;
  }
  scan(html.length);
  return result;
}

function nativeParagraphFlow(table: string, anchor: TableAnchor): 'CELL' | 'TABLE' | undefined {
  if (anchor.reference !== 'paragraph' || anchor.horizontal !== 'margin' || anchor.wrap ||
      anchor.x >= .05 || anchor.y !== 0 || anchor.bottom !== 0) return;
  const border = table.match(/\bdata-hwp-centered-borders="(true|continuation)"/)?.[1];
  return border === 'continuation' ? 'CELL' : border === 'true' ? 'TABLE' : undefined;
}

function projectedInlineInsets(table: string, anchor: TableAnchor) {
  if (anchor.reference !== 'paragraph' || anchor.horizontal !== 'margin' || anchor.wrap || anchor.overlay ||
      /\bdata-hwp-(?:centered-borders|table-continuation)=/.test(table)) return;
  const top = Math.round(anchor.y * 75), bottom = Math.round(anchor.bottom * 75);
  const left = Math.round(anchor.x * 75);
  // A table already converted to ordinary flow can retain its authored
  // before/after reserve without acquiring a larger pasted footprint.
  if (top <= pasteOuterMargin && bottom <= pasteOuterMargin && (top !== pasteOuterMargin || bottom !== pasteOuterMargin))
    return { mode: 'inline-insets', top, bottom, left: left <= pasteOuterMargin ? left : undefined } as const;
}

/** Sentinels in an existing cell identify its real control after HTML import;
 * no marker paragraph or original text is removed. */
export function stageHwpTables(html: string) {
  type Placement = { mode: 'label' } | { mode: 'overlay'; anchor: TableAnchor } |
    { mode: 'flow'; anchor: TableAnchor; breakMode: 'CELL' | 'TABLE' } | { mode: 'inline-insets'; top: number; bottom?: number; left?: number };
  const prefix = `WORDSTUDIO-TABLE-${crypto.randomUUID()}-`, records = new Map<string, Placement>();
  const flowTables = nativeParagraphFlowTables(html);
  const topAttribute = new RegExp(`\\b${hwpInlineTableTopAttribute}="([^"]*)"`);
  let staged = html;
  for (const table of elementRanges(html, 'table').reverse()) {
    const anchor = markupAnchor(table.open), mode = flowTables.get(table.start);
    const breakMode = anchor && mode === 'flow' ? nativeParagraphFlow(table.open, anchor) : undefined;
    const top = anchor ? undefined : hwpInlineTableTop(topAttribute.exec(table.open)?.[1]);
    const placement = mode === 'label' ? { mode } as const
      : anchor && breakMode ? { mode: 'flow', anchor, breakMode } as const
      : anchor?.overlay ? { mode: 'overlay', anchor } as const
      : anchor ? projectedInlineInsets(table.open, anchor)
      : top !== undefined && top !== pasteOuterMargin ? { mode: 'inline-insets', top } as const : undefined;
    if (!placement) continue;
    let depth = 0, at = -1;
    for (const m of table.inner.matchAll(/<\/?(?:table|p)\b[^>]*>/g)) {
      if (/^<table\b/.test(m[0])) depth++;
      else if (/^<\/table\b/.test(m[0])) depth--;
      else if (!depth && /^<p\b/.test(m[0])) { at = m.index! + m[0].length; break; }
    }
    if (at < 0) continue;
    const token = `${prefix}${records.size}-END`; records.set(token, placement);
    const inner = table.inner.slice(0, at) + `<span>${token}</span>` + table.inner.slice(at);
    staged = staged.slice(0, table.start) + table.open + inner + '</table>' + staged.slice(table.end);
  }
  return { html: staged, restore(Core: typeof HwpDocument, doc: HwpDocument) {
    if (!records.size) return doc;
    let size = 0, restored = 0;
    const files = unzipSync(doc.exportHwpx(), { filter: file => {
      if ((size += file.originalSize) > documentLimits.documentInflatedBytes) throw Error('한글 표 배치 데이터가 너무 큽니다.');
      return true;
    } });
    const pasteIndent: { section: number; paragraph: number }[] = [];
    const hasInlineLeft = [...records.values()].some(record => record.mode === 'inline-insets' && record.left !== undefined);
    for (const [name, bytes] of Object.entries(files)) {
      const sectionName = /^Contents\/section(\d+)\.xml$/.exec(name);
      if (!sectionName) continue;
      const section = Number(sectionName[1]);
      let xml = strFromU8(bytes), anchorsChanged = false;
      const hosts = hasInlineLeft ? elementRanges(xml, 'hp:p') : [];
      let owner = hosts.length - 1;
      for (const table of elementRanges(xml, 'hp:tbl').reverse()) {
        const tokens = table.inner.match(new RegExp(`${prefix}\\d+-END`, 'g')) || [];
        if (!tokens.length) continue;
        if (tokens.length !== 1) throw Error('한글 표의 배치 위치가 모호합니다.');
        const record = records.get(tokens[0]); if (!record) throw Error('한글 표의 배치 정보를 찾지 못했습니다.');
        if (record.mode === 'inline-insets') {
          // Restore only the newly pasted default on this exact inline object.
          const inner = table.inner.replace(/^<hp:sz\b[^>]*\/><hp:pos\b[^>]*\/><hp:outMargin\b[^>]*\/>/, prefix =>
            /<hp:pos\b[^>]*\btreatAsChar="1"/.test(prefix) ? prefix.replace(/<hp:outMargin\b[^>]*\/>/, tag => {
              const values: Record<string, number> = {};
              if (/\btop="283"/.test(tag)) values.top = record.top;
              if (record.bottom !== undefined && /\bbottom="283"/.test(tag)) values.bottom = record.bottom;
              if (record.left !== undefined && /\bleft="283"/.test(tag)) {
                while (owner >= 0 && hosts[owner].start > table.start) owner--;
                const host = hosts[owner];
                // The paste engine indents a generated table carrier by 20px.
                // Only move an otherwise empty, single-control carrier whose
                // default inset would shift the measured anchor to the right.
                if (host && host.end > table.end && doc.getTextRange(section, owner, 0, 0xffffffff) === '') {
                  const paragraph = JSON.parse(doc.getParaPropertiesAt(section, owner));
                  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, owner, '[]'));
                  if (paragraph.marginLeft === pasteTableIndent && paragraph.marginRight === 0 && paragraph.indent === 0 &&
                      metrics.controlCount === 1 && /<hp:pos\b[^>]*\bhorzRelTo="PARA"/.test(prefix) &&
                      /<hp:pos\b[^>]*\bhorzAlign="LEFT"/.test(prefix) &&
                      /<hp:pos\b[^>]*\bhorzOffset="0"/.test(prefix)) {
                    values.left = record.left;
                    pasteIndent.push({ section, paragraph: owner });
                  }
                }
              }
              return set(tag, values);
            }) : prefix).replace(tokens[0], '');
          xml = xml.slice(0, table.start) + table.open + inner + '</hp:tbl>' + xml.slice(table.end);
          restored++;
          continue;
        }
        anchorsChanged = true;
        const flow = record.mode === 'flow';
        const open = record.mode === 'label' ? table.open : set(table.open, flow
          ? { textWrap: 'TOP_AND_BOTTOM', pageBreak: record.breakMode }
          : { textWrap: 'IN_FRONT_OF_TEXT' });
        let found = false;
        const inner = table.inner.replace(/<hp:pos\b[^>]*\/>/, tag => {
          found = true;
          if (record.mode === 'label') return tag;
          const { anchor } = record;
          return set(tag, { treatAsChar: 0, affectLSpacing: 0, flowWithText: flow ? 1 : 0, allowOverlap: flow ? 0 : 1, holdAnchorAndSO: 0,
            vertRelTo: flow ? 'PARA' : 'PAPER', horzRelTo: flow ? 'COLUMN' : 'PAPER', vertAlign: flow ? 'TOP' : 'BOTTOM', horzAlign: 'LEFT',
            vertOffset: Math.round(anchor.y * 75), horzOffset: Math.round(anchor.x * 75) });
        // The print group already accounted for its margins. The paste API
        // adds another 283 HU on every edge, even around the inline label.
        }).replace(/<hp:outMargin\b[^>]*\/>/, tag => set(tag, { left: 0, right: 0, top: 0, bottom: 0 })).replace(tokens[0], '');
        if (!found) throw Error('한글 표의 위치 속성을 찾지 못했습니다.');
        xml = xml.slice(0, table.start) + open + inner + '</hp:tbl>' + xml.slice(table.end); restored++;
      }
      files[name] = strToU8(anchorsChanged ? xml.replace(/<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/g, '') : xml);
    }
    if (restored !== records.size) throw Error('일부 떠 있는 표를 한글로 저장하지 못했습니다.');
    const result = new Core(zipSync(files, { level: 1 }));
    try {
      if (pasteIndent.length) {
        result.beginBatch();
        try {
          for (const { section, paragraph } of pasteIndent) {
            const adjusted = JSON.parse(result.applyParaFormat(section, paragraph, JSON.stringify({ marginLeft: 0 })));
            if (!adjusted.ok) throw Error('한글 표의 본문 들여쓰기를 복원하지 못했습니다.');
          }
        } finally { result.endBatch(); }
      }
      return result;
    } catch (error) { result.free(); throw error; }
  } };
}
