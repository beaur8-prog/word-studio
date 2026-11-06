import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { elementRanges } from './markup-ranges';
import { entities } from './cell-content';
import documentLimits from '../../electron/document-limits.json';

type Paragraph = { key: string; tables: number };
type Table = { key: string; cells: number };
type Cell = { key: string; paragraphs: number };
type Caption = { key: string; side: string; gap: number; xml: string };
type Frame = { name: string; paragraph?: Paragraph; table?: Table; cell?: Cell; caption?: { start: number; key: string; side: string; gap: number } };

/** Caption paragraphs are a separate story, not another paragraph in the owner
 * cell. Match ordinary table paths without advancing that cell's paragraph or
 * table ordinal while traversing its caption. Nested caption stories are
 * recovered when their containing story is rendered in isolation. */
export function tableCaptionDefinitions(files: Record<string, Uint8Array>): Caption[] {
  const result: Caption[] = [];
  for (const [file, bytes] of Object.entries(files)) {
    const section = file.match(/^Contents\/section(\d+)\.xml$/); if (!section) continue;
    const xml = strFromU8(bytes), stack: Frame[] = []; let paragraph = -1;
    for (const m of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<[^>]+>/g)) {
      const tag = m[0]; if (tag.startsWith('<!') || tag.startsWith('<?')) continue;
      const name = tag.match(/^<\/?(?:[\w.-]+:)?([\w.-]+)/)?.[1]; if (!name) continue;
      if (tag.startsWith('</')) {
        const frame = stack.pop();
        if (name === 'caption' && frame?.caption) result.push({ ...frame.caption, xml: xml.slice(frame.caption.start, m.index! + tag.length) });
        continue;
      }
      const parent = stack.at(-1), frame: Frame = { ...parent, name, caption: undefined };
      const attr = (key: string) => tag.match(new RegExp(`\\b${key}=["']([^"']*)["']`))?.[1] ?? '';
      if (name === 'p') {
        if (parent?.name === 'sec') frame.paragraph = { key: `${section[1]}:${++paragraph}`, tables: 0 };
        else if (parent?.name === 'subList' && stack.at(-2)?.name === 'tc' && parent.cell)
          frame.paragraph = { key: `${parent.cell.key}/p${parent.cell.paragraphs++}`, tables: 0 };
        else { frame.paragraph = undefined; frame.table = undefined; frame.cell = undefined; }
      } else if (name === 'tbl') {
        frame.table = frame.paragraph ? { key: `${frame.paragraph.key}/t${frame.paragraph.tables++}`, cells: 0 } : undefined;
      } else if (name === 'tc' && frame.table) frame.cell = { key: `${frame.table.key}/c${frame.table.cells++}`, paragraphs: 0 };
      else if (name === 'caption' && parent?.name === 'tbl' && parent.table)
        frame.caption = { start: m.index!, key: parent.table.key, side: attr('side'), gap: Math.max(0, Number(attr('gap')) || 0) / 75 };
      if (!tag.endsWith('/>')) stack.push(frame);
    }
  }
  return result;
}

/** Render captions with the same font, cell and picture conversion as body
 * stories. Keep their original package resources, and never mutate the input. */
export function renderTableCaptions(doc: HwpDocument, definitions: Caption[], render: (doc: HwpDocument) => { html: string; warnings: string[] }) {
  const result = new Map<string, { html: string; side: string; gap: number }>();
  const warnings = new Set<string>();
  if (!definitions.length) return { captions: result, warnings: [] };
  if (definitions.length > 4096) throw Error('표 캡션이 너무 많아 안전하게 변환하지 못했습니다.');
  const prefix = `WORDSTUDIO-CAPTION-${crypto.randomUUID()}-`;
  const marker = (text: string) => `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run></hp:p>`;
  let size = 0;
  const copy = unzipSync(doc.exportHwpx(), { filter: f => {
    if ((size += f.originalSize) > documentLimits.documentInflatedBytes) throw Error('표 캡션 변환 데이터가 너무 큽니다.');
    return true;
  } });
  for (const file of Object.keys(copy).filter(n => /^Contents\/section\d+\.xml$/.test(n))) {
    const xml = strFromU8(copy[file]);
    const body = file === 'Contents/section0.xml' ? definitions.map((definition, i) => {
      const story = elementRanges(definition.xml, 'hp:subList')[0];
      if (!story) throw Error('표 캡션의 본문을 찾지 못했습니다.');
      return marker(`${prefix}${i}-START`) + story.inner + marker(`${prefix}${i}-END`);
    }).join('') : marker('');
    copy[file] = strToU8(xml.replace(/(<hs:sec\b[^>]*>)[\s\S]*?(<\/hs:sec>)/, '$1' + body + '$2'));
  }
  const Core = doc.constructor as typeof HwpDocument, isolated = new Core(zipSync(copy, { level: 1 }));
  try {
    const rendered = render(isolated); rendered.warnings.forEach(w => warnings.add(w));
    let active = -1, from = 0;
    for (const p of elementRanges(rendered.html, 'p')) {
      const text = entities(p.inner.replace(/<[^>]*>/g, '')).trim();
      const start = definitions.findIndex((_, i) => text === `${prefix}${i}-START`);
      if (start >= 0) { active = start; from = p.end; }
      else if (active >= 0 && text === `${prefix}${active}-END`) {
        const d = definitions[active];
        result.set(d.key, { html: rendered.html.slice(from, p.start).trim(), side: d.side, gap: d.gap }); active = -1;
      }
    }
    if (result.size !== definitions.length) throw Error('일부 표 캡션의 본문을 변환하지 못했습니다.');
    return { captions: result, warnings: [...warnings] };
  } finally { isolated.free(); }
}

/** Geometry must finish before inserting caption paragraphs. Otherwise their
 * paragraphs would shift source cell indices and corrupt following formats. */
export function restoreTableCaptions(html: string, captions: Map<string, { html: string; side: string; gap: number }>, warnings: Set<string>): string {
  const used = new Set<string>();
  const withGap = (content: string, gap: number, side: string) => {
    if (!gap || !content) return content;
    const prop = side === 'BOTTOM' ? 'margin-top' : 'margin-bottom';
    const blocks = [...elementRanges(content, 'p'), ...elementRanges(content, 'table')].sort((a, b) => a.start - b.start);
    const block = side === 'BOTTOM' ? blocks[0] : blocks.at(-1); if (!block) return content;
    const oldStyle = block.open.match(/\bstyle="([^"]*)"/), css = oldStyle?.[1] || '';
    const old = [...css.matchAll(new RegExp(`(?:^|;)${prop}:([\\d.]+)px`, 'g'))].at(-1)?.[1];
    const extra = `${prop}:${Math.round((Number(old || 0) + gap) * 100) / 100}px;`;
    const open = oldStyle ? block.open.replace(oldStyle[0], `style="${css};${extra}"`) : block.open.replace(/>$/, ` style="${extra}">`);
    return content.slice(0, block.start) + open + content.slice(block.start + block.open.length);
  };
  const visit = (source: string): string => {
    let out = '', at = 0;
    for (const table of elementRanges(source, 'table')) {
      const key = table.open.match(/\bdata-hwp-caption-key="([^"]+)"/)?.[1], caption = key && captions.get(key);
      const body = table.open.replace(/\sdata-hwp-caption-key="[^"]*"/, '') + visit(table.inner) + '</table>';
      let before = '', after = '';
      if (caption) {
        used.add(key!); const content = withGap(caption.html, caption.gap, caption.side);
        if (caption.side === 'BOTTOM') after = content;
        else before = content;
        if (caption.side !== 'TOP' && caption.side !== 'BOTTOM') warnings.add('표의 좌우 캡션은 본문 흐름으로 가져옵니다. 원본 배치를 확인하세요.');
      }
      out += source.slice(at, table.start) + before + body + after; at = table.end;
    }
    return out + source.slice(at);
  };
  const result = visit(html);
  if (used.size !== captions.size) throw Error('일부 표 캡션의 위치를 연결하지 못했습니다. 원본 파일은 유지됩니다.');
  return result;
}
