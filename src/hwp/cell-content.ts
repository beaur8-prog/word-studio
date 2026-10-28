import type { HwpDocument } from '@rhwp/core';
import { strFromU8 } from 'fflate';
import { serializedParts } from './serialized-parts';
import { hwpBookmark, hwpTabAttribute, hwpTabDefinitions } from './text-navigation';
import type { DocumentTabStop } from '../document-tabs';

export type CellPath = { controlIndex: number; cellIndex: number; cellParaIndex: number }[];
type Equation = { script: string; fontSize: number; color: number; width: number; height: number };
export type SerializedControl = { offset: number; kind: string; numberType?: string; emptyHeader?: boolean; binaryItem?: string; noteId?: string; noteLabel?: string; bookmarkName?: string };
type CellContent = { equations: Equation[]; controls: SerializedControl[]; breaks: number[]; tabStops?: DocumentTabStop[] };
type Paragraph = { key: string; tables: number; cell: boolean; offset: number; tabRef: string; width: number; tabCell?: Cell };
type Cell = { key: string; paragraphs: number; width?: number };
type Table = { key: string; cells: number };
type Frame = { name: string; paragraph?: Paragraph; table?: Table; cell?: Cell; equation?: Equation; control?: SerializedControl };
const fragment = (html: string) => html.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/)?.[1].trim() ?? html;
export const entities = (text: string) => text.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[\da-f]+);/gi, value => {
  const names: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" };
  if (names[value]) return names[value];
  const n = value[2].toLowerCase() === 'x' ? parseInt(value.slice(3, -1), 16) : Number(value.slice(2, -1));
  return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : '\ufffd';
});

/** The flat equation API exposes only the first equation of a cell paragraph.
 * Read the engine's own serialization to preserve every equation, including
 * nested cells, without changing the document or guessing from display text. */
export function cellContents(doc: HwpDocument, files = serializedParts(doc)): Map<string, CellContent> {
  const result = new Map<string, CellContent>();
  const tabs = hwpTabDefinitions(files);
  for (const [name, bytes] of Object.entries(files)) {
    if (!/^Contents\/section\d+\.xml$/.test(name)) continue;
    const section = Number(name.match(/section(\d+)/)![1]), stack: Frame[] = [];
    const pendingTabs = new Map<string, Paragraph>();
    const pd = JSON.parse(doc.getPageDef(section));
    const width = ((pd.landscape ? pd.height : pd.width) - pd.marginLeft - pd.marginRight) / 75;
    const sectionXml = strFromU8(bytes), spacing = Number(sectionXml.match(/<hp:secPr\b[^>]*\btabStop="([^"]*)"/)?.[1]) || 8000;
    let paragraph = -1, note = 0;
    for (const token of strFromU8(bytes).matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<[^>]+>|[^<]+/g)) {
      const text = token[0], parent = stack.at(-1);
      if (text.startsWith('<!--') || text.startsWith('<?')) continue;
      if (text[0] !== '<' || text.startsWith('<![CDATA[')) {
        const decoded = text.startsWith('<![CDATA[') ? text.slice(9, -3) : entities(text);
        if (parent?.control?.kind === 'header' && decoded.trim()) parent.control.emptyHeader = false;
        if (parent?.name === 'script' && parent.equation) parent.equation.script += decoded;
        if (parent?.name === 't' && stack.at(-2)?.name === 'run' && parent.paragraph) parent.paragraph.offset += decoded.length;
        continue;
      }
      if (text.startsWith('</')) { stack.pop(); continue; }
      const name = text.match(/^<(?:[\w.-]+:)?([\w.-]+)/)?.[1]; if (!name) continue;
      const attr = (key: string) => text.match(new RegExp(`\\b${key}=["']([^"']*)["']`))?.[1] ?? '';
      const frame: Frame = { ...parent, name };
      if (name === 'p') {
        if (parent?.name === 'sec') {
          frame.paragraph = { key: `${section}:${++paragraph}`, tables: 0, cell: false, offset: 0, tabRef: attr('paraPrIDRef'), width };
          result.set(frame.paragraph.key, { equations: [], controls: [], breaks: [] });
        }
        else if (parent?.name === 'subList' && stack.at(-2)?.name === 'tc' && parent.cell) {
          frame.paragraph = { key: `${parent.cell.key}/p${parent.cell.paragraphs++}`, tables: 0, cell: true, offset: 0, tabRef: attr('paraPrIDRef'), width, tabCell: parent.cell };
          result.set(frame.paragraph.key, { equations: [], controls: [], breaks: [] });
        }
        else frame.paragraph = undefined; // Header/footer and textbox stories.
      } else if (name === 'tbl' && frame.paragraph) {
        frame.table = { key: `${frame.paragraph.key}/t${frame.paragraph.tables++}`, cells: 0 };
      } else if (name === 'tc' && frame.table) {
        frame.cell = { key: `${frame.table.key}/c${frame.table.cells++}`, paragraphs: 0 };
      } else if (name === 'cellSz' && frame.cell) {
        frame.cell.width = Number(attr('width')) / 75;
      } else if (name === 'equation' && frame.paragraph?.cell) {
        const equation = { script: '', fontSize: Number(attr('baseUnit')) || 1200, color: parseInt(attr('textColor').replace('#', ''), 16) || 0, width: 0, height: 0 };
        result.get(frame.paragraph.key)!.equations.push(equation); frame.equation = equation;
      } else if (name === 'sz' && parent?.name === 'equation' && frame.equation) {
        frame.equation.width = Number(attr('width')); frame.equation.height = Number(attr('height'));
      }
      if (parent?.control?.kind === 'header' && ['rect', 'line', 'pic', 'ellipse', 'polygon', 'container', 'tbl', 'equation', 'ole', 'autoNum', 'fieldBegin'].includes(name)) parent.control.emptyHeader = false;
      if (parent?.paragraph && ((parent.name === 'run' && !['t', 'ctrl'].includes(name)) || parent.name === 'ctrl')) {
        const control: SerializedControl = { offset: parent.paragraph.offset, kind: name };
        if (name === 'autoNum') control.numberType = attr('numType');
        if (name === 'header') control.emptyHeader = true;
        if (name === 'ole') control.binaryItem = attr('binaryItemIDRef');
        if (name === 'bookmark') control.bookmarkName = entities(attr('name'));
        if (name === 'footNote' || name === 'endNote') { control.noteId = `hwp-note-${section}-${note++}`; control.noteLabel = attr('number') || '1'; }
        result.get(parent.paragraph.key)!.controls.push(control); frame.control = control;
      }
      if (parent?.name === 't' && parent.paragraph && ['tab', 'lineBreak', 'hyphen', 'nbSpace', 'fwSpace'].includes(name)) {
        if (name === 'tab') pendingTabs.set(parent.paragraph.key, parent.paragraph);
        if (name === 'lineBreak' && parent.paragraph.cell) result.get(parent.paragraph.key)!.breaks.push(parent.paragraph.offset);
        parent.paragraph.offset++;
      }
      if (!text.endsWith('/>')) stack.push(frame);
    }
    // HWPX writes cellSz after the cell's paragraphs. Resolve those widths
    // only after the whole section, rather than using the outer paper width.
    for (const p of pendingTabs.values()) result.get(p.key)!.tabStops = tabs(p.tabRef, spacing, p.tabCell?.width || p.width);
  }
  return result;
}

/** Restore omitted controls and inline pictures appended outside their source
 * paragraph by the clipboard exporter. Ordinary paragraphs keep the original
 * HTML. Cell paths distinguish nested tables and separate document stories. */
export function restoreCellContent(doc: HwpDocument, section: number, para: number, tableControl: number, tableOrdinal: number, html: string,
  equation: (value: Equation) => string, shape: (path: CellPath, index: number) => string | undefined, warnings: Set<string>, metadata: () => Map<string, CellContent>, headerPreserved = false,
  picture?: (path: CellPath, index: number, html: string) => string, note?: (value: SerializedControl) => string | undefined) {
  const annotate = (content: string, paragraph: number, key: string) => {
    let tableDepth = 0;
    return content.replace(/<\/?(?:table|p)\b[^>]*>/g, tag => {
      if (/^<table\b/.test(tag)) tableDepth++;
      else if (tag === '</table>') tableDepth--;
      else if (!tableDepth && /^<p\b/.test(tag)) return tag.replace(/>$/, ` data-hwp-cell-paragraph="${paragraph}"${hwpTabAttribute(metadata().get(key)?.tabStops)}>`);
      return tag;
    });
  };
  const repair = (source: string, parents: CellPath, control: number, key: string, depth = 0): string => {
    if ((!/<!--\s*rhwp:|<img\b/.test(source) && !/<table\b[\s\S]*<table\b/.test(source)) || depth > 12) return source;
    const tablePath = [...parents, { controlIndex: control, cellIndex: 0, cellParaIndex: 0 }];
    const dims = parents.length ? JSON.parse(doc.getTableDimensionsByPath(section, para, JSON.stringify(tablePath))) : JSON.parse(doc.getTableDimensions(section, para, control));
    if (dims.cellCount > 50000) throw Error('편집할 표의 셀 수가 너무 많습니다.');
    const cells = Array.from({ length: dims.cellCount }, (_, index) => ({ index, ...(parents.length
      ? JSON.parse(doc.getCellInfoByPath(section, para, JSON.stringify([...parents, { controlIndex: control, cellIndex: index, cellParaIndex: 0 }])))
      : JSON.parse(doc.getCellInfo(section, para, control, index))) })).sort((a, b) => a.row - b.row || a.col - b.col);
    let tableDepth = 0, start = 0, cell = 0, output = '', cursor = 0;
    for (const match of source.matchAll(/<\/?(?:table|td)\b[^>]*>/g)) {
      const tag = match[0];
      if (/^<table\b/.test(tag)) { tableDepth++; continue; }
      if (tag === '</table>') { tableDepth--; continue; }
      if (tableDepth !== 1) continue;
      if (/^<td\b/.test(tag)) { start = match.index! + tag.length; continue; }
      const original = source.slice(start, match.index), current = cells[cell++];
      if (!current || !/<!--\s*rhwp:|<(?:img|table)\b/.test(original)) continue;
      const path = [...parents, { controlIndex: control, cellIndex: current.index, cellParaIndex: 0 }];
      const paragraphs = doc.getCellParagraphCountByPath(section, para, JSON.stringify(path));
      const parts: string[] = [];
      for (let p = 0; p < paragraphs; p++) {
        path[path.length - 1].cellParaIndex = p;
        const json = JSON.stringify(path), paragraphKey = `${key}/c${current.index}/p${p}`;
        let content = fragment(doc.exportSelectionInCellHtmlByPath(section, para, json, p, 0, p, 0xffffffff));
        const needsRepair = /<!--\s*rhwp:|<table\b/.test(content) || (/<img\b/.test(content) && (!!picture || !!content.replace(/<[^>]*>/g, '').trim()));
        if (needsRepair) {
          const controls: { index: number; type: string; html: string }[] = [];
          for (let index = 0; index < 1024; index++) {
            try {
              const html = fragment(doc.exportControlHtml(section, para, json, index));
              controls.push({ index, html: (/^<img\b/.test(html) && picture ? picture(path, index, html) : html).replace(/<img\b(?![^>]*\bdata-inline-image=)/g, '<img data-inline-image="true"'),
                type: /^<table\b/.test(html) ? 'Table' : html.match(/^<!--\s*rhwp:\s*셀 안 (\w+)/)?.[1] ?? '' });
            } catch { break; }
          }
          const data = metadata().get(paragraphKey);
          const known = (value: SerializedControl | undefined): string | undefined => {
            if (value?.bookmarkName) return hwpBookmark(value.bookmarkName);
            if (value?.noteId) return note?.(value);
            if (value?.kind === 'autoNum' && ['PAGE', 'TOTAL_PAGE'].includes(value.numberType ?? ''))
              return `<span data-page-number="${value.numberType === 'PAGE' ? 'PAGE' : 'NUMPAGES'}">1</span>`;
            // An empty header definition is not a visible body object. Rich
            // headers stay explicit until their separate story is preserved.
            if (value?.kind === 'header' && (value.emptyHeader || headerPreserved)) return '';
          };
          if (controls.length && !controls.some(c => c.type === 'Table') && data?.controls.length === controls.length) {
            // Selection HTML appends omitted controls after all paragraph text.
            // Reinsert at the serialized model's logical text offsets instead
            // of trusting that placeholder order as the inline position.
            let eq = 0;
            const replacements = controls.map((c, i) => ({ offset: data.controls[i].offset, html: c.type === 'Equation' && data.equations[eq]
              ? equation(data.equations[eq++]) : c.type === 'Shape' ? shape(path, c.index) ?? c.html : known(data.controls[i]) ?? c.html }));
            const paragraph = content.match(/<p\b[^>]*>[\s\S]*?<\/p>/)?.[0] ?? '<p></p>';
            let offset = 0, next = 0, nextBreak = 0, output = '', spanDepth = 0;
            const emit = () => {
              while (true) {
                while (next < replacements.length && replacements[next].offset <= offset) output += replacements[next++].html;
                if (nextBreak >= data.breaks.length || data.breaks[nextBreak] > offset) break;
                output += '<br>'; nextBreak++; offset++;
              }
            };
            for (const m of paragraph.matchAll(/<[^>]+>|&(?:#\d+|#x[\da-f]+|\w+);|[^<&]+|[<&]/gi)) {
              const value = m[0];
              if (value[0] === '<') {
                if (/^<span\b/.test(value)) spanDepth++;
                if (/^<\/span\b/.test(value)) spanDepth--;
                if (/^<\/p/.test(value)) { emit(); while (next < replacements.length) output += replacements[next++].html; }
                // Native selection HTML may omit hard line breaks. Restore
                // their logical slots alongside controls from serialization.
                if (!/^<(?:img|br)\b|^<!--/.test(value)) output += value;
                continue;
              }
              if (!spanDepth && /^[\t\r\n ]*[\r\n][\t\r\n ]*$/.test(value)) continue;
              if (value[0] === '&') { emit(); output += value; offset += entities(value).length; }
              else for (const char of value) { emit(); output += char; offset += char.length; }
            }
            parts.push(annotate(output, p, paragraphKey)); continue;
          }
          let next = 0, eq = 0, nested = 0;
          const tableRanges: [number, number][] = [];
          let level = 0, tableStart = 0;
          for (const m of content.matchAll(/<\/?table\b[^>]*>/g)) {
            if (/^<table\b/.test(m[0])) { if (!level++) tableStart = m.index!; }
            else if (--level === 0) tableRanges.push([tableStart, m.index! + m[0].length]);
          }
          content = content.replace(/<!--\s*rhwp:\s*셀 안 (\w+)[\s\S]*?-->/g, (comment, type: string, offset: number) => {
            if (tableRanges.some(([from, to]) => offset >= from && offset < to)) return comment;
            const index = controls.findIndex((c, i) => i >= next && c.type === type);
            if (index < 0) return comment; next = index + 1;
            if (data?.controls.length === controls.length) {
              const restored = known(data.controls[index]); if (restored !== undefined) return restored;
            }
            if (type === 'Equation') {
              const value = data?.equations[eq++];
              if (value) return equation(value);
            }
            if (type === 'Shape') return shape(path, controls[index].index) ?? comment;
            return comment;
          });
          // Nested tables are already present in selection HTML. Replace each
          // complete outer table without confusing a nested TD with this cell.
          let nesting = 0, from = 0, at = 0, fixed = '';
          for (const m of content.matchAll(/<\/?table\b[^>]*>/g)) {
            if (/^<table\b/.test(m[0])) { if (!nesting++) from = m.index!; }
            else if (--nesting === 0) {
              const table = controls.filter(c => c.type === 'Table')[nested];
              const end = m.index! + m[0].length;
              fixed += content.slice(at, from) + (table ? repair(content.slice(from, end), path, table.index, `${paragraphKey}/t${nested}`, depth + 1) : content.slice(from, end));
              at = end; nested++;
            }
          }
          content = fixed + content.slice(at);
        }
        // Empty paragraphs also disappear from native selection HTML. Keep
        // their slots: tableGeometry applies the source paragraph formats by
        // index, so losing one gives every following picture/text the wrong
        // alignment, leading and indent.
        if (!content.trim()) content = '<p></p>';
        // Control-only paragraphs have no P in the clipboard HTML. Give their
        // inline previews the same paragraph slot as the original source.
        if (needsRepair && /^<img\b/.test(content)) content = `<p>${content}</p>`;
        // A control-only paragraph can consist solely of a nested table,
        // without an outer P. Carry the source index to following paragraphs
        // so tableGeometry does not shift their formatting by one slot.
        content = annotate(content, p, paragraphKey);
        parts.push(content.replace(/<img\b(?![^>]*\bdata-inline-image=)/g, '<img data-inline-image="true"'));
      }
      output += source.slice(cursor, start) + parts.join('\n'); cursor = match.index!;
    }
    return output + source.slice(cursor);
  };
  try { return repair(html, [], tableControl, `${section}:${para}/t${tableOrdinal}`); }
  catch { warnings.add('일부 표 안 개체를 복원하지 못했습니다. 원본 보기의 해당 개체와 비교하세요.'); return html; }
}
