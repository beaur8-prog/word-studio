import { boxedDigitHtml } from '../hwp-boxed-digit';
import { circledComposition, circledCompositionHtml } from './circled-composition';
import type { HwpDocument } from '@rhwp/core';
import { strFromU8 } from 'fflate';
import { serializedParts } from './serialized-parts';
import { hwpBookmark, hwpPageReferenceLink, hwpTabAttribute, hwpTabDefinitions } from './text-navigation';
import type { DocumentTabStop } from '../document-tabs';
import { preserveHwpSourceSpaces, restoreScriptFonts, sourceFontRuns, type HwpScriptTypography } from './script-fonts';
import { restoreHwpFixedSpaces } from './fixed-spaces';

export type CellPath = { controlIndex: number; cellIndex: number; cellParaIndex: number }[];
type Equation = { script: string; fontSize: number; color: number; width: number; height: number };
type PageReference = { target: string; start: number; end: number };
export type SerializedControl = { offset: number; kind: string; numberType?: string; emptyHeader?: boolean; binaryItem?: string; noteId?: string; noteLabel?: string; bookmarkName?: string; fieldType?: string; fieldId?: string; boxedDigit?: string; circledNumber?: string; fieldArgument?: string };
export type SerializedTablePolicy = { pageBreak: 'NONE' | 'TABLE' | 'CELL'; repeatHeader: boolean; treatAsChar?: boolean };
type CellContent = { equations: Equation[]; controls: SerializedControl[]; breaks: number[]; pageReference?: PageReference; fixedSpaces?: number[]; tabStops?: DocumentTabStop[]; tablePolicies?: (SerializedTablePolicy | undefined)[] };
type Paragraph = { key: string; tables: number; cell: boolean; offset: number; tabRef: string; width: number; tabCell?: Cell };
type Cell = { key: string; paragraphs: number; width?: number };
type Table = { key: string; cells: number; policy?: SerializedTablePolicy };
type Frame = { name: string; paragraph?: Paragraph; table?: Table; cell?: Cell; equation?: Equation; control?: SerializedControl; referenceArgument?: SerializedControl };
const fragment = (html: string) => html.match(/<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/)?.[1].trim() ?? html;
export const entities = (text: string) => text.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[\da-f]+);/gi, value => {
  const names: Record<string, string> = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" };
  if (names[value]) return names[value];
  const n = value[2].toLowerCase() === 'x' ? parseInt(value.slice(3, -1), 16) : Number(value.slice(2, -1));
  return n > 0 && n <= 0x10ffff ? String.fromCodePoint(n) : '\ufffd';
});

/** The clipboard provides the cached digit run and emits the unsupported
 * Field comment separately. Retain the exact text/style and link just the
 * authored cached page number to its uniquely identified native bookmark. */
function cellPageReferenceHtml(content: string, source: string, reference: PageReference) {
  const markers = /<!--\s*rhwp:\s*셀 안 Field\b[\s\S]*?-->/g;
  if ([...content.matchAll(markers)].length !== 1 || reference.end !== source.length) return;
  const plain = content.replace(markers, '').trim();
  const paragraph = /^(<p\b[^>]*>)([\s\S]*?)(<\/p>)$/.exec(plain);
  const cache = source.slice(reference.start, reference.end);
  if (!paragraph || !/^[0-9]{1,6}$/.test(cache) ||
      /<\/?(?:a|p|table|img|br|svg)\b/i.test(paragraph[2])) return;
  let offset = 0, depth = 0, literal = '', updated = '', linked = false;
  for (const [part] of paragraph[2].matchAll(/<[^>]*>|[^<]+/g)) {
    if (part.startsWith('<')) {
      if (part.startsWith('<span ')) depth++;
      if (part === '</span>') depth--;
      updated += part;
      continue;
    }
    if (!depth && /^[\t\r\n ]*[\r\n][\t\r\n ]*$/.test(part)) { updated += part; continue; }
    const decoded = entities(part.replace(/&nbsp;/g, '\u00a0'));
    if (offset <= reference.start && part.endsWith(cache) && reference.end - offset === decoded.length) {
      updated += part.slice(0, -cache.length) + hwpPageReferenceLink(reference.target, cache);
      linked = true;
    } else updated += part;
    offset += decoded.length; literal += decoded;
  }
  return linked && literal === source ? paragraph[1] + updated + paragraph[3] : undefined;
}

/** Native selection HTML can omit explicit line-break characters. Offsets
 * come from the logical source text, never from cached visual line wrapping. */
export function restoreHardLineBreaks(html: string, breaks: readonly number[]): string {
  if (!breaks.length) return html;
  let offset = 0, next = 0, output = '', spanDepth = 0;
  const emit = () => {
    while (next < breaks.length && breaks[next] <= offset) { output += '<br>'; next++; offset++; }
  };
  for (const token of html.matchAll(/<!--[\s\S]*?-->|<[^>]+>|&(?:#\d+|#x[\da-f]+|\w+);|[^<&]+|[<&]/gi)) {
    const value = token[0];
    if (value[0] === '<') {
      if (/^<span\b/.test(value)) spanDepth++;
      if (/^<\/span\b/.test(value)) spanDepth--;
      if (/^<br\b/i.test(value)) {
        if (breaks[next] === offset) next++;
        output += value; offset++; continue;
      }
      if (/^<\/p\b/.test(value)) emit();
      output += value; continue;
    }
    if (!spanDepth && /^[\t\r\n ]*[\r\n][\t\r\n ]*$/.test(value)) { output += value; continue; }
    if (value[0] === '&') { emit(); output += value; offset += value === '&nbsp;' ? 1 : entities(value).length; }
    else for (const char of value) {
      if (char === '\n' && breaks[next] === offset) { output += '<br>'; next++; offset++; continue; }
      emit(); output += char; offset += char.length;
    }
  }
  return output;
}

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
        if (parent?.name === 'stringParam' && parent.referenceArgument && parent.referenceArgument.fieldArgument !== '!') {
          const argument = (parent.referenceArgument.fieldArgument ?? '') + decoded;
          parent.referenceArgument.fieldArgument = argument.length <= 1024 ? argument : '!';
        }
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
        else {
          // Captions, headers and textboxes have independent paragraph/table
          // counters. Their nested cells must not advance the owning cell.
          frame.paragraph = undefined; frame.table = undefined; frame.cell = undefined;
        }
      } else if (name === 'tbl' && frame.paragraph) {
        const ordinal = frame.paragraph.tables++;
        frame.table = { key: `${frame.paragraph.key}/t${ordinal}`, cells: 0 };
        const pageBreak = attr('pageBreak');
        if (pageBreak === 'NONE' || pageBreak === 'TABLE' || pageBreak === 'CELL') {
          const data = result.get(frame.paragraph.key)!;
          const policy: SerializedTablePolicy = { pageBreak, repeatHeader: attr('repeatHeader') === '1' };
          (data.tablePolicies ??= [])[ordinal] = policy; frame.table.policy = policy;
        }
      } else if (name === 'pos' && parent?.name === 'tbl' && frame.table?.policy) {
        const inline = attr('treatAsChar');
        if (inline === '0' || inline === '1') frame.table.policy.treatAsChar = inline === '1';
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
      if (name === 'stringParam' && parent?.name === 'parameters' && stack.at(-2)?.name === 'fieldBegin' &&
          parent.control?.fieldType === 'CROSSREF' && attr('name') === 'Command' && !text.endsWith('/>')) {
        if (parent.control.fieldArgument === undefined) {
          parent.control.fieldArgument = '';
          frame.referenceArgument = parent.control;
        } else parent.control.fieldArgument = '!';
      }
      if (parent?.control?.kind === 'header' && ['rect', 'line', 'pic', 'ellipse', 'polygon', 'container', 'tbl', 'equation', 'ole', 'autoNum', 'fieldBegin'].includes(name)) parent.control.emptyHeader = false;
      if (parent?.paragraph && ((parent.name === 'run' && !['t', 'ctrl'].includes(name)) || parent.name === 'ctrl')) {
        const control: SerializedControl = { offset: parent.paragraph.offset, kind: name };
        if (name === 'autoNum') control.numberType = attr('numType');
        if (name === 'header') control.emptyHeader = true;
        if (name === 'ole') control.binaryItem = attr('binaryItemIDRef');
        if (name === 'compose' && attr('composeType') === 'SPREAD' && attr('charSz') === '-3') {
          const kind = attr('circleType'), value = entities(attr('composeText'));
          if (kind === 'SHAPE_RECTANGLE' && /^[0-9]$/.test(value)) control.boxedDigit = value;
          if (kind === 'SHAPE_CIRCLE') control.circledNumber = circledComposition(value);
        }
        if (name === 'bookmark') control.bookmarkName = entities(attr('name'));
        if (name === 'fieldBegin') { control.fieldType = attr('type'); control.fieldId = attr('id'); if (control.fieldType === 'BOOKMARK') control.bookmarkName = entities(attr('name')); }
        if (name === 'fieldEnd') control.fieldId = attr('beginIDRef');
        if (name === 'footNote' || name === 'endNote') { control.noteId = `hwp-note-${section}-${note++}`; control.noteLabel = attr('number') || '1'; }
        result.get(parent.paragraph.key)!.controls.push(control); frame.control = control;
      }
      if (parent?.name === 't' && parent.paragraph && ['tab', 'lineBreak', 'hyphen', 'nbSpace', 'fwSpace'].includes(name)) {
        if (name === 'tab') pendingTabs.set(parent.paragraph.key, parent.paragraph);
        if (name === 'fwSpace') (result.get(parent.paragraph.key)!.fixedSpaces ??= []).push(parent.paragraph.offset);
        if (name === 'lineBreak' && parent.paragraph.cell) result.get(parent.paragraph.key)!.breaks.push(parent.paragraph.offset);
        parent.paragraph.offset++;
      }
      if (!text.endsWith('/>')) stack.push(frame);
    }
    // HWPX writes cellSz after the cell's paragraphs. Resolve those widths
    // only after the whole section, rather than using the outer paper width.
    for (const p of pendingTabs.values()) result.get(p.key)!.tabStops = tabs(p.tabRef, spacing, p.tabCell?.width || p.width);
  }
  // The native control API omits range ends. Only a verified BOOKMARK end
  // may be excluded from index matching; keep unknown field types explicit.
  // IDs may be reused in a later section.
  const bookmarks = new Set([...result].flatMap(([key, value]) => value.controls
    .filter(c => c.kind === 'fieldBegin' && c.fieldType === 'BOOKMARK' && c.fieldId)
    .map(c => `${key.split(':')[0]}:${c.fieldId}`)));
  const destinations = new Map<string, number>();
  for (const value of result.values()) for (const control of value.controls) {
    const name = control.bookmarkName;
    if (name) destinations.set(name, (destinations.get(name) ?? 0) + 1);
  }
  for (const [key, value] of result) {
    const [begin, end] = value.controls;
    if (!key.includes('/c') || value.controls.length !== 2 || begin.kind !== 'fieldBegin' ||
        begin.fieldType !== 'CROSSREF' || end.kind !== 'fieldEnd' || !begin.fieldId ||
        begin.fieldId !== end.fieldId || end.offset <= begin.offset) continue;
    const target = /^\?([^;\u0000-\u001f]{1,1000});6;0;1;0;$/.exec(begin.fieldArgument ?? '')?.[1];
    if (target && destinations.get(target) === 1)
      value.pageReference = { target, start: begin.offset, end: end.offset };
  }
  for (const [key, value] of result) value.controls = value.controls.filter(c =>
    !(c.kind === 'fieldEnd' && bookmarks.has(`${key.split(':')[0]}:${c.fieldId}`)));
  return result;
}

/** Restore omitted controls and inline pictures appended outside their source
 * paragraph by the clipboard exporter. Ordinary paragraphs keep the original
 * HTML. Cell paths distinguish nested tables and separate document stories. */
export function restoreCellContent(doc: HwpDocument, section: number, para: number, tableControl: number, tableOrdinal: number, html: string,
  equation: (value: Equation) => string, shape: (path: CellPath, index: number) => string | undefined, warnings: Set<string>, metadata: () => Map<string, CellContent>, headerPreserved = false,
  picture?: (path: CellPath, index: number, html: string) => string, note?: (value: SerializedControl) => string | undefined,
  nestedTable?: (path: CellPath, html: string, policy?: SerializedTablePolicy) => string, caption?: (key: string, html: string) => string, fontSpace?: readonly boolean[]) {
  const breakKeys = [...metadata()].filter(([, value]) => value.breaks.length).map(([key]) => key);
  const hasBreaks = (prefix: string) => breakKeys.some(key => key.startsWith(prefix + '/'));
  const fontCache = new Map<number, HwpScriptTypography>();
  const scriptFonts = (content: string, path: CellPath, key: string) => {
    // Do not tokenize inline object payloads or guess text/control offsets.
    if (/<(?:img|table)\b|<!--\s*rhwp:/i.test(content)) return content;
    try {
      const json = JSON.stringify(path), text = doc.getTextInCellByPath(section, para, json, 0, 0xffffffff);
      if (!text) return content;
      const runs = sourceFontRuns(doc.getCharShapeRunsInCellByPath(section, para, json, 0, [...text].length), 0, text,
        offset => JSON.parse(doc.getCellCharPropertiesAtByPath(section, para, json, offset)), fontCache);
      const scripts = preserveHwpSourceSpaces(restoreScriptFonts(content, text, runs), text, runs, fontSpace,
        () => doc.getParagraphLineMetrics(section, para, json));
      return restoreHwpFixedSpaces(scripts, text, metadata().get(key)?.fixedSpaces ?? []);
    } catch {
      // Font recovery is optional. An unsupported character shape must not
      // discard controls, fields or pictures already repaired in this table.
      return content;
    }
  };
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
    const finish = (html: string) => {
      const owner = key.match(/^(.*)\/t(\d+)$/);
      const policy = owner ? metadata().get(owner[1])?.tablePolicies?.[Number(owner[2])] : undefined;
      const geometry = parents.length && nestedTable && depth <= 12
        ? nestedTable([...parents, { controlIndex: control, cellIndex: 0, cellParaIndex: 0 }], html, policy) : html;
      return caption?.(key, geometry) ?? geometry;
    };
    if (depth > 12) return finish(source);
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
      if (!current) continue;
      const path = [...parents, { controlIndex: control, cellIndex: current.index, cellParaIndex: 0 }];
      const paragraphs = doc.getCellParagraphCountByPath(section, para, JSON.stringify(path));
      if (!/<!--\s*rhwp:|<(?:img|table)\b/.test(original) && !hasBreaks(`${key}/c${current.index}`)) {
        // Preserve the original clipboard HTML in ordinary cells. Re-exporting
        // it just for fonts could lose omitted empty source paragraph slots.
        const matches = original.match(/<p\b[^>]*>[\s\S]*?<\/p>/g) || [];
        if (matches.length !== paragraphs) continue;
        let p = 0;
        const formatted = original.replace(/<p\b[^>]*>[\s\S]*?<\/p>/g, content => {
          const index = p++, paragraphKey = `${key}/c${current.index}/p${index}`;
          path[path.length - 1].cellParaIndex = index;
          const formatted = scriptFonts(content, path, paragraphKey);
          // Top-level table geometry supplies flat cell tabs separately, but
          // plain nested cells do not visit that callback. Keep their source
          // paragraph stops too, without rebuilding the clipboard contents.
          return metadata().get(paragraphKey)?.tabStops?.length ? annotate(formatted, index, paragraphKey) : formatted;
        });
        if (formatted !== original) { output += source.slice(cursor, start) + formatted; cursor = match.index!; }
        continue;
      }
      const parts: string[] = [];
      for (let p = 0; p < paragraphs; p++) {
        path[path.length - 1].cellParaIndex = p;
        const json = JSON.stringify(path), paragraphKey = `${key}/c${current.index}/p${p}`;
        let content = fragment(doc.exportSelectionInCellHtmlByPath(section, para, json, p, 0, p, 0xffffffff));
        const needsRepair = /<!--\s*rhwp:|<table\b/.test(content) || (/<img\b/.test(content) && (!!picture || !!content.replace(/<[^>]*>/g, '').trim()));
        if (!needsRepair) content = restoreHardLineBreaks(content, metadata().get(paragraphKey)?.breaks || []);
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
          if (data?.pageReference && controls.length === 1 && controls[0].type === 'Field') {
            const linked = cellPageReferenceHtml(content, doc.getTextInCellByPath(section, para, json, 0, 0xffffffff), data.pageReference);
            if (linked) {
              warnings.add('한글 쪽 참조를 책갈피 링크로 가져오며 저장된 쪽 번호를 표시합니다.');
              parts.push(annotate(scriptFonts(linked, path, paragraphKey), p, paragraphKey));
              continue;
            }
          }
          const known = (value: SerializedControl | undefined, type: string): string | undefined => {
            if (type === 'CharOverlap' && value && (value.boxedDigit || value.circledNumber)) {
              const char = JSON.parse(doc.getCellCharPropertiesAtByPath(section, para, json, value.offset));
              return value.circledNumber ? circledCompositionHtml(value.circledNumber, char.fontSize)
                : boxedDigitHtml(value.boxedDigit, char.fontSize);
            }
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
              ? equation(data.equations[eq++]) : c.type === 'Shape' ? shape(path, c.index) ?? c.html : known(data.controls[i], c.type) ?? c.html }));
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
            parts.push(annotate(scriptFonts(output, path, paragraphKey), p, paragraphKey)); continue;
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
              const restored = known(data.controls[index], type); if (restored !== undefined) return restored;
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
          if (controls.length === 1 && controls[0].type === 'Table' &&
              /^<p\b[^>]*>[\s\S]*?<\/p>\s*<table\b/.test(content) &&
              /^ +$/.test(doc.getTextInCellByPath(section, para, json, 0, 0xffffffff))) {
            const { stored, controlCount, lines, tables } = JSON.parse(doc.getParagraphLineMetrics(section, para, json));
            if (stored && controlCount === 1 && lines.length === 1 && tables.some((table: { index: number; inline: boolean }) =>
              table.index === controls[0].index && table.inline)) {
              content = content.replace(/^<p\b[^>]*>/, tag => tag.replace(/>$/, ' data-hwp-inline-table-spaces="true">'));
            }
          }
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
        content = annotate(scriptFonts(content, path, paragraphKey), p, paragraphKey);
        parts.push(content.replace(/<img\b(?![^>]*\bdata-inline-image=)/g, '<img data-inline-image="true"'));
      }
      output += source.slice(cursor, start) + parts.join('\n'); cursor = match.index!;
    }
    return finish(output + source.slice(cursor));
  };
  try { return repair(html, [], tableControl, `${section}:${para}/t${tableOrdinal}`); }
  catch { warnings.add('일부 표 안 개체를 복원하지 못했습니다. 원본 보기의 해당 개체와 비교하세요.'); return html; }
}
