import type { HwpDocument } from '@rhwp/core';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import documentLimits from '../../electron/document-limits.json';
import { entities } from './cell-content';
import { elementRanges } from './markup-ranges';

type Cell = { row: number; rows: number; columns: number; text: string };
type Plan = { cells: Cell[]; heights: (number | undefined)[] };
type TableAddress = { section: number; paragraph: number; global: number; cells: number };
type Box = { x: number; y: number; w: number; h: number };
type CellBox = Box & { cellIdx: number; pageIndex: number };

const attribute = (tag: string, key: string) => tag.match(new RegExp(`\\b${key}="([^"]*)"`))?.[1];
const text = (markup: string) => entities(markup.replace(/<[^>]*>/g, '').replace(/&nbsp;/gi, ' ')).replace(/\s/gu, '');
const size = (tag: string) => Number(attribute(tag, 'style')?.match(/(?:^|;)\s*height\s*:\s*(\d+(?:\.\d+)?)px\s*(?=;|$)/i)?.[1]);
const key = (rows: number, cells: readonly Cell[]) => JSON.stringify([rows, cells]);

function sourcePlan(body: string) {
  if (!body.includes('data-hwp-grid-row="true"') || /<(?:table|img|svg)\b|\bdata-(?:note-reference|document-field|page-number)\b/.test(body)) return;
  const rows = elementRanges(body, 'tr');
  const authored = rows.map(row => size(row.open));
  if (authored.some(height => !Number.isFinite(height) || height <= 0)) return;
  const thin = rows.map(row => row.open.includes('data-hwp-grid-row="true"'));
  const plan: Plan = { cells: [], heights: [] };
  for (const [row, markup] of rows.entries()) {
    const cells = [...elementRanges(markup.inner, 'td'), ...elementRanges(markup.inner, 'th')].sort((a,b) => a.start - b.start);
    for (const cell of cells) {
      const rows = Number(attribute(cell.open,'rowspan') ?? 1), columns = Number(attribute(cell.open,'colspan') ?? 1);
      if (![rows, columns].every(span => Number.isInteger(span) && span > 0) || row + rows > authored.length) return;
      plan.cells.push({ row, rows, columns, text: text(cell.inner) });
      const height = size(cell.open), sum = authored.slice(row, row + rows).reduce((a,b) => a + b, 0);
      // The normal reverse bridge materializes row sizes in cell CSS. A
      // different height can be an authored cell constraint of its own.
      plan.heights.push(Math.abs(height - sum) < .05 && thin.slice(row, row + rows).includes(true)
        ? Math.round(height * 75) : undefined);
    }
  }
  if (plan.heights.some(height => height !== undefined) && plan.cells.some(cell => cell.text)) return { plan, rows: rows.length };
}

function nativeCells(table: string) {
  const rows = elementRanges(table, 'hp:tr');
  const ranges = elementRanges(table, 'hp:tc');
  const cells = ranges.map(cell => {
    const addr = /<hp:cellAddr\b[^>]*\/>/.exec(cell.inner)?.[0] ?? '';
    const span = /<hp:cellSpan\b[^>]*\/>/.exec(cell.inner)?.[0] ?? '';
    const markup = elementRanges(cell.inner, 'hp:t').map(run => run.inner).join('');
    return { row: Number(attribute(addr,'rowAddr')), rows: Number(attribute(span,'rowSpan')),
      columns: Number(attribute(span,'colSpan')), text: text(markup) };
  });
  return { cells, rows, ranges };
}

const overhang = (span: Box, bounds: Box) => Math.max(bounds.x - span.x, bounds.y - span.y,
  span.x + span.w - bounds.x - bounds.w, span.y + span.h - bounds.y - bounds.h, 0);

type Control = { list: number; para: number; controlIndex: number; ctrlId: string };
function paintedTable(doc: HwpDocument, controls: readonly Control[], address: TableAddress) {
  const candidates = controls.filter(c => c.list === 0 && c.para === address.global && c.ctrlId === 'tbl');
  if (candidates.length !== 1) return;
  const control = candidates[0].controlIndex, args = [address.section,address.paragraph,control] as const;
  if (JSON.parse(doc.getTableProperties(...args)).treatAsChar !== true) return;
  const cells: CellBox[] = JSON.parse(doc.getTableCellBboxes(...args));
  if (cells.length !== address.cells || !cells.every(cell => cell.pageIndex === cells[0].pageIndex)) return;
  const info = JSON.parse(doc.getPageInfo(cells[0].pageIndex));
  if (![info.width, info.height, info.marginTop, info.marginHeader, info.marginLeft, info.marginRight, info.marginBottom, info.marginFooter].every(Number.isFinite)) return;
  const paper = { x: 0, y: 0, w: info.width, h: info.height };
  const body = { x: info.marginLeft, y: info.marginTop + info.marginHeader,
    w: info.width - info.marginLeft - info.marginRight,
    h: info.height - info.marginTop - info.marginHeader - info.marginBottom - info.marginFooter };
  type Run = Box & { secIdx: number; parentParaIdx: number; controlIdx: number; cellIdx: number; cellPath?: unknown[]; text: string };
  const nativeRuns: Run[] = JSON.parse(doc.getPageTextLayout(cells[0].pageIndex)).runs;
  const shapes = new Map(cells.map(cell => [cell.cellIdx,cell])), glyphs = new Map<number,string>();
  const overflow = new Map<string, number>();
  for (const run of nativeRuns) {
    if (run.secIdx !== address.section || run.parentParaIdx !== address.paragraph || run.controlIdx !== control ||
        run.cellPath?.length !== 1 || !run.text.trim()) continue;
    const cell = shapes.get(run.cellIdx);
    if (!cell || ![run.x, run.y, run.w, run.h].every(Number.isFinite) || !(run.w > 0 && run.h > 0)) return;
    glyphs.set(run.cellIdx, (glyphs.get(run.cellIdx) ?? '') + run.text.replace(/\s/gu,''));
    for (const [kind, bounds] of [['cell', cell], ['body', body], ['paper', paper]] as const) {
      const id = kind + ':' + run.cellIdx;
      const extra = overhang(run, bounds);
      overflow.set(id, Math.max(overflow.get(id) ?? 0, extra));
    }
  }
  if (glyphs.size) return { literal: JSON.stringify([...glyphs].sort((a,b) => a[0]-b[0])), overflow };
}

/** Native HTML paste can inflate short merged grid rows to its default
 * minimum. Recover their measured cell sizes only after the table, its
 * fonts, and its actual section geometry are present in the native writer. */
export function planHwpGridRows(html: string) {
  const plans = new Map<string, Plan | null>();
  if (html.includes('data-hwp-grid-row="true"')) for (const table of elementRanges(html,'table')) {
    const proof = sourcePlan(table.inner);
    if (!proof) continue;
    const signature = key(proof.rows,proof.plan.cells);
    plans.set(signature, plans.has(signature) ? null : proof.plan);
  }

  return (Core: typeof HwpDocument, doc: HwpDocument): HwpDocument => {
    if (!plans.size) return doc;
    let alternate: HwpDocument | undefined;
    try {
      let size = 0;
      const files = unzipSync(doc.exportHwpx(), { filter: entry => {
        if ((size += entry.originalSize) > documentLimits.documentInflatedBytes) throw Error('Native table proof too large');
        return true;
      } });
      const matches = new Map<string,{ address: TableAddress; file: string; start: number; end: number; xml: string }[]>();
      let global = 0;
      for (let section = 0; section < doc.getSectionCount(); section++) {
        const file = `Contents/section${section}.xml`, source = strFromU8(files[file]);
        const paragraphs = elementRanges(source, 'hp:p');
        if (paragraphs.length !== doc.getParagraphCount(section)) return doc;
        for (const [paragraph, owner] of paragraphs.entries()) {
          for (const table of elementRanges(owner.inner,'hp:tbl')) {
            if (table.inner.includes('<hp:tbl')) continue;
            const native = nativeCells(table.inner), signature = key(native.rows.length,native.cells);
            const expected = plans.get(signature);
            if (!expected) continue;
            let next = table.inner;
            for (const [i, cell] of [...native.ranges.entries()].reverse()) {
              const wanted = expected.heights[i];
              if (wanted === undefined) continue;
              const part = cell.inner.replace(/<hp:cellSz\b[^>]*\/>/, tag => {
                const stored = Number(attribute(tag,'height'));
                return stored > wanted + 3 ? tag.replace(/\bheight="\d+"/, `height="${wanted}"`) : tag;
              });
              next = next.slice(0,cell.start) + cell.open + part + '</hp:tc>' + next.slice(cell.end);
            }
            const start = owner.start + owner.open.length + table.start;
            const found = matches.get(signature) ?? [];
            found.push({ address: { section, paragraph, global: global + paragraph, cells: native.cells.length },
              file, start, end: start + table.end - table.start, xml: table.open + next + '</hp:tbl>' });
            matches.set(signature, found);
          }
        }
        global += paragraphs.length;
      }
      const affected = [];
      const unique = [...matches.values()].filter(entries => entries.length === 1).map(([entry]) => entry)
        .sort((a,b) => b.start - a.start);
      const edited = new Map<string,string>();
      const controls: Control[] = JSON.parse(doc.getControls());
      for (const entry of unique) {
        const original = edited.get(entry.file) ?? strFromU8(files[entry.file]);
        if (original.slice(entry.start,entry.end) === entry.xml) continue;
        const before = paintedTable(doc,controls,entry.address);
        if (!before) continue;
        affected.push({ address: entry.address, before });
        edited.set(entry.file,original.slice(0,entry.start) + entry.xml + original.slice(entry.end));
      }
      if (!affected.length) return doc;
      for (const [file,xml] of edited) files[file] = strToU8(xml);
      alternate = new Core(zipSync(files,{ level: 1 }));
      const trialControls: Control[] = JSON.parse(alternate.getControls());
      for (const entry of affected) {
        const after = paintedTable(alternate,trialControls,entry.address);
        if (!after || after.literal !== entry.before.literal ||
            [...after.overflow].some(([id, overhang]) => overhang > (entry.before.overflow.get(id) ?? 0) + .65)) return doc;
      }
      const output = alternate;
      alternate = undefined;
      return output;
    } catch {
      // Precise geometry is optional. A failed proof keeps the native file
      // made with its normal parsed cells and all of its original contents.
      return doc;
    } finally { alternate?.free(); }
  };
}
