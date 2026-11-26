import type { HwpDocument } from '@rhwp/core';
import type { CellPath, SerializedTablePolicy } from './cell-content';
import { hwpTableGrid } from './table-grid';
import { cellParagraphCss, cellParagraphKey, physicalLineHeight, paragraphAutoSpace } from './paragraph-layout';
import { hwpCellBorders } from './cell-borders';
import { characterScaleProperty } from '../character-scale';
import { hwpSourceWrapAttribute } from '../hwp-text-advances';
import { sourceNestedHangingScale } from './source-nested-hanging';
const escapeAttribute = (value: string) => value.replace(/&/g, '&amp;').replace(/"/g, '&quot;');

type SourceBox = { pageIndex: number; x: number; y: number; w: number; h: number };

/** Require two authored inline controls and matching painted bounds in the
 * same source cell and page before assigning cell-relative positions. */
function inlinePairPosition(doc: HwpDocument, section: number, paragraph: number, path: CellPath, metrics: {
  stored?: boolean; controlCount: number; lines: unknown[]; tables: { index: number; inline?: boolean }[];
}) {
  if (path.length !== 2 || !metrics.stored || metrics.controlCount !== 2 || metrics.lines.length !== 1 ||
    metrics.tables.length !== 2 || metrics.tables.some(t => !t.inline)) return;
  const [owner, nested] = path;
  const parent = JSON.stringify([owner]);
  if (doc.getCellParagraphCountByPath(section, paragraph, parent) !== 1 ||
    /[^ ]/.test(doc.getTextInCellByPath(section, paragraph, parent, 0, 0xffffffff))) return;
  const owners = (JSON.parse(doc.getTableCellBboxes(section, paragraph, owner.controlIndex)) as (SourceBox & { cellIdx: number; rowSpan: number })[])
    .filter(b => b.cellIdx === owner.cellIndex);
  if (owners.length !== 1 || owners[0].rowSpan !== 1) return;
  const home = owners[0], bounds: { left: number; right: number; top: number }[] = [];
  for (const control of metrics.tables) {
    const address = JSON.stringify([owner, { controlIndex: control.index, cellIndex: 0, cellParaIndex: 0 }]);
    const dims = JSON.parse(doc.getTableDimensionsByPath(section, paragraph, address));
    if (dims.rowCount !== 1 || dims.colCount !== dims.cellCount || dims.colCount < 1 || dims.colCount > 16) return;
    const cells = JSON.parse(doc.getTableCellBboxesByPath(section, paragraph, address)) as SourceBox[];
    if (cells.length !== dims.cellCount || cells.some(b => b.pageIndex !== home.pageIndex || !(b.w > 0 && b.h > 0) ||
      b.x < home.x || b.y < home.y || b.x + b.w > home.x + home.w + .2 || b.y + b.h > home.y + home.h + .2 ||
      Math.abs(b.y - cells[0].y) > .2)) return;
    bounds.push({ left: Math.min(...cells.map(b => b.x)), right: Math.max(...cells.map(b => b.x + b.w)), top: cells[0].y });
  }
  const [first, second] = bounds;
  const slot = metrics.tables.findIndex(t => t.index === nested.controlIndex);
  if (slot < 0 || first.right + .5 >= second.left || Math.abs(first.top - second.top) > .2) return;
  const x = Math.round((bounds[slot].left - home.x) * 15);
  const y = Math.round(((first.top + second.top) / 2 - home.y) * 15);
  return x >= 0 && x <= 32767 && y >= 0 && y <= 1440 ? `${x} ${y}` : undefined;
}

/** The clipboard HTML carries nested text and padding but omits its grid.
 * Restore source widths and minimum row heights without clipping content or scaling text. */
export function nestedTableGeometry(doc: HwpDocument, section: number, paragraph: number, path: CellPath, html: string, pitches: Map<string, number>, policy?: SerializedTablePolicy): string {
  const json = JSON.stringify(path), dims = JSON.parse(doc.getTableDimensionsByPath(section, paragraph, json));
  if (dims.cellCount > 50000 || dims.colCount > 1000 || dims.rowCount > 10000) throw Error('편집할 표의 셀 수가 너무 많습니다.');
  const cells = Array.from({ length: dims.cellCount }, (_, index) => ({
    index,
    ...JSON.parse(doc.getCellInfoByPath(section, paragraph, JSON.stringify([...path.slice(0, -1), { ...path.at(-1)!, cellIndex: index }]))),
    ...JSON.parse(doc.getCellPropertiesByPath(section, paragraph, json, index)),
  })).sort((a, b) => a.row - b.row || a.col - b.col);
  const columns = hwpTableGrid(cells, 'col', dims.colCount);
  const rows = hwpTableGrid(cells, 'row', dims.rowCount);
  const parent = JSON.stringify(path.slice(0, -1));
  const metrics = JSON.parse(doc.getParagraphLineMetrics(section, paragraph, parent));
  const table = metrics.tables.find((t: { index: number }) => t.index === path.at(-1)!.controlIndex);
  const cellPosition = inlinePairPosition(doc, section, paragraph, path, metrics);
  let spacing = Math.max(0, Number(table?.outerBottom) || 0);
  // A control-only inline table replaces its host line. Preserve that line's
  // trailing leading, rather than adding the editor's default 18px table gap.
  if (table?.inline && metrics.controlCount === 1 && !doc.getTextInCellByPath(section, paragraph, parent, 0, 0xffffffff).trim()) {
    const properties = JSON.parse(doc.getCellParaPropertiesAtByPath(section, paragraph, parent));
    const char = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, parent, 0));
    const leading = metrics.lines[0]?.spacing ??
      (['Percent', 'SpaceOnly'].includes(properties.lineSpacingType) ? physicalLineHeight(properties, char) - Number(char.fontSize) / 75 : 0);
    if (Number.isFinite(leading)) spacing += Math.max(0, leading);
  }
  const declaredCellBreak = policy?.pageBreak === 'CELL' && policy.treatAsChar === false;
  const splittableRows = new Set<number>();
  if (declaredCellBreak) for (const c of cells)
    if (c.rowSpan === 1 && (!c.isHeader || !policy.repeatHeader)) splittableRows.add(c.row);
  let physicalBoxes: { cellIdx: number; pageIndex: number; y: number; h: number; w: number }[] | undefined;
  let depth = 0, cell = 0, cellParagraph = 0, row = 0;
  const style = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old};${css}"`) : tag.replace(/>$/, ` style="${css}">`);
  return html.replace(/<\/?(?:table|tr|td|p)\b[^>]*>/g, (tag, offset: number) => {
    if (/^<table\b/.test(tag)) {
      depth++;
      if (depth !== 1) return tag;
      if (cellPosition) tag = tag.replace(/>$/, ` data-cell-table-position="${cellPosition}">`);
      return style(tag, `width:${columns.reduce((sum, n) => sum + n, 0)}px;margin-bottom:${Math.round(spacing * 100) / 100}px;`);
    }
    if (tag === '</table>') { depth--; return tag; }
    if (depth !== 1) return tag;
    // CSS table-row height is a minimum; text and nested content may grow it.
    if (/^<tr\b/.test(tag)) {
      if (splittableRows.has(row)) tag = tag.replace(/>$/, ' data-row-split="allow">');
      return style(tag, `height:${rows[row++] ?? 1}px;`);
    }
    if (/^<p\b/.test(tag)) {
      const c = cells[cell - 1]; if (!c) return tag;
      const sourceIndex = tag.match(/\bdata-hwp-cell-paragraph="(\d+)"/)?.[1];
      if (sourceIndex !== undefined) cellParagraph = Number(sourceIndex);
      const current = [...path.slice(0, -1), { ...path.at(-1)!, cellIndex: c.index, cellParaIndex: cellParagraph }];
      const currentJson = JSON.stringify(current), key = cellParagraphKey(section, paragraph, current);
      const properties = JSON.parse(doc.getCellParaPropertiesAtByPath(section, paragraph, currentJson));
      const char = JSON.parse(doc.getCellCharPropertiesAtByPath(section, paragraph, currentJson, 0));
      const count = doc.getCellParagraphCountByPath(section, paragraph, currentJson);
      const text = html.slice(offset + tag.length, html.indexOf('</p>', offset));
      const final = ++cellParagraph === count, simple = !/<(?:img|table)\b|<!--\s*rhwp:/i.test(text);
      if (path.length === 2 && c.rowSpan === 1 && c.colSpan === 1 && !c.isHeader && !c.textDirection &&
          !c.diagonalSlash && !c.diagonalBackSlash && simple && !table?.inline && !cellPosition &&
          !properties.marginLeft && !properties.marginRight && properties.indent < 0 &&
          text.includes(characterScaleProperty) && !tag.includes(hwpSourceWrapAttribute)) {
        try {
          const boxes = physicalBoxes ??= JSON.parse(doc.getTableCellBboxesByPath(section, paragraph, json)) as NonNullable<typeof physicalBoxes>;
          const witness = sourceNestedHangingScale(doc, section, paragraph, current, boxes, pitches, -properties.indent);
          if (witness) tag = tag.replace(/>$/, ` ${hwpSourceWrapAttribute}="${escapeAttribute(JSON.stringify(witness))}">`);
        } catch { /* Missing nested geometry retains the original model. */ }
      }
      // An empty or whitespace-only line has no painted run. Its native line
      // still exists and needs the same terminal budget as an ordinary cell;
      // absence from the paint index does not mean its layout is unknown.
      const lines = pitches.get(`lines:${key}`) ?? (final && simple
        ? JSON.parse(doc.getParagraphLineMetrics(section, paragraph, currentJson)).lines.length : undefined);
      const css = cellParagraphCss(properties, char, { pitch: pitches.get(key), inkHeight: pitches.get(`height:${key}`), final, lines, simple });
      const tableSpaces = tag.includes('data-hwp-inline-table-spaces="true"');
      return paragraphAutoSpace(style(tag, css + (tableSpaces ? 'line-height:0.1pt;' : '')), properties);
    }
    if (!/^<td\b/.test(tag)) return tag;
    cellParagraph = 0;
    const c = cells[cell++]; if (!c) return tag;
    if (declaredCellBreak && c.rowSpan > 1 && (!c.isHeader || !policy.repeatHeader)) tag = tag.replace(/>$/, ' data-cell-split="allow">');
    const widths = columns.slice(c.col, c.col + c.colSpan).map(n => Math.max(1, Math.round(n)));
    return style(tag.replace(/>$/, ` colwidth="${widths.join(',')}">`), `width:${Math.round(c.width / 75 * 100) / 100}px;vertical-align:${['top', 'middle', 'bottom'][c.verticalAlign] ?? 'top'};${hwpCellBorders(c)}`);
  });
}
