import type { HwpDocument } from '@rhwp/core';
import type { CellPath } from './cell-content';
import { hwpTableGrid } from './table-grid';

/** The clipboard HTML carries nested text and padding but omits its grid.
 * Restore real column widths without fixing row heights or scaling text. */
export function nestedTableGeometry(doc: HwpDocument, section: number, paragraph: number, path: CellPath, html: string): string {
  const json = JSON.stringify(path), dims = JSON.parse(doc.getTableDimensionsByPath(section, paragraph, json));
  if (dims.cellCount > 50000 || dims.colCount > 1000 || dims.rowCount > 10000) throw Error('편집할 표의 셀 수가 너무 많습니다.');
  const cells = Array.from({ length: dims.cellCount }, (_, index) => ({
    ...JSON.parse(doc.getCellInfoByPath(section, paragraph, JSON.stringify([...path.slice(0, -1), { ...path.at(-1)!, cellIndex: index }]))),
    ...JSON.parse(doc.getCellPropertiesByPath(section, paragraph, json, index)),
  })).sort((a, b) => a.row - b.row || a.col - b.col);
  const columns = hwpTableGrid(cells, 'col', dims.colCount);
  let depth = 0, cell = 0;
  const style = (tag: string, css: string) => tag.includes('style="') ? tag.replace(/style="([^"]*)"/, (_, old) => `style="${old};${css}"`) : tag.replace(/>$/, ` style="${css}">`);
  return html.replace(/<\/?(?:table|td)\b[^>]*>/g, tag => {
    if (/^<table\b/.test(tag)) {
      depth++;
      return depth === 1 ? style(tag, `width:${columns.reduce((sum, n) => sum + n, 0)}px;`) : tag;
    }
    if (tag === '</table>') { depth--; return tag; }
    if (depth !== 1 || !/^<td\b/.test(tag)) return tag;
    const c = cells[cell++]; if (!c) return tag;
    const widths = columns.slice(c.col, c.col + c.colSpan).map(n => Math.max(1, Math.round(n)));
    return style(tag.replace(/>$/, ` colwidth="${widths.join(',')}">`), `width:${Math.round(c.width / 75 * 100) / 100}px;vertical-align:${['top', 'middle', 'bottom'][c.verticalAlign] ?? 'top'};`);
  });
}
