import { W, all, child, val } from './docx-format';

/** Word permits omitted leading/trailing cells and legacy horizontal merges.
 * Materialize that grid before Mammoth computes vertical merges. Keep strict
 * validation for user JSON; a ragged DOCX table must not reject the whole file. */
export function normalizeDocxTables(body: Document): number {
  let repaired = 0;
  const children = (el: Element, name: string) => Array.from(el.children).filter(c => c.localName === name);
  const count = (el: Element | undefined, fallback = 0) => {
    const n = Number(val(el));
    if (!Number.isInteger(n) || n < 0 || n > 1000) throw new Error('표의 열 수가 올바르지 않습니다.');
    return n || fallback;
  };
  const make = (name: string, attrs: Record<string, string> = {}) => {
    const el = body.createElementNS(W, `w:${name}`);
    for (const [key, value] of Object.entries(attrs)) el.setAttributeNS(W, `w:${key}`, value);
    return el;
  };
  for (const table of all(body, 'tbl')) {
    const rows = children(table, 'tr');
    const span = (cell: Element) => count(child(child(cell, 'tcPr'), 'gridSpan'), 1);
    let changed = false;
    for (const row of rows) {
      let first: Element | undefined;
      for (const cell of children(row, 'tc')) {
        const props = child(cell, 'tcPr'), merge = child(props, 'hMerge');
        if (!merge) { first = undefined; continue; }
        if (val(merge) === 'restart' || !first) first = cell;
        else {
          const firstProps = child(first, 'tcPr')!;
          const total = span(first) + span(cell);
          if (total > 1000) throw new Error('표의 열 수가 너무 많습니다.');
          child(firstProps, 'gridSpan')?.remove();
          firstProps.append(make('gridSpan', { val: String(total) }));
          for (const block of Array.from(cell.children)) if (block !== props) first.append(block);
          cell.remove();
        }
        merge.remove(); changed = true;
      }
      const header = child(child(row, 'trPr'), 'tblHeader');
      if (header && ['0', 'false', 'off'].includes(val(header))) header.remove();
    }
    const widths = rows.map(row => children(row, 'tc').reduce((n, c) => n + span(c),
      count(child(child(row, 'trPr'), 'gridBefore')) + count(child(child(row, 'trPr'), 'gridAfter'))));
    const grid = child(table, 'tblGrid');
    const width = Math.max(grid?.children.length ?? 0, ...widths, 0);
    if (width > 1000) throw new Error('표의 열 수가 너무 많습니다.');
    const blank = (columns: number) => {
      const cell = make('tc'), props = make('tcPr'), borders = make('tcBorders');
      props.append(make('gridSpan', { val: String(columns) }));
      for (const side of ['top', 'right', 'bottom', 'left']) borders.append(make(side, { val: 'nil' }));
      props.append(borders); cell.append(props, make('p')); return cell;
    };
    for (const row of rows) {
      const props = child(row, 'trPr'), before = child(props, 'gridBefore'), after = child(props, 'gridAfter');
      const leading = count(before);
      if (leading) { row.insertBefore(blank(leading), children(row, 'tc')[0] ?? null); changed = true; }
      const occupied = children(row, 'tc').reduce((n, c) => n + span(c), 0);
      if (occupied < width) { row.append(blank(width - occupied)); changed = true; }
      before?.remove(); after?.remove();
    }
    if (changed) repaired++;
  }
  return repaired;
}
