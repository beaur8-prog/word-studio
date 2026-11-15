import { documentScale } from './document-scale';

const measured = new WeakMap<HTMLTableElement, { html: string; font: string; available: number; generation: number }>();
let generation = 0;
if (typeof document !== 'undefined') document.fonts?.addEventListener('loadingdone', () => generation++);

/** The imported grid is a preference under OOXML autofit. When every column's
 * unwrapped content fits the containing text area, let deficient columns grow
 * to that content instead of imposing another line. Keep the authored grid in
 * the model; these measured widths belong only to the view and print tree.
 * Nested tables retain their existing grid until parent/child constraints can
 * be resolved together. Larger/overflowing tables retain the existing bounded grid, rather than
 * inventing a new shrink or wrapping policy here. */
export function layoutWordAutoFitTables(root: HTMLElement, freeze = false) {
  for (const table of [...root.querySelectorAll<HTMLTableElement>('table[data-word-table-layout="autofit"][data-word-table-width-auto="true"]')].reverse()) {
    if (table.dataset.printAutoFitFrozen === 'true') continue;
    if (table.querySelector('table') || table.parentElement?.closest('table')) continue;
    if (table.dataset.tableAnchor || table.closest('.table-overlay-frame,.table-wrap-frame')) continue;
    const cols = [...table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')];
    const first = table.rows[0];
    if (!cols.length || !first) continue;
    // A non-auto tcW is an explicit cell preference, not just the table grid.
    // Retain it; this bounded content-fitting path handles unconstrained cells.
    if ([...table.rows].some(row => [...row.cells].some(cell => cell.dataset.wordCellWidthAuto !== 'true'))) continue;
    // Cell colwidth attributes are authored model values. The colgroup itself
    // may already hold this view's result from the previous layout pass.
    const preferred = [...first.cells].flatMap(cell => (cell.getAttribute('colwidth') || '').split(',').map(Number));
    if (preferred.length !== cols.length || preferred.some(n => !Number.isFinite(n) || n <= 0)) continue;
    const parent = table.parentElement;
    if (!parent) continue;
    const css = getComputedStyle(parent), scale = documentScale(parent);
    let available = parent.getBoundingClientRect().width / scale -
      (parseFloat(css.paddingLeft) || 0) - (parseFloat(css.paddingRight) || 0) -
      (parseFloat(css.borderLeftWidth) || 0) - (parseFloat(css.borderRightWidth) || 0);
    const columns = table.closest<HTMLElement>('[data-column-count]');
    if (columns && Number(columns.dataset.columnCount) > 1) {
      const count = Number(columns.dataset.columnCount), gap = Number(columns.dataset.columnSpace) || 0;
      available = Math.min(available, (columns.getBoundingClientRect().width / documentScale(columns) - gap * (count - 1)) / count);
    }
    if (!(available > 0)) continue;
    const font = getComputedStyle(table).font, cached = measured.get(table);
    if (cached?.html === table.innerHTML && cached.font === font && cached.available === available && cached.generation === generation) continue;
    const host = document.createElement('div'); host.className = 'tiptap'; host.setAttribute('aria-hidden', 'true');
    host.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-100000px;top:0;padding:0;min-height:0;width:max-content;font:${font}`;
    const copy = table.cloneNode(true) as HTMLTableElement;
    copy.querySelectorAll('.pagination-gap,.pagination-cell-gap,[data-flow-widget],.ProseMirror-separator').forEach(n => n.remove());
    copy.querySelectorAll('[id]').forEach(n => n.removeAttribute('id'));
    copy.querySelector(':scope > colgroup')?.remove();
    const group = document.createElement('colgroup');
    preferred.forEach(() => group.append(document.createElement('col'))); copy.prepend(group);
    Object.assign(copy.style, { tableLayout: 'auto', width: 'max-content', maxWidth: 'none', margin: '0', position: 'static', transform: 'none', float: 'none' });
    host.append(copy); document.body.append(host);
    let widths: number[];
    try {
      const natural = [...group.children].map(col => col.getBoundingClientRect().width);
      const extra = copy.getBoundingClientRect().width - natural.reduce((a, b) => a + b, 0);
      const proposed = preferred.map((width, i) => Math.max(width, natural[i]));
      widths = natural.every(n => Number.isFinite(n) && n > 0) && proposed.reduce((a, b) => a + b, 0) + extra <= available
        ? proposed : preferred;
    } finally { host.remove(); }
    cols.forEach((col, i) => { const width = `${widths[i]}px`; if (col.style.width !== width) col.style.width = width; });
    const width = `${widths.reduce((a, b) => a + b, 0)}px`;
    if (table.style.width !== width) table.style.width = width;
    if (freeze) table.dataset.printAutoFitFrozen = 'true';
    measured.set(table, { html: table.innerHTML, font, available, generation });
  }
}
