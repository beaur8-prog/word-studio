import { layoutHwpCellLines } from './hwp-cell-line-metrics';

const pixels = /^\d+(?:\.\d+)?px$/;
const variableContents = 'table,img,svg,[data-docx-tab],[data-word-overlap],[data-note-reference],' +
  '[data-document-field],[data-page-number],[data-column-count],[data-document-textbox],[data-flow-widget]';

function independentlySized(table: HTMLTableElement, body: HTMLElement) {
  if (table.parentElement !== body || !pixels.test(table.style.width) || table.style.zoom ||
      table.querySelector(variableContents) || getComputedStyle(table).tableLayout !== 'fixed') return false;
  const columns = [...table.querySelectorAll<HTMLTableColElement>(':scope > colgroup > col')];
  return columns.length > 0 && columns.every(column => pixels.test(column.style.width));
}

/** Repeated fit checks append blocks to the same page. An unchanged fixed
 * grid keeps its text width, so its inline terminal struts need no new pass. */
export function nativeFixedCellMeasurement() {
  let body: HTMLElement | undefined;
  let stable = new WeakMap<HTMLTableElement, boolean>();
  const invalidate = (entries: readonly MutationRecord[]) => {
    for (const { target } of entries) {
      const element = target instanceof Element ? target : target.parentElement;
      const table = element?.closest('table');
      if (table && table.parentElement === body) stable.delete(table);
    }
  };
  const observer = new MutationObserver(invalidate);
  let fontVersion = 0, measuredVersion = fontVersion;
  const fontChanged = () => { fontVersion++; };
  document.fonts.addEventListener('loadingdone', fontChanged);
  document.fonts.addEventListener('loadingerror', fontChanged);
  let dpr = window.devicePixelRatio;
  const prepare = (content: HTMLElement) => {
    if (body !== content) {
      observer.disconnect();
      body = content;
      observer.observe(body, { attributes: true, childList: true, characterData: true, subtree: true });
      stable = new WeakMap();
    }
    invalidate(observer.takeRecords());
    if (measuredVersion !== fontVersion || dpr !== window.devicePixelRatio || document.fonts.status !== 'loaded') {
      stable = new WeakMap();
      measuredVersion = fontVersion;
      dpr = window.devicePixelRatio;
    }
  };
  return {
    measure(content: HTMLElement) {
      prepare(content);
      const checked = new Map<HTMLTableElement, boolean>();
      layoutHwpCellLines(content, cell => {
        const table = cell.closest('table');
        if (!table || table.parentElement !== content) return false;
        if (stable.get(table) === !table.previousElementSibling) return true;
        if (!checked.has(table)) checked.set(table, independentlySized(table, content));
        return false;
      });
      // The pending records produced by layoutHwpCellLines are our own paint
      // changes, and must not invalidate its freshly measured result.
      observer.takeRecords();
      for (const [table, fixed] of checked) {
        if (fixed) stable.set(table, !table.previousElementSibling);
        else stable.delete(table);
      }
    },
    dispose() {
      observer.disconnect();
      document.fonts.removeEventListener('loadingdone', fontChanged);
      document.fonts.removeEventListener('loadingerror', fontChanged);
    },
  };
}
