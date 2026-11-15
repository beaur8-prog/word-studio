import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { columnResizingPluginKey } from '@tiptap/pm/tables';
import { editorLayoutResume, editorLayoutSuspended } from './editor-layout-task';

const probeFamily = 'Studio DotumChe spacing probe';
const acceptedFamilies = new Set(['돋움체', 'dotumche']);
const familyName = (family: string) => family.split(',')[0].replace(/["']/g, '').trim().toLowerCase();
const fontStyle = 'font-kerning:none;font-variant-ligatures:none';
const marker = 'data-word-cell-ink';
const runSelector = 'table[data-word-table-layout="fixed"] td > p[data-word-line-rule="exact"] > span[style*="letter-spacing"]';
const overlaySelector = 'img[data-image-floating],.document-textbox,table[data-table-anchor],.table-overlay-active,.table-wrap-frame,.pagination-footnote,.print-footnotes';
const unit = 64;
let knownFace: Promise<boolean> | undefined;
let canvas: CanvasRenderingContext2D | null = null;

function authoredFaceRegistered() {
  return [...document.fonts].some(face => acceptedFamilies.has(familyName(face.family)));
}

function candidates(root: HTMLElement, selector = runSelector) {
  return [...root.querySelectorAll<HTMLElement>(selector)].filter(run =>
    acceptedFamilies.has(familyName(run.style.fontFamily)) && parseFloat(run.style.letterSpacing) < 0
  );
}

function glyphStyle(spacing: number) {
  return `letter-spacing:${spacing}px;${fontStyle}`;
}

function fallbackAvailable() {
  if (!knownFace) knownFace = (async () => {
    const native = await Promise.all(['돋움체', 'DotumChe', 'DotumChe-Regular'].map(async name => {
      try {
        await new FontFace('Studio DotumChe native probe', `local(${JSON.stringify(name)})`).load();
        return true;
      } catch { return false; }
    }));
    if (native.some(Boolean)) return false;
    try {
      const probe = await new FontFace(probeFamily, 'local("AppleSDGothicNeo-Regular")').load();
      document.fonts.add(probe);
      return true;
    } catch { return false; }
  })();
  return knownFace;
}

function paragraphRun(paragraph: HTMLElement) {
  if (paragraph.tagName !== 'P' || paragraph.dataset.wordLineRule !== 'exact' ||
      paragraph.hasAttribute('data-hwp-cell-terminal') || paragraph.hasAttribute('data-hwp-cell-continuation') ||
      paragraph.querySelector('[data-hwp-cell-prefix],[data-flow-widget],[data-note-reference]')) return;
  const cell = paragraph.parentElement;
  if (!(cell instanceof HTMLTableCellElement) || cell.tagName !== 'TD' || cell.colSpan !== 1 || cell.rowSpan !== 1 ||
      [...cell.children].some(child => child !== paragraph && !child.classList.contains('column-resize-handle')) ||
      cell.dataset.wordCellWidthAuto === 'true' ||
      !(parseFloat(cell.getAttribute('colwidth') || cell.style.width) > 0)) return;
  const table = cell.closest('table');
  if (!table || table.dataset.wordTableLayout !== 'fixed' || table.dataset.wordTableWidthAuto === 'true' ||
      table.hasAttribute('data-table-anchor') || table.parentElement?.closest('td,th,.document-furniture,.document-textbox,[data-column-count]:not([data-column-count="1"])') ||
      table.style.zoom || cell.parentElement?.style.zoom) return;
  const run = paragraph.firstElementChild;
  if (run instanceof HTMLElement && run.tagName === 'SPAN' && !!run.style.letterSpacing) return run;
}

function overlapsFloatingObject(paragraph: HTMLElement, objects: Iterable<Element>) {
  const body = paragraph.getBoundingClientRect();
  for (const object of objects) {
    const box = object.getBoundingClientRect();
    if (Math.min(body.right, box.right) > Math.max(body.left, box.left) &&
        Math.min(body.bottom, box.bottom) > Math.max(body.top, box.top)) return true;
  }
  return false;
}

function textLine(run: HTMLElement) {
  const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT);
  const range = document.createRange();
  let line: { top: number; bottom: number } | undefined;
  while (walker.nextNode()) {
    range.selectNodeContents(walker.currentNode);
    for (const box of range.getClientRects()) {
      if (line && (box.top !== line.top || box.bottom !== line.bottom)) return;
      line = { top: box.top, bottom: box.bottom };
    }
  }
  return line;
}

function layoutChoice(paragraph: HTMLElement, run: HTMLElement, value: string) {
  const css = getComputedStyle(run), paragraphStyle = getComputedStyle(paragraph);
  const originalSpacing = parseFloat(css.letterSpacing);
  const insets = paragraphStyle.boxSizing === 'border-box' ?
    parseFloat(paragraphStyle.paddingLeft) + parseFloat(paragraphStyle.paddingRight) +
    parseFloat(paragraphStyle.borderLeftWidth) + parseFloat(paragraphStyle.borderRightWidth) : 0;
  const width = Math.round((parseFloat(paragraphStyle.width) - insets) * unit) / unit;
  if (!/^[!-~]{2,96}$/.test(value) || value !== run.textContent || !(originalSpacing < 0) || !(width > 0) ||
      !acceptedFamilies.has(familyName(css.fontFamily)) || css.fontStyle !== 'normal' || css.fontWeight !== '400' ||
      css.textTransform !== 'none' || css.direction !== 'ltr' || css.writingMode !== 'horizontal-tb' ||
      css.fontVariantCaps !== 'normal' || !['normal', '"liga" 0'].includes(css.fontFeatureSettings) ||
      parseFloat(paragraphStyle.textIndent) || !['left', 'start'].includes(paragraphStyle.textAlign)) return;
  if (!textLine(run)) return;
  canvas ??= document.createElement('canvas').getContext('2d');
  const context = canvas;
  if (!context) return;
  const sample = document.createElement('span');
  sample.lang = paragraph.closest<HTMLElement>('[lang]')?.lang ?? '';
  sample.style.cssText = `position:fixed;visibility:hidden;pointer-events:none;left:-10000px;white-space:pre;font-size:${css.fontSize};${fontStyle}`;
  document.body.append(sample);
  try {
    const shape = (family: string) => {
      sample.style.fontFamily = family;
      return ['ATW', 'il0', '@.', value].map(text => {
        sample.textContent = text;
        return sample.getBoundingClientRect().width;
      });
    };
    const actual = shape(css.fontFamily), known = shape(JSON.stringify(probeFamily));
    if (!actual.every((value, index) => value > 0 && value === known[index])) return;
    context.font = `${css.fontSize} ${JSON.stringify(probeFamily)}`;
    context.fontKerning = 'none';
    const glyphs = Array.from(value, char => context.measureText(char));
    const minima = glyphs.slice(0, -1).map((glyph, index) => Math.max(
      glyph.actualBoundingBoxRight + glyphs[index + 1].actualBoundingBoxLeft, 1 / unit
    ) - glyph.width);
    if (!minima.some(minimum => minimum > originalSpacing)) return;
    const adjustments = (gap: number) => minima.map(minimum => Math.ceil(Math.max(originalSpacing, minimum + gap) * unit) / unit);
    const base = glyphs.reduce((sum, glyph) => sum + glyph.width, 0);
    const fits = (gap: number) => base + adjustments(gap).reduce((sum, spacing) => sum + spacing, 0) <= width - 1 / unit;
    if (!fits(0)) return;
    let low = 0, high = unit;
    while (low < high) {
      const midpoint = Math.ceil((low + high) / 2);
      if (fits(midpoint / unit)) low = midpoint; else high = midpoint - 1;
    }
    const proposed = [...adjustments(low / unit), 0];
    sample.style.fontFamily = css.fontFamily;
    sample.replaceChildren(...Array.from(value, (char, index) => {
      const span = document.createElement('span');
      span.style.letterSpacing = `${proposed[index]}px`;
      span.textContent = char;
      return span;
    }));
    if (sample.getBoundingClientRect().width <= width) return proposed;
  } finally {
    sample.remove();
  }
}

/** Final print rows are already fitted to paper. Preserve every cell and
 * row boundary, and retain the original layout if the spans rewrap. */
export async function layoutWordCellSpacing(root: HTMLElement) {
  const runs = candidates(root, '.print-content ' + runSelector);
  if (!runs.length || authoredFaceRegistered() || !await fallbackAvailable() || authoredFaceRegistered()) return;
  for (const run of runs) {
    const paragraph = run.parentElement!;
    if (paragraphRun(paragraph) !== run || paragraph.childNodes.length !== 1 || run.childNodes.length !== 1 ||
        !(run.firstChild instanceof Text) || paragraph.matches('[data-print-paragraph-fragment],[data-print-soft-end]') ||
        paragraph.closest('[data-print-merge-continuation]')) continue;
    const paper = paragraph.closest('.print-page');
    if (!paper || overlapsFloatingObject(paragraph, paper.querySelectorAll(overlaySelector))) continue;
    const text = run.firstChild;
    const proposed = layoutChoice(paragraph, run, text.data);
    if (!proposed) continue;
    const cell = paragraph.parentElement!, row = cell.parentElement!, table = row.closest('table')!;
    const geometry = () => [paragraph, cell, row, table].map(element => {
      const box = element.getBoundingClientRect();
      return [box.x, box.y, box.width, box.height].join(',');
    }).join(';');
    const before = geometry(), line = textLine(run);
    // Exact-height lines can paint slightly outside the body at the top of
    // a page. Keep the existing vertical ink bounds on the physical sheet.
    const bounds = paper.getBoundingClientRect();
    if (!line || line.top < bounds.top || line.bottom > bounds.bottom) continue;
    const spans = Array.from(text.data, (char, index) => {
      const span = document.createElement('span');
      span.style.cssText = glyphStyle(proposed[index]);
      span.setAttribute(marker, 'true');
      span.textContent = char;
      return span;
    });
    run.replaceChildren(...spans);
    const placed = textLine(run);
    if (geometry() !== before || placed?.top !== line.top || placed.bottom !== line.bottom)
      run.replaceChildren(text);
  }
}

const key = new PluginKey<DecorationSet>('wordCellSpacing');
export const WordCellSpacing = Extension.create({
  name: 'wordCellSpacing',
  addProseMirrorPlugins: () => [new Plugin({
    key,
    state: { init: () => DecorationSet.empty, apply: (tr, prior) => tr.getMeta(key) ?? prior.map(tr.mapping, tr.doc) },
    props: { decorations: state => key.getState(state) },
    view(view) {
      let frame = 0, ready: boolean | undefined, checking = false, destroyed = false, signature = '';
      let doc = view.state.doc;
      const schedule = () => {
        if (destroyed || frame) return;
        frame = requestAnimationFrame(() => {
          frame = 0;
          if (!view.dom.isConnected || view.composing || editorLayoutSuspended() || view.dom.classList.contains('table-row-resizing')) return;
          const runs = candidates(view.dom);
          const original = runs.length > 0 && authoredFaceRegistered();
          if (ready === undefined && !checking && runs.length && !original) {
            checking = true;
            void fallbackAvailable().then(available => { ready = available; schedule(); });
          }
          const decos: Decoration[] = [], values: string[] = [];
          if (ready && runs.length && !original) {
            const resize = columnResizingPluginKey.getState(view.state);
            const resized = resize?.dragging ? view.nodeDOM(resize.activeHandle) : null;
            const resizingTable = resized instanceof HTMLElement ? resized.closest('table') : null;
            const objects = view.dom.querySelectorAll(overlaySelector);
            for (const run of runs) {
              const paragraph = run.parentElement!;
              if (paragraphRun(paragraph) !== run ||
                  resize?.dragging && (!resizingTable || paragraph.closest('table') === resizingTable) ||
                  overlapsFloatingObject(paragraph, objects)) continue;
              const place = view.state.doc.resolve(view.posAtDOM(paragraph, 0));
              const node = place.parent, content = node.firstChild;
              if (node.type.name !== 'paragraph' || node.childCount !== 1 || place.node(-1).childCount !== 1 || !content?.isText ||
                  content.marks.length !== 1 || content.marks[0].type.name !== 'textStyle') continue;
              const proposed = layoutChoice(paragraph, run, content.text!);
              if (!proposed) continue;
              const start = place.start();
              for (const [index, amount] of proposed.entries()) {
                const at = start + index;
                decos.push(Decoration.inline(at, at + 1, { nodeName: 'span', [marker]: 'true', style: glyphStyle(amount) }));
              }
              values.push(`${start}:${proposed.join(',')}`);
            }
          }
          const next = values.join(';');
          if (next === signature && (next === '' || doc === view.state.doc)) return;
          signature = next;
          doc = view.state.doc;
          view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(doc, decos)).setMeta('addToHistory', false));
        });
      };
      const resize = new ResizeObserver(schedule);
      resize.observe(view.dom);
      view.dom.addEventListener('compositionend', schedule);
      document.addEventListener(editorLayoutResume, schedule);
      document.fonts.addEventListener('loadingdone', schedule);
      schedule();
      return {
        update(next, prior) {
          if (next.state.doc !== prior.doc ||
              columnResizingPluginKey.getState(next.state)?.dragging !== columnResizingPluginKey.getState(prior)?.dragging)
            schedule();
        },
        destroy() {
          destroyed = true;
          cancelAnimationFrame(frame);
          resize.disconnect();
          view.dom.removeEventListener('compositionend', schedule);
          document.removeEventListener(editorLayoutResume, schedule);
          document.fonts.removeEventListener('loadingdone', schedule);
        },
      };
    },
  })],
});
