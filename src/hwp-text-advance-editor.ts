import { Extension } from '@tiptap/core';
import type { Node as ModelNode } from '@tiptap/pm/model';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { columnResizingPluginKey } from '@tiptap/pm/tables';
import { Mapping } from '@tiptap/pm/transform';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import { documentScale } from './document-scale';
import { editorLayoutResume, editorLayoutSuspended } from './editor-layout-task';
import { editorHwpSingleCellLine, hwpEditorSingleLine } from './hwp-source-cell-lines';
import { parseHwpSourceWrap } from './hwp-text-advances';
import { hasUnchangedSourceWrap, hwpSourceAdvanceParagraph, hwpSourceAdvanceRun, mixedHwpCellRunParagraph, planHwpSourceAdvanceDecorations, shortHwpCellInkFits, shortHwpCellWidthParagraph, sourceCellWidthAdjustment, sourceFlowAllowance, type HwpCellWidthAdjustment, type HwpSourceAdvancePlan } from './hwp-text-advance-layout';

const key = new PluginKey<DecorationSet>('hwpSourceAdvances');
const measuredStyles = ['width','boxSizing','paddingLeft','paddingRight','paddingTop','paddingBottom','borderLeftWidth','borderRightWidth',
  'fontFamily','fontSize','fontWeight','fontStyle','lineHeight','minHeight','letterSpacing','wordSpacing',
  'whiteSpace','overflowWrap','wordBreak','lineBreak','textAlign','textIndent','direction','writingMode'] as const;

const sourceFlowDecoration = 'data-hwp-source-flow-base-height';
const sourceWidthDecoration = 'data-hwp-source-width-delta';
type CachedAdvance = { signature: string; width: number; plan: HwpSourceAdvancePlan };
type EditorAdvanceProbe = {
  paragraph: HTMLElement; node: ModelNode; start: number; width: number; height: number;
  marginRight: number; signature: string; lang: string;
  styles: readonly (readonly [typeof measuredStyles[number], string])[];
  mixed: boolean; allowance?: number; widthAdjustment?: HwpCellWidthAdjustment;
};
const probesPerBatch = 64;

async function editorAdvances(view: EditorView, cache: WeakMap<ModelNode, CachedAdvance>, isCurrent: () => boolean) {
  const host = document.createElement('div');
  host.className = 'tiptap';
  host.style.cssText = 'position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none';
  host.setAttribute('aria-hidden', 'true');
  const probes: EditorAdvanceProbe[] = [];
  const decorations: Decoration[] = [];
  const stack = view.dom.closest<HTMLElement>('.document-stack');
  let papers: DOMRect[] | undefined, scale: number | undefined;
  const append = (node: ModelNode, start: number, plan: HwpSourceAdvancePlan, marginRight: number) => {
    if (plan.baseHeight || plan.widthDelta) {
      const attributes: Record<string, string> = {};
      if (plan.baseHeight) attributes[sourceFlowDecoration] = String(plan.baseHeight);
      if (plan.widthDelta) {
        attributes[sourceWidthDecoration] = String(plan.widthDelta);
        attributes.style = `margin-right:${marginRight - plan.widthDelta}px`;
      }
      decorations.push(Decoration.node(start - 1, start - 1 + node.nodeSize, attributes,
        { style: JSON.stringify(attributes) }));
    }
    for (const { from, to, style } of plan.runs)
      decorations.push(Decoration.inline(start + from, start + to,
        { nodeName: 'span', [hwpSourceAdvanceRun]: 'true', style }, { style }));
  };
  view.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph' || !node.attrs.hwpSourceWrap && node.attrs.hwpSourceSpaces !== true ||
        node.content.content.some(child => !child.isText)) return;
    const actual = view.nodeDOM(pos);
    if (!(actual instanceof HTMLElement)) return;
    if (actual.dataset.hwpSourceWrap?.includes('"singleCellLine"')) {
      const indent = editorHwpSingleCellLine(actual, scale ??= documentScale(stack ?? view.dom));
      if (indent !== undefined) {
        const style = `--hwp-editor-single-line-indent:${indent}px`;
        decorations.push(Decoration.node(pos, pos + node.nodeSize, { [hwpEditorSingleLine]: 'true', style }, { style }));
        return;
      }
    }
    // A terminal prefix blocks a fresh width trial. Its compact line boxes
    // can coexist with an already accepted advance on the unchanged model.
    const approved = actual.dataset.hwpCellTerminal === 'true' && actual.hasAttribute(sourceFlowDecoration)
      ? cache.get(node) : undefined;
    const cell = actual.parentElement;
    if (approved?.plan.baseHeight && !approved.plan.widthDelta && approved.plan.runs.length &&
        cell instanceof HTMLTableCellElement && cell.tagName === 'TD' && cell.rowSpan === 1 && cell.colSpan === 1) {
      const css = getComputedStyle(actual), bounds = cell.getBoundingClientRect();
      const range = document.createRange(); range.selectNodeContents(actual);
      if (Math.abs(parseFloat(css.width) - approved.width) <= 1 / 64 && hasUnchangedSourceWrap(actual) &&
          [...range.getClientRects()].every(box => !(box.width > 0 && box.height > 0) ||
            box.left >= bounds.left - .5 && box.right <= bounds.right + .5 &&
            box.top >= bounds.top - .5 && box.bottom <= bounds.bottom + .5)) {
        append(node, pos + 1, approved.plan, parseFloat(css.marginRight) || 0);
        return;
      }
    }
    const short = shortHwpCellWidthParagraph(actual);
    const mixed = actual.hasAttribute('data-hwp-source-wrap') &&
      mixedHwpCellRunParagraph(actual, scale ??= documentScale(stack ?? view.dom));
    if (!short && !mixed && !hwpSourceAdvanceParagraph(actual)) return;
    const css = getComputedStyle(actual);
    const activeWidth = Number(actual.getAttribute(sourceWidthDecoration)) || 0;
    const marginRight = (parseFloat(css.marginRight) || 0) + activeWidth;
    const width = parseFloat(css.width) - activeWidth;
    const height = Number(actual.getAttribute(sourceFlowDecoration)) || parseFloat(css.height);
    if (!(width > 0 && height > 0)) return;
    const lang = actual.closest<HTMLElement>('[lang]')?.lang || '';
    let allowance: number | undefined, widthAdjustment: HwpCellWidthAdjustment | undefined;
    if (actual.hasAttribute('data-hwp-source-wrap')) {
      scale ??= documentScale(stack ?? view.dom);
      if (actual.parentElement instanceof HTMLTableCellElement) widthAdjustment = sourceCellWidthAdjustment(actual, scale, activeWidth);
      let paper;
      if (stack && !actual.closest('td,th')) {
        papers ??= [...stack.querySelectorAll('.section-paper,.paper-backgrounds .paper')]
          .map(p => p.getBoundingClientRect()).filter(box => box.height > 0).sort((a,b) => a.top - b.top);
        const y = actual.getBoundingClientRect().top;
        const at = papers.findIndex(box => y >= box.top - .05 && y < box.bottom);
        if (at >= 0) {
          const context = actual.closest<HTMLElement>('[data-bottom-margin]');
          const margin = Number(context?.dataset.bottomMargin);
          if (Number.isFinite(margin)) paper = { index: at + 1, box: papers[at], bottom: papers[at].bottom - margin * scale,
            scale, liveEditor: true, topMargin: Number(context?.dataset.margin) };
        }
      }
      allowance = sourceFlowAllowance(actual, paper);
    }
    // A restored one-line decoration already owns the proved width. Its
    // current run strut may differ from the unadjusted clone during undo.
    // Keep the live result while the current cell and painted line still fit.
    if (short && activeWidth && widthAdjustment && Math.abs(activeWidth - widthAdjustment.delta) <= .05 &&
        shortHwpCellInkFits(actual)) {
      append(node, pos + 1, { runs: [], baseHeight: height, widthDelta: activeWidth }, marginRight);
      return;
    }
    const signature = [lang, mixed, height, width, widthAdjustment?.delta ?? 0, allowance || 0,
      ...measuredStyles.slice(1).map(property => css[property])].join('|');
    const previous = cache.get(node);
    if (previous?.signature === signature) { append(node, pos + 1, previous.plan, marginRight); return; }
    probes.push({ paragraph: actual, node, start: pos + 1, width, height, marginRight, signature, lang,
      styles: measuredStyles.map(property => [property, css[property]] as const), mixed, allowance, widthAdjustment });
  });
  // The editor probes are independent, absolutely positioned paragraphs. They
  // can be measured in bounded batches without changing their context or the
  // order in which the finished decorations are applied.
  for (let offset = 0; offset < probes.length; offset += probesPerBatch) {
    if (offset > 0) await new Promise<void>(resolve => setTimeout(resolve, 0));
    if (!isCurrent()) return;
    const addresses = new Map<HTMLElement, EditorAdvanceProbe>();
    const allowances = new Map<HTMLElement, number>();
    const widths = new Map<HTMLElement, HwpCellWidthAdjustment>();
    const mixedRuns = new Set<HTMLElement>();
    for (const measurement of probes.slice(offset, offset + probesPerBatch)) {
      const { paragraph, node, width, styles, lang, mixed, allowance, widthAdjustment } = measurement;
      const probe = paragraph.cloneNode(true) as HTMLElement;
      for (const span of probe.querySelectorAll('[' + hwpSourceAdvanceRun + ']')) span.replaceWith(...span.childNodes);
      probe.normalize();
      if (node.textContent !== probe.textContent) continue;
      for (const [property, value] of styles) probe.style[property] = value;
      probe.style.width = width + 'px';
      probe.style.maxWidth = 'none';
      probe.style.position = 'absolute';
      probe.style.top = '0';
      probe.lang = lang;
      addresses.set(probe, measurement);
      if (mixed) mixedRuns.add(probe);
      if (allowance) allowances.set(probe, allowance);
      if (widthAdjustment) widths.set(probe, widthAdjustment);
      host.append(probe);
    }
    if (!addresses.size) continue;
    document.body.append(host);
    try {
      for (const [probe, { node, width, height, signature }] of addresses) {
        const css = getComputedStyle(probe);
        if ((!mixedRuns.has(probe) && Math.abs(parseFloat(css.height) - height) > 1 / 64) ||
            Math.abs(parseFloat(css.width) - width) > 1 / 64) {
          cache.set(node, { signature, width, plan: { runs: [] } });
          allowances.delete(probe);
          widths.delete(probe);
          mixedRuns.delete(probe);
          addresses.delete(probe);
          probe.remove();
        }
      }
      const approved = await planHwpSourceAdvanceDecorations(host, allowances, widths, mixedRuns);
      if (!isCurrent()) return;
      for (const [paragraph, { node, start, width, signature, marginRight }] of addresses) {
        const plan = paragraph.textContent === node.textContent ? approved.get(paragraph) : undefined;
        const value = plan ?? { runs: [] };
        cache.set(node, { signature, width, plan: value });
        append(node, start, value, marginRight);
      }
    } finally { host.remove(); host.replaceChildren(); }
  }
  return decorations.sort((a, b) => a.from - b.from || a.to - b.to);
}

function captionTextAtPointer(view: EditorView, event: MouseEvent): boolean {
  if (!view.editable || !(event.target instanceof HTMLElement)) return false;
  const paragraph = event.target.closest(`p[${hwpEditorSingleLine}]`);
  if (!paragraph) return false;
  const resize = columnResizingPluginKey.getState(view.state);
  if (!resize || resize.dragging) return false;
  const range = document.createRange();
  range.selectNodeContents(paragraph);
  const overText = [...range.getClientRects()].some(box => box.width > 0 && box.height > 0 &&
    event.clientX >= box.left && event.clientX < box.right && event.clientY >= box.top && event.clientY < box.bottom);
  if (overText && resize.activeHandle >= 0)
    view.dispatch(view.state.tr.setMeta(columnResizingPluginKey, { setHandle: -1 }));
  return overText;
}

export const HwpSourceAdvanceEditor = Extension.create({
  name: 'hwpSourceAdvanceEditor',
  priority: 101,
  addProseMirrorPlugins: () => [new Plugin({
    key,
    state: { init: () => DecorationSet.empty, apply: (tr, before) => tr.getMeta(key) ?? before.map(tr.mapping, tr.doc) },
    props: {
      decorations: state => key.getState(state),
      handleDOMEvents: {
        // Only the text itself takes precedence over the column handle.
        // The cell border and any resize that has begun keep working.
        mousemove: captionTextAtPointer,
        mousedown(view, event) { captionTextAtPointer(view, event); return false; },
      },
    },
    appendTransaction(transactions, previous, current) {
      if (!transactions.some(transaction => transaction.docChanged)) return null;
      const from = previous.doc.content.findDiffStart(current.doc.content);
      if (from === null || from >= current.doc.content.size) return null;
      const end = previous.doc.content.findDiffEnd(current.doc.content);
      const to = Math.max(from + 1, end?.b ?? current.doc.content.size);
      const mapping = new Mapping();
      for (const transaction of transactions) mapping.appendMapping(transaction.mapping);
      const oldPositions = mapping.invert();
      const change = current.tr;
      current.doc.nodesBetween(from, to, (paragraph, position) => {
        if (paragraph.type.name !== 'paragraph') return;
        const source = paragraph.attrs.hwpSourceWrap;
        const reference = parseHwpSourceWrap(source);
        if (!reference || !reference.cellWidth && !reference.bodyWidth && !reference.origin &&
            reference.columnWidth === undefined && !reference.authoredBodyLines) return false;
        const original = previous.doc.nodeAt(oldPositions.map(position, -1));
        if (original?.type === paragraph.type && original.attrs.hwpSourceWrap === source &&
          original.textContent === paragraph.textContent && !original.eq(paragraph)) {
          // An unchanged string can reflow after formatting. The same history
          // event retires the saved visual lines and restores them on undo.
          change.setNodeMarkup(position, undefined, { ...paragraph.attrs, hwpSourceWrap: null });
        }
        return false;
      });
      return change.docChanged ? change.setStoredMarks(current.storedMarks) : null;
    },
    view(view) {
      let frame = 0, pending = false, requested = false, destroyed = false;
      let cache = new WeakMap<ModelNode, CachedAdvance>();
      const update = () => {
        if (destroyed || frame) return;
        if (pending) { requested = true; return; }
        frame = requestAnimationFrame(async () => {
          frame = 0;
          if (destroyed || view.composing || editorLayoutSuspended() || !view.dom.isConnected) return;
          if (!view.dom.querySelector('[data-hwp-source-spaces="true"],[data-hwp-source-wrap]') && key.getState(view.state) === DecorationSet.empty) return;
          const current = view.state.doc, measurements = cache;
          const isCurrent = () => !destroyed && !view.composing && !editorLayoutSuspended() &&
            current === view.state.doc && measurements === cache;
          pending = true;
          try {
            const decorations = await editorAdvances(view, measurements, isCurrent);
            if (!decorations || !isCurrent()) return;
            const existing = key.getState(view.state)?.find() || [];
            const same = existing.length === decorations.length && decorations.every((item, index) => {
              const previous = existing[index];
              return item.from === previous.from && item.to === previous.to && item.spec.style === previous.spec.style;
            });
            if (!same)
              view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(current, decorations)).setMeta('addToHistory', false));
          } finally {
            pending = false;
            const retry = requested || current !== view.state.doc || measurements !== cache;
            requested = false;
            if (retry) update();
          }
        });
      };
      const fontsChanged = () => { cache = new WeakMap<ModelNode, CachedAdvance>(); update(); };
      const resize = new ResizeObserver(update);
      resize.observe(view.dom);
      view.dom.addEventListener('compositionend', update);
      document.fonts.addEventListener('loadingdone', fontsChanged);
      document.addEventListener(editorLayoutResume, update);
      update();
      return {
        update(next, previous) { if (next.state.doc !== previous.doc) update(); },
        destroy() {
          destroyed = true;
          cancelAnimationFrame(frame);
          resize.disconnect();
          view.dom.removeEventListener('compositionend', update);
          document.fonts.removeEventListener('loadingdone', fontsChanged);
          document.removeEventListener(editorLayoutResume, update);
        },
      };
    },
  })],
});
