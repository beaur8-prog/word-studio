import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { HwpInfo, HwpPage } from './reader';
import documentLimits from '../../electron/document-limits.json';
import { HWP_VIEW_PADDING, hwpPageAtOffset, hwpViewportLayout, hwpVisiblePages } from './viewport-layout';

interface Paint { page?: HwpPage; error?: string }
interface Props {
  info: HwpInfo; zoom: number; navigation: { page: number; id: number }; paused: boolean;
  getPage(index: number): Promise<HwpPage>; keepPages(indices: number[]): void; onPageChange(index: number): void; onLoadingChange(value: boolean): void;
}

/** Exact paper-size spacers retain the full scrollbar. Only nearby SVGs enter
 * the DOM; both retained paint and the reader cache have byte budgets. */
export function HwpViewport({ info, zoom, navigation, paused, getPage, keepPages, onPageChange, onLoadingChange }: Props) {
  const viewport = useRef<HTMLElement>(null), frame = useRef<number | undefined>(undefined);
  const [view, setView] = useState({ top: 0, height: 700 });
  const [paint, setPaint] = useState(new Map<number, Paint>()), paintRef = useRef(paint);
  const [retry, setRetry] = useState(0);
  const layout = useMemo(() => hwpViewportLayout(info.pages, zoom / 100), [info, zoom]);
  const previous = useRef({ layout, zoom });
  const visible = hwpVisiblePages(layout, Math.max(0, view.top - HWP_VIEW_PADDING), view.height);
  const start = Math.max(0, visible.first - 1), end = Math.min(info.count - 1, visible.last + 1);
  const refresh = () => {
    if (frame.current !== undefined) return;
    frame.current = requestAnimationFrame(() => {
      frame.current = undefined;
      if (viewport.current) setView({ top: viewport.current.scrollTop, height: viewport.current.clientHeight });
    });
  };
  useEffect(() => {
    const element = viewport.current!;
    const observer = new ResizeObserver(refresh); observer.observe(element); refresh();
    return () => { observer.disconnect(); if (frame.current !== undefined) cancelAnimationFrame(frame.current); frame.current = undefined; };
  }, []);
  useLayoutEffect(() => {
    const old = previous.current, element = viewport.current!;
    if (old.zoom !== zoom) {
      const page = hwpPageAtOffset(old.layout, Math.max(0, element.scrollTop - HWP_VIEW_PADDING));
      const within = element.scrollTop - HWP_VIEW_PADDING - old.layout.offsets[page];
      element.scrollTop = HWP_VIEW_PADDING + layout.offsets[page] + within * zoom / old.zoom;
      refresh();
    }
    previous.current = { layout, zoom };
  }, [layout, zoom]);
  useLayoutEffect(() => {
    const element = viewport.current!;
    element.scrollTop = HWP_VIEW_PADDING + layout.offsets[navigation.page]; refresh();
    // Scroll-originated page changes never update navigation, so they cannot
    // snap the viewport back to the top of the active page.
  }, [navigation.id]);
  useEffect(() => { onPageChange(visible.current); }, [visible.current, onPageChange]);
  useEffect(() => { onLoadingChange(!paint.has(visible.current)); }, [paint, visible.current, onLoadingChange]);
  useEffect(() => {
    let active = true;
    const priority = Array.from({ length: end - start + 1 }, (_, i) => start + i).sort((a, b) => {
      const rank = (n: number) => n === visible.current ? 0 : n >= visible.first && n <= visible.last ? 1 : 2;
      return rank(a) - rank(b) || Math.abs(a - visible.current) - Math.abs(b - visible.current);
    });
    const retain = (source: Map<number, Paint>) => {
      let bytes = 0;
      const result = new Map<number, Paint>();
      for (const number of priority) {
        const item = source.get(number); if (!item) continue;
        const size = (item.page?.svg.length ?? 0) * 2;
        if (number !== visible.current && bytes + size > documentLimits.hwpPageCacheBytes) continue;
        bytes += size; result.set(number, item);
      }
      return result;
    };
    paintRef.current = retain(paintRef.current); setPaint(paintRef.current);
    if (!paused) keepPages(priority);
    if (!paused) void (async () => {
      for (const number of priority) {
        if (!active) return;
        if (paintRef.current.has(number)) continue;
        const bytes = [...paintRef.current.values()].reduce((sum, item) => sum + (item.page?.svg.length ?? 0) * 2, 0);
        if (number !== visible.current && bytes >= documentLimits.hwpPageCacheBytes) break;
        let item: Paint;
        try { item = { page: await getPage(number) }; }
        catch (error) { item = { error: error instanceof Error ? error.message : String(error) }; }
        if (!active) return;
        paintRef.current = retain(new Map(paintRef.current).set(number, item)); setPaint(paintRef.current);
      }
    })();
    return () => { active = false; };
  }, [start, end, visible.first, visible.last, visible.current, paused, getPage, keepPages, retry]);
  const extraBottom = Math.max(0, view.height - (layout.heights.at(-1) ?? 0) - HWP_VIEW_PADDING * 2);
  return <main ref={viewport} className="hwp-viewport" aria-label="HWP 문서 스크롤" aria-busy={!paint.has(visible.current) || paused} tabIndex={0} onScroll={refresh}
    onPointerDown={event => { if (!(event.target as Element).closest('button,input,select,a')) viewport.current?.focus({ preventScroll: true }); }}
    onKeyDown={event => {
      if (event.ctrlKey || event.metaKey || (event.target as Element).closest('button,input,select')) return;
      const element = viewport.current!;
      if (event.key === 'PageDown' || event.key === 'PageUp') { event.preventDefault(); element.scrollBy({ top: element.clientHeight * .9 * (event.key === 'PageDown' ? 1 : -1) }); }
      else if (event.key === 'Home' || event.key === 'End') { event.preventDefault(); element.scrollTop = event.key === 'Home' ? 0 : element.scrollHeight; }
    }}>
    <div className="hwp-pages" style={{ width: layout.width, height: layout.height + extraBottom }}>
      {Array.from({ length: end - start + 1 }, (_, offset) => {
        const number = start + offset, size = info.pages[number], item = paint.get(number), current = number === visible.current;
        return <section key={number} className="hwp-sheet-space" data-hwp-page={number + 1} aria-label={`${number + 1}쪽`} aria-current={current ? 'page' : undefined}
          style={{ top: layout.offsets[number], width: size.width * zoom / 100, height: size.height * zoom / 100 }}>
          {item?.page ? <div className="hwp-sheet" data-testid={current ? 'hwp-page' : undefined} aria-label={`${number + 1}쪽`}
            style={{ width: size.width, height: size.height, transform: `scale(${zoom / 100})` }} dangerouslySetInnerHTML={{ __html: item.page.svg }} />
            : <div className="hwp-page-placeholder" aria-busy={!item?.error}>{item?.error ? <div role="alert"><strong>{number + 1}쪽을 표시하지 못했습니다</strong><p>{item.error}</p>
              <button onClick={() => { paintRef.current = new Map(paintRef.current); paintRef.current.delete(number); setRetry(value => value + 1); }}>다시 읽기</button></div> : `${number + 1}쪽을 읽는 중…`}</div>}
          <span className="hwp-page-label" aria-hidden="true">{number + 1} / {info.count}</span>
        </section>;
      })}
    </div>
  </main>;
}
