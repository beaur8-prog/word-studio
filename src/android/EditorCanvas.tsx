import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import type { PageSettings } from '../types';
import { pageMargins, paperDimensions } from '../types';
import { visibleEditorPages } from '../editor-visible-pages';
import { useDocumentZoom } from '../document-zoom-gesture';

const PX_PER_MM = 96 / 25.4;
/** Paper dimensions are millimetres. The editor is always measured at its real
 * CSS pixel size; only the outer display transform follows the phone width. */
export function EditorCanvas({ settings, initialPage, onCompositionStart, onCompositionEnd, children }: {
  settings: PageSettings; initialPage: number; onCompositionStart(): void; onCompositionEnd(): void; children: ReactNode;
}) {
  const viewport = useRef<HTMLDivElement>(null), paper = useRef<HTMLDivElement>(null), requestedPages = useRef(1);
  const [viewWidth, setViewWidth] = useState(0), [height, setHeight] = useState(0), [pages, setPages] = useState(1), [manualScale, setManualScale] = useState<number | null>(null);
  const [widthMM, heightMM] = paperDimensions(settings), margins = pageMargins(settings);
  const width = widthMM * PX_PER_MM, pageHeight = heightMM * PX_PER_MM;
  const fit = viewWidth ? Math.min(1.5, Math.max(.1, (viewWidth - 12) / width)) : 1, scale = manualScale ?? fit;
  const changeZoom = useDocumentZoom(viewport, paper, scale, setManualScale);
  useLayoutEffect(() => {
    const element = viewport.current!, stack = paper.current!;
    const resize = (requested?: number) => {
      if (requested !== undefined && Number.isFinite(requested)) requestedPages.current = requested;
      const editor = stack.querySelector<HTMLElement>('.tiptap');
      const count = editor ? visibleEditorPages(editor, pageHeight, requestedPages.current) : 1;
      setPages(count); setViewWidth(element.clientWidth); setHeight(Math.max(pageHeight, stack.scrollHeight, count * (pageHeight + 24) - 24));
    };
    const observer = new ResizeObserver(() => resize()); observer.observe(element); observer.observe(stack); resize();
    const pagecount = (event: Event) => requestAnimationFrame(() => resize((event as CustomEvent<number>).detail));
    stack.addEventListener('pagecount', pagecount);
    return () => { observer.disconnect(); stack.removeEventListener('pagecount', pagecount); };
  }, [width, pageHeight]);
  const positioned = useRef(false);
  useEffect(() => {
    if (!positioned.current && viewWidth && viewport.current) { viewport.current.scrollTop = initialPage * (pageHeight + 24) * scale; positioned.current = true; }
  }, [viewWidth, initialPage, pageHeight, scale]);
  return <section className="android-editor-pane"><div ref={viewport} className="android-editor-scroll" data-scale={scale} onCompositionStart={onCompositionStart} onCompositionEnd={onCompositionEnd}>
    <div className="android-editor-stage" style={{ width: width * scale, height: Math.max(pageHeight, height) * scale }}>
      <div ref={paper} className="document-stack" data-page-height={pageHeight} data-margin={margins.top * PX_PER_MM} data-bottom-margin={margins.bottom * PX_PER_MM}
        style={{ width, minHeight: pages * (pageHeight + 24) - 24, transform: `scale(${scale})`, '--page-width': `${width}px`, '--page-height': `${pageHeight}px`, '--page-margin': `${margins.top * PX_PER_MM}px`,
          '--page-padding': `${margins.top * PX_PER_MM}px ${margins.right * PX_PER_MM}px ${margins.bottom * PX_PER_MM}px ${margins.left * PX_PER_MM}px` } as CSSProperties}>
        <div className="paper-backgrounds" aria-hidden="true">{Array.from({ length: pages }, (_, n) => <div key={n} className="paper" style={{ top: n * (pageHeight + 24), height: pageHeight }} />)}</div>
        {children}
      </div>
    </div>
  </div>
    <div className="android-editor-zoom" role="toolbar" aria-label="편집 배율">
      <span>{pages}쪽</span><button aria-label="편집 축소" onMouseDown={e => e.preventDefault()} onClick={() => changeZoom(scale / 1.25)}>−</button>
      <button aria-label="편집 폭 맞춤" onMouseDown={e => e.preventDefault()} onClick={() => changeZoom(fit)}>{Math.round(scale * 100)}%</button>
      <button aria-label="편집 확대" onMouseDown={e => e.preventDefault()} onClick={() => changeZoom(scale * 1.25)}>＋</button>
    </div>
  </section>;
}
