import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import type { PageSettings } from '../types';
import { pageMargins, paperDimensions } from '../types';

const PX_PER_MM = 96 / 25.4;
/** Paper dimensions are millimetres. The editor is always measured at its real
 * CSS pixel size; only the outer display transform follows the phone width. */
export function EditorCanvas({ settings, initialPage, onCompositionStart, onCompositionEnd, children }: {
  settings: PageSettings; initialPage: number; onCompositionStart(): void; onCompositionEnd(): void; children: ReactNode;
}) {
  const viewport = useRef<HTMLDivElement>(null), paper = useRef<HTMLDivElement>(null);
  const [viewWidth, setViewWidth] = useState(0), [height, setHeight] = useState(0);
  const [widthMM, heightMM] = paperDimensions(settings), margins = pageMargins(settings);
  const width = widthMM * PX_PER_MM, pageHeight = heightMM * PX_PER_MM;
  const scale = viewWidth ? Math.min(1.5, Math.max(.1, (viewWidth - 12) / width)) : 1;
  useLayoutEffect(() => {
    const element = viewport.current!, stack = paper.current!;
    const resize = () => { setViewWidth(element.clientWidth); setHeight(Math.max(pageHeight, stack.scrollHeight)); };
    const observer = new ResizeObserver(resize); observer.observe(element); observer.observe(stack); resize();
    const pagecount = () => requestAnimationFrame(resize);
    stack.addEventListener('pagecount', pagecount);
    return () => { observer.disconnect(); stack.removeEventListener('pagecount', pagecount); };
  }, [width, pageHeight]);
  const positioned = useRef(false);
  useEffect(() => {
    if (!positioned.current && viewWidth && viewport.current) { viewport.current.scrollTop = initialPage * (pageHeight + 24) * scale; positioned.current = true; }
  }, [viewWidth, initialPage, pageHeight, scale]);
  return <div ref={viewport} className="android-editor-scroll" onCompositionStart={onCompositionStart} onCompositionEnd={onCompositionEnd}>
    <div className="android-editor-stage" style={{ width: width * scale, height: Math.max(pageHeight, height) * scale }}>
      <div ref={paper} className="document-stack" data-page-height={pageHeight} data-margin={margins.top * PX_PER_MM} data-bottom-margin={margins.bottom * PX_PER_MM}
        style={{ width, minHeight: pageHeight, transform: `scale(${scale})`, '--page-width': `${width}px`, '--page-height': `${pageHeight}px`, '--page-margin': `${margins.top * PX_PER_MM}px`,
          '--page-padding': `${margins.top * PX_PER_MM}px ${margins.right * PX_PER_MM}px ${margins.bottom * PX_PER_MM}px ${margins.left * PX_PER_MM}px` } as CSSProperties}>
        {children}
      </div>
    </div>
  </div>;
}
