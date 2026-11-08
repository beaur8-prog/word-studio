import { useEffect, useLayoutEffect, useRef, useState } from 'react';

type Size = { width: number; height: number };
type Drag = { x: number; y: number; left: number; top: number; at: number; moved: boolean; multi: boolean; leftEdge: boolean; rightEdge: boolean; topEdge: boolean; bottomEdge: boolean };
type Pinch = { distance: number; factor: number; x: number; y: number };
/** A single real document page, with touch gestures confined to its viewport.
 * Zoom never modifies the page/model; after a pinch, the remaining finger
 * cannot accidentally advance the document. */
export function PagePresentation({ source, sizes, index, loadPage, onPage, onExit, onEdit, notice = '' }: {
  source: object; sizes: Size[]; index: number; loadPage(index: number): Promise<HTMLElement>;
  onPage(index: number): void; onExit(): void; onEdit?(): void; notice?: string;
}) {
  const root = useRef<HTMLElement>(null);
  const viewport = useRef<HTMLDivElement>(null), content = useRef<HTMLDivElement>(null);
  const [box, setBox] = useState({ width: 1, height: 1 }), [factor, setFactor] = useState(1);
  const [loaded, setLoaded] = useState(false), [error, setError] = useState(''), [retry, setRetry] = useState(0);
  const [controls, setControls] = useState(true);
  const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const points = useRef(new Map<number, { x: number; y: number }>()), drag = useRef<Drag | null>(null), pinch = useRef<Pinch | null>(null);
  const anchor = useRef<{ x: number; y: number; cx: number; cy: number } | null>(null);
  const lastTap = useRef({ at: 0, x: 0, y: 0 });
  const callbacks = useRef({ loadPage, onPage, onExit }); callbacks.current = { loadPage, onPage, onExit };
  const size = sizes[index] ?? sizes[0] ?? { width: 794, height: 1123 };
  const base = box.width / size.width, scale = base * factor;
  const width = size.width * scale, height = size.height * scale;
  const stageWidth = Math.max(box.width, width), stageHeight = Math.max(box.height, height);
  const left = (stageWidth - width) / 2, top = (stageHeight - height) / 2;
  const whole = Math.min(1, box.height / (size.height * base)), minimum = Math.min(.25, whole);
  const showControls = () => { setControls(true); clearTimeout(timer.current); timer.current = setTimeout(() => setControls(false), 2200); };
  useEffect(() => {
    const element = root.current!; const close = () => callbacks.current.onExit();
    element.addEventListener('studio-presentation-close', close);
    return () => element.removeEventListener('studio-presentation-close', close);
  }, []);
  useEffect(() => { showControls(); return () => clearTimeout(timer.current); }, []);
  useLayoutEffect(() => {
    const root = viewport.current!;
    const resize = () => setBox({ width: root.clientWidth, height: root.clientHeight });
    const observer = new ResizeObserver(resize); observer.observe(root); resize();
    return () => observer.disconnect();
  }, []);
  useEffect(() => {
    let active = true; const root = content.current!;
    setLoaded(false); setError(''); setFactor(1); points.current.clear(); drag.current = null; pinch.current = null; anchor.current = null;
    void callbacks.current.loadPage(index).then(element => {
      if (!active) return;
      element.style.transform = 'none'; element.style.margin = '0';
      root.replaceChildren(element); setLoaded(true);
    }).catch(reason => { if (active) setError(reason instanceof Error ? reason.message : String(reason)); });
    return () => { active = false; root.replaceChildren(); };
  }, [source, index, retry]);
  useLayoutEffect(() => {
    const root = viewport.current!, point = anchor.current;
    if (point) { root.scrollLeft = point.x * scale + left - point.cx; root.scrollTop = point.y * scale + top - point.cy; anchor.current = null; }
    else if (factor === 1) { root.scrollLeft = 0; root.scrollTop = 0; }
  }, [scale, left, top, loaded]);
  const zoom = (next: number, cx = box.width / 2, cy = box.height / 2, fixed?: { x: number; y: number }) => {
    const root = viewport.current!;
    anchor.current = { x: fixed?.x ?? (root.scrollLeft + cx - left) / scale, y: fixed?.y ?? (root.scrollTop + cy - top) / scale, cx, cy };
    setFactor(Math.max(minimum, Math.min(8, next)));
  };
  const go = (next: number) => { if (next >= 0 && next < sizes.length && next !== index) { callbacks.current.onPage(next); showControls(); } };
  const local = (x: number, y: number) => { const rect = viewport.current!.getBoundingClientRect(); return { x: x - rect.left, y: y - rect.top }; };
  const beginPinch = () => {
    const [a, b] = [...points.current.values()], root = viewport.current!;
    if (!a || !b) return;
    const center = local((a.x + b.x) / 2, (a.y + b.y) / 2);
    pinch.current = { distance: Math.max(1, Math.hypot(a.x - b.x, a.y - b.y)), factor,
      x: (root.scrollLeft + center.x - left) / scale, y: (root.scrollTop + center.y - top) / scale };
    if (drag.current) drag.current.multi = true;
  };
  useEffect(() => {
    const key = (event: KeyboardEvent) => {
      if ((event.target as HTMLElement).closest('input,textarea,select')) return;
      const actions: Record<string, () => void> = { ArrowRight: () => go(index + 1), ArrowDown: () => go(index + 1), PageDown: () => go(index + 1),
        ArrowLeft: () => go(index - 1), ArrowUp: () => go(index - 1), PageUp: () => go(index - 1), Home: () => go(0), End: () => go(sizes.length - 1),
        '+': () => zoom(factor * 1.25), '-': () => zoom(factor / 1.25), Escape: () => callbacks.current.onExit() };
      if (actions[event.key]) { event.preventDefault(); event.stopPropagation(); actions[event.key](); }
    };
    window.addEventListener('keydown', key, true); return () => window.removeEventListener('keydown', key, true);
  }, [index, factor, scale, box, sizes.length]);
  return <section ref={root} className="android-presentation" aria-label="전체 화면 문서 뷰어" data-page={index + 1} data-controls={controls}>
    <div ref={viewport} className="android-presentation-viewport" tabIndex={0} aria-label={`${index + 1}쪽 확대·이동`} onPointerDown={event => {
      if (event.pointerType === 'mouse' || !loaded) return;
      event.currentTarget.setPointerCapture(event.pointerId); points.current.set(event.pointerId, { x: event.clientX, y: event.clientY });
      if (points.current.size === 1) { const root = viewport.current!; drag.current = { x: event.clientX, y: event.clientY, left: root.scrollLeft, top: root.scrollTop,
        at: performance.now(), moved: false, multi: false, leftEdge: root.scrollLeft <= 2, rightEdge: root.scrollLeft >= root.scrollWidth - root.clientWidth - 2,
        topEdge: root.scrollTop <= 2, bottomEdge: root.scrollTop >= root.scrollHeight - root.clientHeight - 2 }; }
      else beginPinch();
    }} onPointerMove={event => {
      if (!points.current.has(event.pointerId)) return;
      points.current.set(event.pointerId, { x: event.clientX, y: event.clientY });
      const [a, b] = [...points.current.values()], gesture = pinch.current;
      if (a && b && gesture) {
        const center = local((a.x + b.x) / 2, (a.y + b.y) / 2);
        zoom(gesture.factor * Math.hypot(a.x - b.x, a.y - b.y) / gesture.distance, center.x, center.y, gesture);
      } else if (drag.current && !drag.current.multi) {
        const dx = event.clientX - drag.current.x, dy = event.clientY - drag.current.y;
        if (Math.hypot(dx, dy) > 10) drag.current.moved = true;
        viewport.current!.scrollLeft = drag.current.left - dx; viewport.current!.scrollTop = drag.current.top - dy;
      }
    }} onPointerUp={event => {
      const active = points.current.has(event.pointerId), gesture = drag.current;
      points.current.delete(event.pointerId);
      if (!active || !gesture) return;
      if (!points.current.size) {
        if (!gesture.multi) {
          const dx = event.clientX - gesture.x, dy = event.clientY - gesture.y, elapsed = performance.now() - gesture.at;
          if (gesture.moved && elapsed < 750) {
            const horizontal = Math.abs(dx) > Math.abs(dy) * 1.2;
            const vertical = Math.abs(dy) > Math.abs(dx) * 1.2;
            if (horizontal && Math.abs(dx) >= 56 && (dx < 0 ? gesture.rightEdge : gesture.leftEdge)) go(index + (dx < 0 ? 1 : -1));
            else if (vertical && Math.abs(dy) >= 56 && (dy < 0 ? gesture.bottomEdge : gesture.topEdge)) go(index + (dy < 0 ? 1 : -1));
          } else if (!gesture.moved) {
            const now = performance.now(), point = local(event.clientX, event.clientY);
            if (now - lastTap.current.at < 320 && Math.hypot(lastTap.current.x - event.clientX, lastTap.current.y - event.clientY) < 28) { zoom(factor > 1.1 ? 1 : 2, point.x, point.y); lastTap.current.at = 0; showControls(); }
            else { lastTap.current = { at: now, x: event.clientX, y: event.clientY }; if (controls) { setControls(false); clearTimeout(timer.current); } else showControls(); }
          }
        }
        drag.current = null; pinch.current = null;
      }
    }} onPointerCancel={() => { points.current.clear(); drag.current = null; pinch.current = null; }}>
      <div className="print-layout android-presentation-canvas" style={{ width: stageWidth, height: stageHeight }}>
        <div ref={content} className="android-presentation-page" style={{ left, top, width: size.width, height: size.height, transform: `scale(${scale})` }} />
      </div>
    </div>
    <button className="android-presentation-exit" aria-label="뷰어 닫기" onClick={onExit}>‹</button>
    <div className="android-presentation-controls" hidden={!controls}>
      <div className="android-presentation-top"><span role="status">{index + 1} / {sizes.length}</span>{onEdit && <button onClick={onEdit}>편집</button>}</div>
      <footer><button aria-label="이전 페이지" disabled={index === 0} onClick={() => go(index - 1)}>‹</button>
        <button aria-label="뷰어 축소" onClick={() => { zoom(factor / 1.25); showControls(); }}>−</button>
        <button aria-label="뷰어 폭 맞춤" onClick={() => { zoom(1); showControls(); }}>{Math.round(scale * 100)}%</button>
        <button aria-label="뷰어 확대" onClick={() => { zoom(factor * 1.25); showControls(); }}>＋</button>
        <button aria-label="다음 페이지" disabled={index + 1 === sizes.length} onClick={() => go(index + 1)}>›</button>
      </footer>
      <div className="android-presentation-options"><button onClick={() => { zoom(whole); showControls(); }}>쪽 전체</button><span>좌우로 넘기기 · 확대 후 끌어서 이동</span></div>
      {notice && <p role="status">{notice}</p>}
    </div>
    {!loaded && <div className="android-presentation-loading" role={error ? 'alert' : 'status'}>{error ? <>{error}<button onClick={() => setRetry(value => value + 1)}>다시 읽기</button></> : `${index + 1}쪽을 읽는 중…`}</div>}
  </section>;
}
