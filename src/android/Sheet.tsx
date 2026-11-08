import { useEffect, useRef, type ReactNode } from 'react';
export function Sheet({ title, onClose, children }: { title: string; onClose(): void; children: ReactNode }) {
  const ref = useRef<HTMLDialogElement>(null);
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    ref.current?.showModal();
    return () => { if (previous?.isConnected) previous.focus({ preventScroll: true }); };
  }, []);
  return <dialog ref={ref} className="android-sheet" aria-label={title} onCancel={event => { event.preventDefault(); event.stopPropagation(); onClose(); }}
    onClick={event => { if (event.target === ref.current) onClose(); }}>
    <header><h2>{title}</h2><button aria-label={`${title} 닫기`} onClick={onClose}>닫기</button></header>
    <div className="android-sheet-body">{children}</div>
  </dialog>;
}
