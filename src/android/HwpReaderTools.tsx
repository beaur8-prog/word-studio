import { useEffect, useState, type RefObject } from 'react';
import type { HwpInfo } from '../hwp/reader';
import { Sheet } from './Sheet';
import { ModeSwitch } from './ModeSwitch';
import { MoreHorizontal } from 'lucide-react';
import { shareAndroidFile } from './platform';

export function HwpReaderTools({ file, info, index, zoom, disabled, conversionBlocked, query, searching, searchInput, searchControl, onQuery, onFind, onPage, onZoom, onFit, onClose, onOpen, onCopy, onEdit, onPresent }: {
  file: File; info?: HwpInfo; index: number; zoom: number; disabled: boolean; conversionBlocked: boolean; query: string; searching: boolean;
  searchInput: RefObject<HTMLInputElement | null>; searchControl: RefObject<(() => void) | null>; onQuery(value: string): void; onFind(): void; onPage(value: number): void; onZoom(value: number): void;
  onFit(): void; onClose(): void; onOpen(): void; onCopy(): void; onEdit(): void; onPresent(): void;
}) {
  const [panel, setPanel] = useState<'menu' | 'search' | 'page' | 'zoom' | null>(null), [message, setMessage] = useState('');
  useEffect(() => { searchControl.current = () => setPanel('search'); return () => { searchControl.current = null; }; }, [searchControl]);
  return <>
    <header className="android-topbar"><button aria-label="한글 원본 보기 닫기" disabled={disabled} onClick={onClose}>‹</button><div><strong>{file.name}</strong><small>한글 원본 보기</small></div>
      <button disabled={disabled} aria-label="한글 문서 메뉴" onClick={() => setPanel('menu')}><MoreHorizontal size={23} aria-hidden="true" /></button>
      <ModeSwitch editing={false} disabled={!info || disabled} editDisabled={conversionBlocked} onEdit={onEdit} onView={onPresent} viewLabel="한글 전체 화면 뷰어" /></header>
    <nav className="android-reader-tools android-hwp-tools" aria-label="한글 문서 도구">
      <button disabled={!info || disabled} onClick={() => setPanel('search')}>찾기</button>
      <button disabled={!info || disabled} onClick={() => setPanel('page')}>쪽 {index + 1}/{info?.count ?? '…'}</button>
      <button disabled={!info || disabled} onClick={() => setPanel('zoom')}>보기</button>
      <button disabled={disabled} onClick={() => { void shareAndroidFile(file).then(result => setMessage(result === 'completed' ? '공유창에 원본을 전달했습니다' : '공유를 취소했습니다')).catch(error => setMessage(String(error.message || error))); }}>공유</button>
      <button disabled={!info || disabled || conversionBlocked} onClick={onEdit}>편집본</button>
    </nav>
    {message && <p className="android-hwp-message" role="status">{message}</p>}
    {panel && <Sheet title={panel === 'menu' ? '한글 문서 메뉴' : panel === 'search' ? '한글 문서에서 찾기' : panel === 'page' ? '한글 쪽 이동' : '한글 보기 설정'} onClose={() => setPanel(null)}>
      {panel === 'menu' && <><button onClick={() => { setPanel(null); onPresent(); }}>전체 화면 뷰어</button><button onClick={() => { setPanel(null); onOpen(); }}>다른 파일 열기</button><button disabled={!info} onClick={onCopy}>문서 텍스트 복사</button><button onClick={onClose}>원본 보기 닫기</button><p>원본 파일은 변경하지 않습니다. PDF·인쇄는 Android 출력 연결 후 제공됩니다.</p></>}
      {panel === 'search' && <form onSubmit={event => { event.preventDefault(); onFind(); setPanel(null); }}><label>찾을 내용<input autoFocus ref={searchInput} aria-label="HWP 문서에서 찾기" value={query} onChange={event => onQuery(event.target.value)} /></label><button disabled={searching || !query.trim()}>다음 찾기</button></form>}
      {panel === 'page' && <><label>쪽 번호<input type="number" aria-label="HWP 쪽 번호" min="1" max={info?.count} value={index + 1} onChange={event => onPage(Number(event.target.value) - 1)} /></label>
        <div className="android-action-row"><button disabled={index <= 0} onClick={() => onPage(index - 1)}>이전 쪽</button><button disabled={!info || index + 1 >= info.count} onClick={() => onPage(index + 1)}>다음 쪽</button></div><button onClick={() => setPanel(null)}>원본으로 돌아가기</button></>}
      {panel === 'zoom' && <><button onClick={onFit}>폭에 맞추기</button><label>확대<input aria-label="HWP 확대" type="range" min="10" max="300" value={zoom} onChange={event => onZoom(Number(event.target.value))} /><output>{zoom}%</output></label><p>표·직인·그림의 원래 배치를 유지합니다.</p></>}
    </Sheet>}
  </>;
}
