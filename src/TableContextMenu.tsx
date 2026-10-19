import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { deleteCellSelection } from '@tiptap/pm/tables';
import { AlignLeft, AlignCenter, AlignRight, AlignJustify, Bold, Italic, Underline, Strikethrough, X, Copy, Scissors, ClipboardPaste } from 'lucide-react';
import { pasteClipboard } from './clipboard';
import {
  selectedCells, selectContextCell, setCellStyles, formatCellText, cellTextValue,
  allCellsHaveMark, toggleCellMark, setCellHighlight, formatCellParagraphs, clearCellFormatting,
  applyCellBorders, selectTableRegion, distributeColumns, divideSelectedCells, type BorderPreset,
} from './table-format';

const colors = [
  '#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47',
  '#f2f2f2', '#7f7f7f', '#d0cece', '#d6dce4', '#d9e2f3', '#fbe4d5', '#ededed', '#fff2cc', '#deebf7', '#e2efd9',
  '#d9d9d9', '#595959', '#aeaaaa', '#adb9ca', '#b4c6e7', '#f8cbad', '#dbdbdb', '#ffe599', '#bdd7ee', '#c5e0b3',
  '#bfbfbf', '#3f3f3f', '#757171', '#8496b0', '#8eaadb', '#f4b183', '#c9c9c9', '#ffd966', '#9dc3e6', '#a8d08d',
  '#a6a6a6', '#262626', '#3a3838', '#323e4f', '#2f5496', '#c55a11', '#7b7b7b', '#bf9000', '#2e75b5', '#548235',
  '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0',
];
const fonts = ['맑은 고딕', '바탕', '굴림', '돋움', '궁서', 'Arial', 'Calibri', 'Cambria', 'Times New Roman', 'Georgia', 'Consolas'];
const presets: [BorderPreset, string][] = [
  ['all', '모든 테두리'], ['outer', '바깥쪽 테두리'], ['inner', '안쪽 테두리'], ['none', '테두리 없음'],
  ['top', '위쪽 테두리'], ['bottom', '아래쪽 테두리'], ['left', '왼쪽 테두리'], ['right', '오른쪽 테두리'],
  ['horizontal', '안쪽 가로선'], ['vertical', '안쪽 세로선'],
];

function Palette({ label, value, onChange, onReset }: { label: string; value: string; onChange: (color: string) => void; onReset?: () => void }) {
  return <div className="cell-palette">
    <div className="cell-swatches" role="group" aria-label={`${label} 팔레트`}>
      {colors.map((color, i) => <button key={i} type="button" aria-label={`${label} ${color}`} title={color}
        aria-pressed={value.toLowerCase() === color} style={{ backgroundColor: color }} onClick={() => onChange(color)} />)}
    </div>
    <div className="cell-custom-color">
      {onReset && <button type="button" onClick={onReset}>기본값 / 색 없음</button>}
      <label>사용자 지정 <input type="color" aria-label={`${label} 사용자 지정`} value={/^#[\da-f]{6}$/i.test(value) ? value : '#000000'} onChange={e => onChange(e.target.value)} /></label>
    </div>
  </div>;
}

export function TableContextMenu({ editor, readOnly, notify }: { editor: Editor; readOnly: boolean; notify: (message: string) => void }) {
  const [point, setPoint] = useState<{ x: number; y: number } | null>(null);
  const [tab, setTab] = useState('format');
  const [colorTarget, setColorTarget] = useState<'color' | 'backgroundColor' | 'highlight'>('backgroundColor');
  const [lastColors, setLastColors] = useState({ color: '#000000', backgroundColor: '#fff2cc', highlight: '#ffff00' });
  const [borderColor, setBorderColor] = useState('#44546a');
  const [borderStyle, setBorderStyle] = useState('solid');
  const [borderWidth, setBorderWidth] = useState('0.75');
  const [font, setFont] = useState('');
  const [fontSize, setFontSize] = useState('');
  const [padding, setPadding] = useState('8');
  const [splitRows, setSplitRows] = useState(1), [splitColumns, setSplitColumns] = useState(2);
  const [, setRevision] = useState(0);
  const root = useRef<HTMLDivElement>(null);
  const restoreFocus = () => { setPoint(null); editor.view.focus(); };

  useEffect(() => {
    if (readOnly) { setPoint(null); return; }
    const dom = editor.view.dom;
    const open = (event: MouseEvent | KeyboardEvent) => {
      if (!editor.isEditable || editor.view.composing) return;
      const target = event instanceof MouseEvent && event.target instanceof Element ? event.target.closest<HTMLElement>('td,th') : undefined;
      if (event instanceof MouseEvent && (!target || !dom.contains(target))) return;
      if (!selectContextCell(editor, target ?? undefined)) return;
      event.preventDefault(); event.stopPropagation();
      const cell = selectedCells(editor)[0];
      const box = (editor.view.nodeDOM(cell.pos) as HTMLElement).getBoundingClientRect();
      setPoint(event instanceof MouseEvent ? { x: event.clientX, y: event.clientY } : { x: box.left + 12, y: box.top + 12 });
      setFont(cellTextValue(editor, 'fontFamily').replace(/["']/g, ''));
      const size = cellTextValue(editor, 'fontSize');
      setFontSize(size ? String(parseFloat(size) * (size.endsWith('px') ? 0.75 : 1)) : '');
      setPadding(String(parseFloat(cell.node.attrs.paddingTop || '8')));
    };
    const key = (e: KeyboardEvent) => { if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) open(e); };
    const update = () => setRevision(n => n + 1);
    const selectionChanged = () => setPoint(null);
    dom.addEventListener('contextmenu', open);
    dom.addEventListener('keydown', key);
    editor.on('transaction', update);
    editor.on('selectionUpdate', selectionChanged);
    return () => {
      dom.removeEventListener('contextmenu', open); dom.removeEventListener('keydown', key);
      editor.off('transaction', update); editor.off('selectionUpdate', selectionChanged);
    };
  }, [editor, readOnly]);

  useLayoutEffect(() => {
    if (!point || !root.current) return;
    const el = root.current, rect = el.getBoundingClientRect();
    el.style.left = `${Math.max(8, Math.min(point.x, window.innerWidth - rect.width - 8))}px`;
    el.style.top = `${Math.max(8, Math.min(point.y, window.innerHeight - rect.height - 8))}px`;
  }, [point, tab]);
  useEffect(() => {
    if (!point) return;
    root.current?.querySelector<HTMLButtonElement>('.cell-menu-close')?.focus({ preventScroll: true });
    const outside = (e: Event) => { if (!root.current?.contains(e.target as Node)) setPoint(null); };
    const resize = () => setPoint(null);
    document.addEventListener('pointerdown', outside, true);
    document.addEventListener('scroll', outside, true);
    window.addEventListener('resize', resize);
    return () => {
      document.removeEventListener('pointerdown', outside, true); document.removeEventListener('scroll', outside, true);
      window.removeEventListener('resize', resize);
    };
  }, [point]);

  if (!point || readOnly) return null;
  const cells = selectedCells(editor);
  if (!cells.length) return null;
  const run = (action: () => unknown) => { action(); restoreFocus(); };
  const applyColor = (color: string | null) => {
    if (color) setLastColors(previous => ({ ...previous, [colorTarget]: color }));
    if (colorTarget === 'color') formatCellText(editor, 'color', color);
    else if (colorTarget === 'backgroundColor') setCellStyles(editor, { backgroundColor: color ?? 'transparent' });
    else setCellHighlight(editor, color);
  };
  const commitSize = () => {
    const size = Number(fontSize);
    if (Number.isFinite(size) && size >= 1 && size <= 400) {
      if (cellTextValue(editor, 'fontSize') !== `${size}pt`) formatCellText(editor, 'fontSize', `${size}pt`);
    }
    else setFontSize(cellTextValue(editor, 'fontSize').replace(/pt$/, ''));
  };
  const copy = (cut: boolean) => {
    editor.view.focus();
    if (document.execCommand('copy')) {
      if (cut) deleteCellSelection(editor.state, tr => editor.view.dispatch(closeHistory(tr)));
      notify(cut ? '선택한 셀 내용을 잘라냈습니다.' : '선택한 셀을 복사했습니다.');
    } else notify('문서에서 Ctrl+C 또는 Ctrl+X를 눌러 주세요.');
    setPoint(null);
  };
  const paste = async () => {
    restoreFocus();
    await pasteClipboard(editor, notify);
  };
  const label = { color: '글자 색', backgroundColor: '셀 배경', highlight: '형광펜' }[colorTarget];
  return createPortal(<div ref={root} className="table-context-menu" role="dialog" aria-label="표 셀 편집" style={{ left: point.x, top: point.y }}
    onContextMenu={e => e.preventDefault()} onKeyDown={e => {
      e.stopPropagation();
      if (e.key === 'Escape') { e.preventDefault(); restoreFocus(); }
      if ((e.ctrlKey || e.metaKey) && ['z', 'y'].includes(e.key.toLowerCase()) && e.target instanceof HTMLButtonElement) {
        e.preventDefault(); run(() => e.key.toLowerCase() === 'y' || e.shiftKey ? editor.commands.redo() : editor.commands.undo());
      }
      if (e.key === 'Tab') {
        const focusable = Array.from(root.current!.querySelectorAll<HTMLElement>('button:not(:disabled),input,select'));
        const first = focusable[0], last = focusable[focusable.length - 1];
        if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
        else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      }
    }}>
    <header><div><strong>표 셀 편집</strong><span>{cells.length}개 셀 선택됨</span></div><button className="cell-menu-close" aria-label="표 메뉴 닫기" onClick={restoreFocus}><X size={17} /></button></header>
    <div className="cell-menu-clipboard">
      <button onClick={() => copy(false)}><Copy size={14} />복사</button>
      <button onClick={() => copy(true)}><Scissors size={14} />잘라내기</button>
      <button onClick={paste}><ClipboardPaste size={14} />붙여넣기</button>
    </div>
    <div className="cell-menu-tabs" role="tablist" aria-label="셀 편집 종류">
      {[['format', '글자·셀 서식'], ['borders', '테두리'], ['table', '표 편집']].map(([id, name]) =>
        <button key={id} role="tab" id={`cell-tab-${id}`} aria-selected={tab === id} aria-controls={`cell-panel-${id}`} onClick={() => setTab(id)}>{name}</button>)}
    </div>
    <div className="cell-menu-body" role="tabpanel" id={`cell-panel-${tab}`} aria-labelledby={`cell-tab-${tab}`}>
      {tab === 'format' && <>
        <div className="cell-font-fields">
          <label>글꼴<select aria-label="셀 글꼴" value={font} onChange={e => { setFont(e.target.value); formatCellText(editor, 'fontFamily', e.target.value || null); }}>
            <option value="">기본값 / 혼합</option>{!fonts.includes(font) && font && <option>{font}</option>}{fonts.map(name => <option key={name}>{name}</option>)}
          </select></label>
          <label>크기 (pt)<input type="number" aria-label="셀 글꼴 크기" min="1" max="400" step="0.5" placeholder="혼합" value={fontSize} onChange={e => setFontSize(e.target.value)} onBlur={commitSize} onKeyDown={e => { if (e.key === 'Enter') commitSize(); }} /></label>
        </div>
        <div className="cell-inline-tools" role="group" aria-label="셀 글자 서식">
          {([['bold', '굵게', Bold], ['italic', '기울임', Italic], ['underline', '밑줄', Underline], ['strike', '취소선', Strikethrough]] as const).map(([name, title, Icon]) =>
            <button key={name} title={title} aria-label={title} aria-pressed={allCellsHaveMark(editor, name)} onClick={() => toggleCellMark(editor, name)}><Icon size={17} /></button>)}
          <span />
          {([['left', '왼쪽 정렬', AlignLeft], ['center', '가운데 정렬', AlignCenter], ['right', '오른쪽 정렬', AlignRight], ['justify', '양쪽 정렬', AlignJustify]] as const).map(([value, title, Icon]) =>
            <button key={value} title={title} aria-label={title} onClick={() => formatCellParagraphs(editor, { textAlign: value })}><Icon size={17} /></button>)}
        </div>
        <div className="cell-color-target" role="group" aria-label="색 적용 대상">{(['color', 'backgroundColor', 'highlight'] as const).map(value =>
          <button key={value} aria-pressed={colorTarget === value} onClick={() => setColorTarget(value)}>{{ color: '글자 색', backgroundColor: '셀 배경', highlight: '형광펜' }[value]}</button>)}</div>
        <Palette label={label} value={lastColors[colorTarget]} onChange={applyColor} onReset={() => applyColor(null)} />
        <div className="cell-settings-row">
          <label>세로 정렬<select aria-label="셀 세로 정렬" value={cells.every(c => c.node.attrs.verticalAlign === cells[0].node.attrs.verticalAlign) ? cells[0].node.attrs.verticalAlign || 'top' : ''} onChange={e => setCellStyles(editor, { verticalAlign: e.target.value })}>
            <option value="" disabled>혼합</option><option value="top">위쪽</option><option value="middle">가운데</option><option value="bottom">아래쪽</option>
          </select></label>
          <label>줄 간격<select aria-label="셀 줄 간격" defaultValue="" onChange={e => formatCellParagraphs(editor, { lineSpacing: e.target.value })}>
            <option value="" disabled>선택</option>{['1', '1.15', '1.5', '2'].map(value => <option key={value}>{value}</option>)}
          </select></label>
        </div>
        <button className="cell-full-button" onClick={() => clearCellFormatting(editor)}>선택한 셀 서식 지우기</button>
      </>}
      {tab === 'borders' && <>
        <div className="cell-settings-row">
          <label>선 종류<select aria-label="테두리 선 종류" value={borderStyle} onChange={e => setBorderStyle(e.target.value)}>
            <option value="solid">실선</option><option value="dashed">파선</option><option value="dotted">점선</option><option value="double">이중선</option>
          </select></label>
          <label>굵기 (pt)<select aria-label="테두리 굵기" value={borderWidth} onChange={e => setBorderWidth(e.target.value)}>{['0.5', '0.75', '1', '1.5', '2.25', '3', '4.5', '6'].map(value => <option key={value}>{value}</option>)}</select></label>
        </div>
        <div className="cell-border-preview"><span style={{ borderBottom: `${borderWidth}pt ${borderStyle} ${borderColor}` }} />선 모양을 고른 뒤 적용할 테두리를 누르세요.</div>
        <Palette label="테두리 색" value={borderColor} onChange={setBorderColor} />
        <div className="cell-action-grid">{presets.map(([id, title]) => <button key={id} onClick={() => applyCellBorders(editor, id, `${borderWidth}pt ${borderStyle} ${borderColor}`)}>{title}</button>)}</div>
      </>}
      {tab === 'table' && <>
        <h4>선택</h4><div className="cell-action-grid three">{([['row', '행 선택'], ['column', '열 선택'], ['table', '표 전체 선택']] as const).map(([region, title]) => <button key={region} onClick={() => run(() => selectTableRegion(editor, region))}>{title}</button>)}</div>
        <h4>삽입</h4><div className="cell-action-grid">
          <button onClick={() => run(() => editor.commands.addRowBefore())}>위에 행 삽입</button><button onClick={() => run(() => editor.commands.addRowAfter())}>아래에 행 삽입</button>
          <button onClick={() => run(() => editor.commands.addColumnBefore())}>왼쪽에 열 삽입</button><button onClick={() => run(() => editor.commands.addColumnAfter())}>오른쪽에 열 삽입</button>
        </div>
        <h4>구조와 크기</h4><div className="cell-action-grid">
          <button disabled={!editor.can().mergeCells()} onClick={() => run(() => editor.commands.mergeCells())}>셀 병합</button>
          <button disabled={!editor.can().splitCell()} onClick={() => run(() => editor.commands.splitCell())}>병합 셀 분할</button>
          <button onClick={() => run(() => editor.commands.toggleHeaderRow())}>머리글 행 전환</button>
          <button onClick={() => run(() => distributeColumns(editor))}>모든 열 너비 같게</button>
        </div>
        <div className="cell-padding"><label>셀 안쪽 여백 (px)<input type="number" min="0" max="96" aria-label="셀 안쪽 여백" value={padding} onChange={e => setPadding(e.target.value)} /></label>
          <button disabled={padding.trim() === '' || !Number.isFinite(Number(padding)) || Number(padding) < 0 || Number(padding) > 96} onClick={() => setCellStyles(editor, { paddingTop: `${padding}px`, paddingRight: `${padding}px`, paddingBottom: `${padding}px`, paddingLeft: `${padding}px` })}>적용</button></div>
        <div className="cell-split-fields"><label>나눌 행 수<input type="number" aria-label="나눌 행 수" min="1" max="10" value={splitRows} onChange={e => setSplitRows(Number(e.target.value))} /></label>
          <label>나눌 열 수<input type="number" aria-label="나눌 열 수" min="1" max="10" value={splitColumns} onChange={e => setSplitColumns(Number(e.target.value))} /></label>
          <button disabled={!Number.isInteger(splitRows) || !Number.isInteger(splitColumns) || splitRows < 1 || splitColumns < 1 || splitRows > 10 || splitColumns > 10 || splitRows * splitColumns < 2}
            onClick={() => run(() => { if (!divideSelectedCells(editor, splitRows, splitColumns)) notify('표가 너무 커서 셀을 나누지 못했습니다. 더 작은 범위를 선택해 주세요.'); })}>각 셀 나누기</button></div>
        <h4>내용과 삭제</h4><div className="cell-action-grid">
          <button onClick={() => run(() => deleteCellSelection(editor.state, tr => editor.view.dispatch(closeHistory(tr))))}>셀 내용 지우기</button>
          <button onClick={() => run(() => editor.commands.deleteRow())}>선택한 행 삭제</button>
          <button onClick={() => run(() => editor.commands.deleteColumn())}>선택한 열 삭제</button>
          <button className="cell-delete-table" onClick={() => run(() => editor.commands.deleteTable())}>표 삭제</button>
        </div>
      </>}
    </div>
    <footer>선택한 셀에 적용 · Esc로 닫기 · Ctrl+Z로 실행 취소</footer>
  </div>, document.body);
}
