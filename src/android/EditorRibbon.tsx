import { useEffect, useId, useRef, useState, type ReactNode, type MouseEvent } from 'react';
import type { Editor } from '@tiptap/core';
import { AlignCenter, AlignJustify, AlignLeft, AlignRight, Bold, ChevronDown, ChevronUp, Highlighter, Italic,
  List, ListOrdered, Minus, Pilcrow, Redo2, RemoveFormatting, Save, Table2, Underline, Undo2, WrapText } from 'lucide-react';

type Tab = 'home' | 'insert' | 'table';
const tabs: [Tab, string][] = [['home', '홈'], ['insert', '삽입'], ['table', '표']];
const fonts = [['Malgun Gothic', '맑은 고딕'], ['Noto Sans KR', 'Noto Sans KR'], ['Noto Serif KR', 'Noto Serif KR'],
  ['Arial', 'Arial'], ['Times New Roman', 'Times New Roman'], ['Courier New', 'Courier New']];

/** Desktop-style groups, with touch targets and a collapsible ribbon for phones. */
export function EditorRibbon({ editor, disabled, onChange, onSave, onTable }: {
  editor: Editor; disabled: boolean; onChange(): void; onSave(): void; onTable(): void;
}) {
  const [tab, setTab] = useState<Tab>('home'), [expanded, setExpanded] = useState(true);
  const id = useId(), tabButtons = useRef(new Map<Tab, HTMLButtonElement>());
  const format = editor.getAttributes('textStyle');
  const size = parseFloat(format.fontSize || '11'), family = format.fontFamily || 'Malgun Gothic';
  const [draftSize, setDraftSize] = useState(String(size));
  useEffect(() => { setDraftSize(String(size)); }, [size]);
  const preserveSelection = (event: MouseEvent<HTMLButtonElement>) => {
    const field = document.activeElement;
    if (field instanceof HTMLInputElement && field.closest('.android-editor-ribbon')) field.blur();
    event.preventDefault();
  };
  const act = (action: () => void) => { if (!disabled && !editor.view.composing) { action(); onChange(); } };
  const tool = (label: string, icon: ReactNode, action: () => void, active?: boolean, blocked = false, caption = false) =>
    <button type="button" title={label} aria-label={label} aria-pressed={active} disabled={disabled || blocked}
      onMouseDown={preserveSelection} onClick={() => act(action)} className={caption ? 'android-ribbon-caption-button' : undefined}>
      {icon}{caption && <span>{label}</span>}
    </button>;
  const group = (name: string, content: ReactNode) => <section className="android-ribbon-group" aria-label={name}>
    <div className="android-ribbon-group-body">{content}</div><small>{name}</small>
  </section>;
  const commitSize = () => {
    const next = Number(draftSize);
    if (Number.isFinite(next) && next >= 6 && next <= 96) { if (next !== size) act(() => { editor.chain().setFontSize(`${next}pt`).run(); }); }
    else setDraftSize(String(size));
  };
  const chooseTab = (next: Tab) => { setTab(next); setExpanded(true); };
  return <section className="android-editor-ribbon" aria-label="편집 리본">
    <div className="android-ribbon-tab-row">
      <div className="android-quick-tools" role="toolbar" aria-label="빠른 실행">
        {tool('실행 취소', <Undo2 size={20} />, () => { editor.chain().focus().undo().run(); }, undefined, !editor.can().undo())}
        {tool('다시 실행', <Redo2 size={20} />, () => { editor.chain().focus().redo().run(); }, undefined, !editor.can().redo())}
        {tool('문서 사본 저장', <Save size={20} />, onSave)}
      </div>
      <div className="android-ribbon-tabs" role="tablist" aria-label="편집 도구 탭">
        {tabs.map(([key, label], index) => <button type="button" role="tab" key={key} id={`${id}-${key}`}
          ref={element => { if (element) tabButtons.current.set(key, element); else tabButtons.current.delete(key); }}
          aria-selected={tab === key} aria-controls={`${id}-panel`} tabIndex={tab === key ? 0 : -1}
          onMouseDown={preserveSelection} onClick={() => chooseTab(key)} onKeyDown={event => {
            const next = event.key === 'ArrowRight' ? (index + 1) % tabs.length : event.key === 'ArrowLeft' ? (index + tabs.length - 1) % tabs.length
              : event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : -1;
            if (next >= 0) { event.preventDefault(); chooseTab(tabs[next][0]); tabButtons.current.get(tabs[next][0])?.focus(); }
          }}>{label}</button>)}
      </div>
      <button type="button" className="android-ribbon-collapse" aria-label={expanded ? '도구 모음 접기' : '도구 모음 펼치기'}
        aria-expanded={expanded} aria-controls={`${id}-panel`} onMouseDown={preserveSelection} onClick={() => setExpanded(value => !value)}>
        {expanded ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
      </button>
    </div>
    <div className="android-ribbon-content" id={`${id}-panel`} role="tabpanel" aria-labelledby={`${id}-${tab}`} hidden={!expanded}>
      {tab === 'home' && <>
        {group('글꼴', <><div className="android-ribbon-row">
          <select aria-label="글꼴" value={family} disabled={disabled} onChange={event => act(() => { editor.chain().focus().setFontFamily(event.target.value).run(); })}>
            {!fonts.some(([value]) => value === family) && <option value={family}>{family}</option>}
            {fonts.map(([value, name]) => <option key={value} value={value}>{name}</option>)}
          </select>
          <input aria-label="글꼴 크기" title="글꼴 크기 (pt)" type="number" min="6" max="96" step="0.5" inputMode="decimal"
            disabled={disabled} value={draftSize} onChange={event => setDraftSize(event.target.value)} onBlur={commitSize}
            onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); commitSize(); editor.commands.focus(); } }} />
        </div><div className="android-ribbon-row">
          {tool('굵게', <Bold size={20} />, () => { editor.chain().focus().toggleBold().run(); }, editor.isActive('bold'))}
          {tool('기울임', <Italic size={20} />, () => { editor.chain().focus().toggleItalic().run(); }, editor.isActive('italic'))}
          {tool('밑줄', <Underline size={20} />, () => { editor.chain().focus().toggleUnderline().run(); }, editor.isActive('underline'))}
          {tool('형광펜', <Highlighter size={20} />, () => { editor.chain().focus().toggleHighlight().run(); }, editor.isActive('highlight'))}
        </div></>)}
        {group('단락', <><div className="android-ribbon-row">
          {tool('왼쪽 정렬', <AlignLeft size={20} />, () => { editor.chain().focus().setTextAlign('left').run(); }, editor.isActive({ textAlign: 'left' }))}
          {tool('가운데 정렬', <AlignCenter size={20} />, () => { editor.chain().focus().setTextAlign('center').run(); }, editor.isActive({ textAlign: 'center' }))}
          {tool('오른쪽 정렬', <AlignRight size={20} />, () => { editor.chain().focus().setTextAlign('right').run(); }, editor.isActive({ textAlign: 'right' }))}
          {tool('양쪽 정렬', <AlignJustify size={20} />, () => { editor.chain().focus().setTextAlign('justify').run(); }, editor.isActive({ textAlign: 'justify' }))}
        </div><div className="android-ribbon-row">
          {tool('글머리 기호', <List size={20} />, () => { editor.chain().focus().toggleBulletList().run(); }, editor.isActive('bulletList'))}
          {tool('번호 매기기', <ListOrdered size={20} />, () => { editor.chain().focus().toggleOrderedList().run(); }, editor.isActive('orderedList'))}
          {tool('서식 지우기', <RemoveFormatting size={20} />, () => { editor.chain().focus().unsetAllMarks().clearNodes().run(); })}
          {tool('표 도구', <Table2 size={20} />, onTable)}
        </div></>)}
      </>}
      {tab === 'insert' && <>
        {group('페이지', <div className="android-ribbon-row">
          {tool('쪽 나누기', <WrapText size={24} />, () => { editor.chain().focus().insertContent([{ type: 'pageBreak' }, { type: 'paragraph' }]).run(); }, undefined, false, true)}
          {tool('구분선', <Minus size={24} />, () => { editor.chain().focus().setHorizontalRule().run(); }, undefined, false, true)}
        </div>)}
        {group('표', <div className="android-ribbon-row">{tool('표 삽입', <Table2 size={24} />, onTable, undefined, false, true)}</div>)}
      </>}
      {tab === 'table' && <>
        {group('표', <div className="android-ribbon-row">{tool('표 도구', <Table2 size={24} />, onTable, undefined, false, true)}</div>)}
        {group('행·열', <div className="android-ribbon-row">
          {tool('아래에 행 삽입', <Pilcrow size={24} />, () => { editor.chain().focus().addRowAfter().run(); }, undefined, !editor.isActive('table'), true)}
          {tool('오른쪽에 열 삽입', <Table2 size={24} />, () => { editor.chain().focus().addColumnAfter().run(); }, undefined, !editor.isActive('table'), true)}
          {tool('셀 병합', <Table2 size={24} />, () => { editor.chain().focus().mergeCells().run(); }, undefined, !editor.can().mergeCells(), true)}
          {tool('셀 나누기', <Table2 size={24} />, () => { editor.chain().focus().splitCell().run(); }, undefined, !editor.can().splitCell(), true)}
        </div>)}
      </>}
    </div>
  </section>;
}
