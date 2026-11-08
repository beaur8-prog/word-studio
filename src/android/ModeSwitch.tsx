import type { MouseEvent } from 'react';
import { BookOpen, Pencil } from 'lucide-react';

/** The same two targets stay in the top right of the editor and full-screen viewer. */
export function ModeSwitch({ editing, disabled = false, editDisabled = false, onEdit, onView, viewLabel = '전체 화면 뷰어' }: {
  editing: boolean; disabled?: boolean; editDisabled?: boolean; onEdit?(): void; onView(): void; viewLabel?: string;
}) {
  const preserveSelection = (event: MouseEvent<HTMLButtonElement>) => {
    const field = document.activeElement;
    if (field instanceof HTMLInputElement && field.closest('.android-editor-ribbon')) field.blur();
    event.preventDefault();
  };
  return <div className="android-mode-switch" role="group" aria-label="문서 모드">
    <button type="button" aria-label="편집 모드" title="편집 모드 · 뷰어에서 Esc" aria-pressed={editing}
      disabled={disabled || editDisabled || !onEdit} onMouseDown={preserveSelection} onClick={onEdit}>
      <Pencil size={19} aria-hidden="true" /><span>편집</span>
    </button>
    <button type="button" aria-label={viewLabel} title="뷰어 모드" aria-pressed={!editing}
      disabled={disabled} onMouseDown={preserveSelection} onClick={onView}>
      <BookOpen size={19} aria-hidden="true" /><span>뷰어</span>
    </button>
  </div>;
}
