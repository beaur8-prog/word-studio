import { useEffect, useRef, useState } from "react";
import type { CSSProperties, ReactNode } from "react";
const escapeHtml = (text: string) => { const el = document.createElement('span'); el.textContent = text; return el.innerHTML; };
import { EditorContent, useEditor } from "@tiptap/react";
import { generateJSON } from "@tiptap/core";
import { EditorState } from '@tiptap/pm/state';
import {
  Save,
  Undo2,
  Redo2,
  FileText,
  FolderOpen,
  Search,
  ChevronDown,
  ChevronUp,
  ClipboardPaste,
  Scissors,
  Copy,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Subscript,
  Superscript,
  Highlighter,
  Type,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  List,
  ListOrdered,
  IndentIncrease,
  IndentDecrease,
  RemoveFormatting,
  Table2,
  Image as ImageIcon,
  Link,
  MessageSquare,
  FilePlus2,
  FileDown,
  Printer,
  PanelLeft,
  PanelRight,
  Sun,
  Moon,
  Minus,
  Plus,
  X,
  Check,
  CheckCircle2,
  ZoomIn,
  BookOpen,
  LayoutTemplate,
  BetweenHorizontalStart,
  Columns2,
  RectangleHorizontal,
  RectangleVertical,
  Paintbrush,
  TextCursorInput,
  ArrowUp,
  ArrowDown,
  Trash2,
  Heading1,
  CalendarDays,
} from "lucide-react";
import { editorExtensions, parseDocumentHtml } from "./extensions";
import { TableContextMenu } from './TableContextMenu';
import { defaultPage, pageMargins, pageFurniture } from "./types";
import { readHistory, readPreviousVersions, saveRecovery, recoveryNotice } from './recovery';
import { createPrintLayout, type PrintLayout } from './print-layout';
import type { StudioDocument, PageSettings, Comment } from "./types";
import {
  exportDocx,
  importFile,
  validateDocument,
  sanitizeImportedHtml,
  DOCX_EXPORT_NOTICE,
} from "./document-io";

const sample = `<p style="color:#526d92;font-size:10pt">STUDIO / WORKING DOCUMENT</p><h1>아이디어를 기록하고,<br>다음 단계를 만드세요.</h1><p>생각의 시작부터 완성된 문서까지. Word Studio에서 글을 쓰고, 내용을 정리하고, 나만의 문서를 만들어 보세요.</p><h2>01. 문서 작성 시작하기</h2><p>이 문장은 자유롭게 수정할 수 있습니다. 텍스트를 선택한 다음 <strong>홈</strong> 탭에서 글꼴과 크기, 색상을 바꿔 보세요. 제목 스타일을 적용하면 왼쪽 탐색 창에 문서의 구조가 표시됩니다.</p><h2>02. 한눈에 보는 작업 계획</h2><table><tr><th><p>단계</p></th><th><p>주요 작업</p></th><th><p>상태</p></th></tr><tr><td><p>기획</p></td><td><p>핵심 아이디어와 목표 정리</p></td><td><p>진행 중</p></td></tr><tr><td><p>작성</p></td><td><p>자료 수집과 초안 작성</p></td><td><p>예정</p></td></tr><tr><td><p>검토</p></td><td><p>내용 검토와 최종 문서 공유</p></td><td><p>예정</p></td></tr></table><h2>03. 작은 체크리스트</h2><ul><li><p>읽는 사람을 떠올리며 목적을 분명하게 씁니다.</p></li><li><p>표와 이미지를 추가해 내용을 더 쉽게 전달합니다.</p></li><li><p>완성된 문서는 DOCX 또는 PDF로 내보냅니다.</p></li></ul><p><em>좋은 문서는 작은 메모에서 시작됩니다.</em></p>`;
const blank = (title = "문서1", html = "<p></p>"): StudioDocument => ({
  id: crypto.randomUUID(),
  format: "word-studio",
  version: 1,
  title,
  content: generateJSON(html, editorExtensions),
  page: { ...defaultPage },
  comments: [],
  updatedAt: new Date().toISOString(),
});
function readInitial() {
  if (readHistory().length) return readHistory()[0];
  try {
    const raw = localStorage.getItem("word-studio-current");
    return raw ? validateDocument(JSON.parse(raw)) : blank("시작하기", sample);
  } catch {
    return blank("시작하기", sample);
  }
}
// Timestamps change on render; compare only the editable document state.
function documentContents(doc: StudioDocument) {
  const { updatedAt, ...contents } = doc;
  return JSON.stringify(contents);
}

function SaveFailureDialog({ title, onSave, onDiscard, onCancel }: {
  title: string; onSave(): Promise<boolean>; onDiscard(): void; onCancel(): void;
}) {
  const ref = useRef<HTMLDialogElement>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  useEffect(() => { ref.current?.showModal(); }, []);
  return <dialog ref={ref} className="dialog save-failure-dialog" aria-labelledby="save-failure-title"
    onCancel={event => { event.preventDefault(); if (!saving) onCancel(); }}>
    <h2 id="save-failure-title">현재 문서를 저장하지 못했습니다</h2>
    <p>‘{title}’의 최신 변경 내용을 이 PC에 보관하지 못했습니다. 다른 문서를 열기 전에 파일로 저장할 수 있습니다.</p>
    <p className="muted">저장하지 않고 계속하면 저장되지 않은 변경 내용을 잃을 수 있습니다.</p>
    {error && <p role="alert">{error}</p>}
    <div className="dialog-actions">
      <button autoFocus disabled={saving} onClick={onCancel}>취소</button>
      <button disabled={saving} onClick={onDiscard}>저장하지 않고 계속</button>
      <button className="primary" disabled={saving} onClick={async () => {
        setSaving(true);
        setError('');
        try {
          if (!await onSave()) setError('파일 저장을 완료하지 못했습니다. 다시 저장하거나 취소를 선택하세요.');
        } finally { setSaving(false); }
      }}>{saving ? '저장 중…' : '파일로 저장 후 계속'}</button>
    </div>
  </dialog>;
}
function FontSizeInput({
  value,
  onCommit,
  onDone,
}: {
  value: number;
  onCommit: (size: number) => void;
  onDone: () => void;
}) {
  const [draft, setDraft] = useState(String(value));
  useEffect(() => setDraft(String(value)), [value]);
  const commit = () => {
    const n = Number(draft);
    if (Number.isFinite(n) && n >= 6 && n <= 96) {
      if (n !== value) onCommit(n);
    } else setDraft(String(value));
  };
  return (
    <input
      aria-label="글꼴 크기"
      type="number"
      min="6"
      max="96"
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          commit();
          onDone();
        }
        if (e.key === "Escape") {
          setDraft(String(value));
          onDone();
        }
      }}
    />
  );
}
function IconButton({
  label,
  children,
  onClick,
  active,
  disabled = false,
  className = "",
}: {
  label: string;
  children: ReactNode;
  onClick: () => void;
  active?: boolean;
  disabled?: boolean;
  className?: string;
}) {
  return (
    <button
      type="button"
      title={label}
      aria-label={label}
      aria-pressed={active}
      disabled={disabled}
      className={`icon-button ${active ? "active" : ""} ${className}`}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}
function Tool({
  label,
  icon,
  onClick,
  active = false,
}: {
  label: string;
  icon: ReactNode;
  onClick: () => void;
  active?: boolean;
}) {
  return (
    <button
      type="button"
      className={`tool ${active ? "active" : ""}`}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {icon}
      <span>{label}</span>
    </button>
  );
}
function Group({
  label,
  children,
  className = "",
}: {
  label: string;
  children: ReactNode;
  className?: string;
}) {
  return (
    <section className={`ribbon-group ${className}`}>
      <div className="group-content">{children}</div>
      <span className="group-label">{label}</span>
    </section>
  );
}
function Dropdown({
  label,
  icon,
  children,
  wide = false,
}: {
  label: string;
  icon?: ReactNode;
  children: ReactNode;
  wide?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const close = (e: PointerEvent) => {
      if (!ref.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("pointerdown", close);
    return () => document.removeEventListener("pointerdown", close);
  }, []);
  return (
    <div className={`dropdown ${wide ? "wide" : ""}`} ref={ref}>
      <button
        className={icon ? "tool dropdown-tool" : "dropdown-button"}
        aria-expanded={open}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen(!open)}
      >
        {icon}
        <span>{label}</span>
        <ChevronDown size={11} />
      </button>
      {open && (
        <div
          className="popover"
          onClick={(e) => {
            if ((e.target as HTMLElement).closest("button")) setOpen(false);
          }}
        >
          {children}
        </div>
      )}
    </div>
  );
}
function MenuItem({
  children,
  onClick,
}: {
  children: ReactNode;
  onClick: () => void;
}) {
  return (
    <button
      className="menu-item"
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  );
}
function download(blob: Blob, name: string) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

export default function App() {
  const [initial] = useState(readInitial),
    [documentId, setDocumentId] = useState(initial.id || crypto.randomUUID()),
    [title, setTitle] = useState(initial.title),
    [page, setPage] = useState<PageSettings>(initial.page),
    [comments, setComments] = useState<Comment[]>(initial.comments),
    [tab, setTab] = useState("홈"),
    [revision, setRevision] = useState(0),
    [selectionVersion, setSelectionVersion] = useState(0),
    [theme, setTheme] = useState<"dark" | "light">(() =>
      localStorage.getItem("word-studio-theme") === "light" ? "light" : "dark",
    ),
    [darkPaper, setDarkPaper] = useState(false),
    [zoom, setZoom] = useState(100),
    [pages, setPages] = useState(1),
    [currentPage, setCurrentPage] = useState(1),
    [ruler, setRuler] = useState(true),
    [navigation, setNavigation] = useState(false),
    [navTab, setNavTab] = useState("제목"),
    [commentsOpen, setCommentsOpen] = useState(false),
    [query, setQuery] = useState(""),
    [replace, setReplace] = useState(""),
    [findIndex, setFindIndex] = useState(0),
    [readOnly, setReadOnly] = useState(false),
    [marks, setMarks] = useState(false),
    [saved, setSaved] = useState("이 PC에 저장됨"),
    [toast, setToast] = useState(""),
    [modal, setModal] = useState<
      "new" | "link" | "header" | "comment" | "stats" | "about" | null
    >(null),
    [modalValue, setModalValue] = useState(""),
    [commentQuote, setCommentQuote] = useState(""),
    [fileMenu, setFileMenu] = useState(false),
    [command, setCommand] = useState(false),
    [commandQuery, setCommandQuery] = useState(""),
    [collapsed, setCollapsed] = useState(false),
    [busy, setBusy] = useState(false),
    [tableRows, setTableRows] = useState(3),
    [tableCols, setTableCols] = useState(3);
  const [printPreview, setPrintPreview] = useState<PrintLayout | null>(null);
  const [failedTransition, setFailedTransition] = useState<StudioDocument | null>(null);
  const transitioning = useRef(false);
  const outputLayout = useRef<PrintLayout | null>(null);
  const fileInput = useRef<HTMLInputElement>(null),
    imageInput = useRef<HTMLInputElement>(null),
    stackRef = useRef<HTMLDivElement>(null),
    workspaceRef = useRef<HTMLDivElement>(null),
    currentRef = useRef<StudioDocument>(initial),
    toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const pendingRecovery = useRef(false), recoveryEpoch = useRef(0);
  const commentRange = useRef<{ from: number; to: number } | null>(null);
  const editor = useEditor({
    extensions: editorExtensions,
    content: initial.content,
    onUpdate() {
      pendingRecovery.current = true;
      setRevision((x) => x + 1);
      setSaved("저장 중…");
    },
    onSelectionUpdate() {
      setSelectionVersion((x) => x + 1);
    },
    onTransaction() {
      setSelectionVersion((x) => x + 1);
    },
    editorProps: {
      attributes: { "aria-label": "문서 편집 영역", spellcheck: "false" },
      transformPastedHTML: (html) => sanitizeImportedHtml(html).html,
    },
  });
  const notify = (message: string) => {
    setToast(message);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(""), 4500);
  };
  currentRef.current = {
    id: documentId,
    format: "word-studio",
    version: 1,
    title,
    content: editor?.getJSON() ?? initial.content,
    page,
    comments,
    updatedAt: new Date().toISOString(),
  };
  useEffect(() => {
    pendingRecovery.current = true;
    setSaved('저장 중…');
    const epoch = ++recoveryEpoch.current;
    const snapshot = structuredClone(currentRef.current);
    const id = setTimeout(() => {
      void saveRecovery(snapshot).then(() => {
        if (epoch === recoveryEpoch.current) {
          pendingRecovery.current = false;
          setSaved("이 PC에 저장됨");
        }
      }).catch(() => {
        if (epoch === recoveryEpoch.current) setSaved("자동 저장 실패 · 파일로 저장하세요");
      });
    }, 600);
    return () => clearTimeout(id);
  }, [documentId, title, page, comments, revision]);
  useEffect(() => {
    const persist = (event: BeforeUnloadEvent) => {
      if (pendingRecovery.current) {
        event.preventDefault();
        event.returnValue = '';
        void saveRecovery(structuredClone(currentRef.current)).catch(() => {});
      }
    };
    window.addEventListener("beforeunload", persist);
    return () => window.removeEventListener("beforeunload", persist);
  }, []);
  useEffect(() => { if (recoveryNotice) notify(recoveryNotice); }, []);
  useEffect(() => {
    document.documentElement.dataset.chrome = theme;
    localStorage.setItem("word-studio-theme", theme);
    window.wordDesktop?.setTheme(theme);
  }, [theme]);
  useEffect(() => {
    document.title = `${title} — Word Studio`;
  }, [title]);
  useEffect(() => {
    editor?.setEditable(!readOnly);
  }, [editor, readOnly]);
  useEffect(() => {
    const el = stackRef.current;
    const change = (e: Event) => setPages((e as CustomEvent).detail);
    el?.addEventListener("pagecount", change);
    return () => el?.removeEventListener("pagecount", change);
  }, [editor]);
  useEffect(() => {
    const stack = stackRef.current; if (!stack || !editor) return;
    let frame = 0;
    const update = () => {
      cancelAnimationFrame(frame); frame = requestAnimationFrame(() => {
        const margins = pageMargins(page), px = 96 / 25.4;
        const scale = stack.getBoundingClientRect().width / stack.offsetWidth || 1;
        const occupied = (selector: string, distance: number) => Math.max(0, ...[...stack.querySelectorAll<HTMLElement>(selector)].map(el =>
          el.textContent?.trim() || el.querySelector('img,table') ? distance * px + el.getBoundingClientRect().height / scale + 2 : 0));
        const top = Math.max(margins.top * px, occupied('.paper-header', page.headerDistance ?? Math.max(0, margins.top / 2 - 3)));
        const bottom = Math.max(margins.bottom * px, occupied('.paper-footer', page.footerDistance ?? Math.max(0, margins.bottom / 2 - 3)));
        stack.style.setProperty('--page-padding', `${top}px ${margins.right * px}px ${bottom}px ${margins.left * px}px`);
        stack.dataset.margin = String(top); stack.dataset.bottomMargin = String(bottom);
        for (const field of stack.querySelectorAll<HTMLElement>('.paper [data-page-number]')) {
          const index = [...stack.querySelectorAll('.paper')].indexOf(field.closest('.paper')!);
          field.textContent = String(field.dataset.pageNumber === 'NUMPAGES' ? pages : index + 1);
        }
      });
    };
    const observer = new ResizeObserver(update);
    stack.querySelectorAll('.paper-header,.paper-footer').forEach(el => observer.observe(el));
    update(); document.fonts.addEventListener('loadingdone', update);
    return () => { cancelAnimationFrame(frame); observer.disconnect(); document.fonts.removeEventListener('loadingdone', update); };
  }, [editor, page, pages]);
  useEffect(() => {
    if (editor)
      requestAnimationFrame(() =>
        editor.view.dispatch(editor.state.tr.setMeta("addToHistory", false)),
      );
  }, [page, editor]);
  useEffect(() => () => clearTimeout(toastTimer.current), []);
  const activateDocument = (doc: StudioDocument) => {
    if (!editor) return;
    // A new EditorState resets history and selection plugins together. Merely
    // excluding setContent from history would leave the old undo/redo branches.
    editor.view.updateState(EditorState.create({
      schema: editor.schema, doc: editor.schema.nodeFromJSON(doc.content), plugins: editor.state.plugins,
    }));
    ++recoveryEpoch.current;
    pendingRecovery.current = true;
    outputLayout.current?.dispose(); outputLayout.current = null;
    document.body.classList.remove('studio-printing'); setPrintPreview(null);
    setFailedTransition(null);
    transitioning.current = false;
    setDocumentId(doc.id || crypto.randomUUID());
    setTitle(doc.title);
    setPage(doc.page);
    setComments(doc.comments);
    setRevision(value => value + 1);
    setModal(null);
    setFileMenu(false);
    setQuery("");
    setReadOnly(false);
    setPages(1);
    setCurrentPage(1);
    workspaceRef.current?.scrollTo({ top: 0 });
    notify("문서를 열었습니다");
  };
  const loadDoc = async (doc: StudioDocument, html?: string) => {
    if (!editor || transitioning.current) return false;
    const next = validateDocument({ ...doc, content: html === undefined ? doc.content : parseDocumentHtml(html) });
    // Parse before replacing anything, including the current undo history.
    editor.schema.nodeFromJSON(next.content).check();
    transitioning.current = true;
    try {
      // Users can keep editing while disk I/O is pending. Flush a fresh snapshot
      // if the document changed during the write, then switch without an await.
      let snapshot: StudioDocument;
      do {
        snapshot = structuredClone(currentRef.current);
        await saveRecovery(snapshot);
      } while (documentContents(snapshot) !== documentContents(currentRef.current));
      activateDocument(next);
      return true;
    } catch {
      pendingRecovery.current = true;
      setSaved('자동 저장 실패 · 파일로 저장하세요');
      setModal(null); setFileMenu(false); setCommand(false);
      setFailedTransition(next);
      return false;
    }
  };
  const readFile = async (file: File) => {
    setBusy(true);
    try {
      const result = await importFile(file);
      const loaded = await loadDoc(result.document, result.html);
      if (loaded && result.warnings.length) notify(result.warnings.join(" "));
    } catch (e) {
      notify(
        `파일을 열지 못했습니다: ${e instanceof Error ? e.message : String(e)}`,
      );
    } finally {
      setBusy(false);
    }
  };
  const open = async () => {
    setFileMenu(false);
    if (window.wordDesktop) {
      try {
        const f = await window.wordDesktop.openFile();
        if (f) await readFile(new File([new Uint8Array(f.data)], f.name));
      } catch (e) {
        notify(String(e));
      }
    } else fileInput.current?.click();
  };
  const save = async (
    extension: "wordstudio" | "docx" | "html" | "txt" = "wordstudio",
  ): Promise<boolean> => {
    if (!editor) return false;
    setBusy(true);
    setFileMenu(false);
    try {
      const doc = currentRef.current;
      const blob =
        extension === "docx"
          ? await exportDocx(doc)
          : extension === "wordstudio"
            ? new Blob([JSON.stringify(doc, null, 2)], {
                type: "application/json",
              })
            : extension === "txt"
              ? new Blob([editor.getText()], {
                  type: "text/plain;charset=utf-8",
                })
              : new Blob(
                  [
                    `<!doctype html><html lang="ko"><meta charset="utf-8"><title>${title.replace(/[<>&]/g, "")}</title><body>${editor.getHTML()}</body></html>`,
                  ],
                  { type: "text/html;charset=utf-8" },
                );
      const name = `${title.replace(/[<>:"/\\|?*]/g, "-") || "문서"}.${extension}`;
      if (window.wordDesktop) {
        if (
          !(await window.wordDesktop.saveFile({
            name,
            extension,
            data: new Uint8Array(await blob.arrayBuffer()),
          }))
        )
          return false;
      } else download(blob, name);
      notify(
        `${extension === "wordstudio" ? "문서" : extension.toUpperCase()} 파일을 저장했습니다` +
          (extension === "docx" && comments.length
            ? ` · ${DOCX_EXPORT_NOTICE}`
            : ""),
      );
      return true;
    } catch (e) {
      notify(`저장 실패: ${e instanceof Error ? e.message : String(e)}`);
      return false;
    } finally {
      setBusy(false);
    }
  };
  const preview = async () => {
    outputLayout.current?.dispose();
    const layout = await createPrintLayout(editor!.getHTML(), currentRef.current.page);
    outputLayout.current = layout;
    layout.element.classList.replace('measuring', 'preview');
    document.body.classList.add('studio-printing');
    setPrintPreview(layout);
    setFileMenu(false);
    return layout;
  };
  const closePreview = () => {
    outputLayout.current?.dispose(); outputLayout.current = null;
    document.body.classList.remove('studio-printing'); setPrintPreview(null);
  };
  const print = async (pdf = false) => {
    setFileMenu(false);
    try {
      if (!outputLayout.current) await preview();
      if (window.wordDesktop) {
        if (pdf) {
          if (
            await window.wordDesktop.exportPdf({
              name: `${title}.pdf`,
              header: page.header,
              footer: page.footer,
              pageNumbers: page.pageNumbers,
              margin: page.margin,
              paginated: true,
            })
          )
            notify("PDF를 저장했습니다");
        } else await window.wordDesktop.print();
      } else window.print();
    } catch (e) {
      notify(`출력 실패: ${String(e)}`);
    }
  };
  const addComment = () => {
    const s = editor?.state.selection;
    commentRange.current = s ? { from: s.from, to: s.to } : null;
    setCommentQuote(s ? editor!.state.doc.textBetween(s.from, s.to, " ") : "");
    setModalValue("");
    setModal("comment");
  };
  const actions = [
    { name: "새 문서", hint: "Ctrl+N", action: () => setModal("new") },
    { name: "파일 열기", hint: "Ctrl+O", action: open },
    { name: "문서 저장", hint: "Ctrl+S", action: () => save() },
    { name: "Word 문서로 내보내기", hint: "DOCX", action: () => save("docx") },
    { name: "PDF로 내보내기", hint: "PDF", action: () => print(true) },
    { name: '인쇄 미리보기', hint: '페이지 확인', action: () => { void preview().catch(e => notify(String(e))); } },
    { name: "인쇄", hint: "Ctrl+P", action: () => print() },
    {
      name: "찾기 및 바꾸기",
      hint: "Ctrl+F",
      action: () => {
        setNavigation(true);
        setNavTab("검색");
      },
    },
    { name: "새 메모", hint: "검토", action: addComment },
    { name: "단어 개수", hint: "통계", action: () => setModal("stats") },
    {
      name: "화면 테마 전환",
      hint: theme === "dark" ? "밝게" : "어둡게",
      action: () => setTheme(theme === "dark" ? "light" : "dark"),
    },
  ];
  const actionRef = useRef(actions);
  actionRef.current = actions;
  useEffect(() => {
    const key = (e: KeyboardEvent) => {
      if (e.isComposing || e.keyCode === 229) return;
      if (document.querySelector('.save-failure-dialog[open]')) return;
      if (e.key === "Escape") {
        setModal(null);
        setCommand(false);
        setFileMenu(false);
        return;
      }
      if (!(e.ctrlKey || e.metaKey)) return;
      const map: Record<string, string> = {
        s: "문서 저장",
        o: "파일 열기",
        n: "새 문서",
        p: "인쇄",
        f: "찾기 및 바꾸기",
      };
      if (map[e.key.toLowerCase()]) {
        e.preventDefault();
        actionRef.current
          .find((a) => a.name === map[e.key.toLowerCase()])
          ?.action();
      }
      if (e.key.toLowerCase() === "k" && e.shiftKey) {
        e.preventDefault();
        setCommand(true);
      }
    };
    window.addEventListener("keydown", key);
    const off = window.wordDesktop?.onMenu((a) => {
      if (document.querySelector('.save-failure-dialog[open]')) return;
      const map: Record<string, string> = {
        new: "새 문서",
        open: "파일 열기",
        save: "문서 저장",
        "save-as": "문서 저장",
        "export-docx": "Word 문서로 내보내기",
        "export-pdf": "PDF로 내보내기",
        print: "인쇄",
      };
      actionRef.current.find((x) => x.name === map[a])?.action();
    });
    return () => {
      window.removeEventListener("keydown", key);
      off?.();
    };
  }, []);
  if (!editor) return <div className="loading">Word Studio를 여는 중…</div>;
  const text = editor.getText(),
    words = text.trim() ? text.trim().split(/\s+/).length : 0;
  const headings: { title: string; level: number; pos: number }[] = [];
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === "heading")
      headings.push({ title: node.textContent, level: node.attrs.level, pos });
  });
  const matches: { from: number; to: number; preview: string }[] = [];
  if (query) {
    editor.state.doc.descendants((node, pos) => {
      if (node.isTextblock) {
        const value = node.textBetween(0, node.content.size, "\n", "\ufffc");
        let index = 0;
        while (
          (index = value
            .toLocaleLowerCase()
            .indexOf(query.toLocaleLowerCase(), index)) >= 0
        ) {
          matches.push({
            from: pos + 1 + index,
            to: pos + 1 + index + query.length,
            preview: value,
          });
          index += Math.max(1, query.length);
        }
        return false;
      }
    });
  }
  const selectMatch = (i: number) => {
    if (!matches.length) return;
    const index = (i + matches.length) % matches.length;
    setFindIndex(index);
    editor
      .chain()
      .focus()
      .setTextSelection(matches[index])
      .scrollIntoView()
      .run();
  };
  const replaceMatch = (all = false) => {
    if (!matches.length) return;
    const tr = editor.state.tr;
    const targets = all
      ? [...matches].reverse()
      : [matches[Math.min(findIndex, matches.length - 1)]];
    for (const m of targets) tr.insertText(replace, m.from, m.to);
    editor.view.dispatch(tr);
    notify(`${targets.length}개 항목을 바꿨습니다`);
  };
  const setIndent = (delta: number) => {
    const type = editor.isActive("heading") ? "heading" : "paragraph";
    const current = editor.getAttributes(type).indent || 0;
    editor
      .chain()
      .focus()
      .updateAttributes(type, {
        indent: Math.max(0, Math.min(240, current + delta)),
      })
      .run();
  };
  const insertImage = async (file: File) => {
    if (!/^image\/(png|jpeg|gif|webp)$/.test(file.type)) {
      notify("PNG, JPEG, GIF 또는 WebP 이미지를 선택하세요");
      return;
    }
    if (file.size > 15 * 1024 * 1024) {
      notify("이미지는 15MB 이하로 선택하세요");
      return;
    }
    const reader = new FileReader();
    reader.onload = () =>
      editor
        .chain()
        .focus()
        .setImage({ src: String(reader.result), alt: file.name })
        .run();
    reader.readAsDataURL(file);
  };
  const dims = page.size === "A4" ? [793.7, 1122.52] : [816, 1056],
    pageWidth = page.orientation === "portrait" ? dims[0] : dims[1],
    pageHeight = page.orientation === "portrait" ? dims[1] : dims[0],
    margin = (pageMargins(page).top * 96) / 25.4;
  const setPageValue = (change: Partial<PageSettings>) =>
    setPage((p) => ({ ...p, ...change, ...(change.margin !== undefined ? { margins: undefined } : {}),
      ...Object.fromEntries(['header','footer','firstHeader','firstFooter','evenHeader','evenFooter'].filter(key => key in change).map(key => [`${key}Html`, undefined])) }));
  const format = editor.getAttributes("textStyle");
  void selectionVersion;
  const fit = () => {
    if (workspaceRef.current)
      setZoom(
        Math.max(
          40,
          Math.min(
            150,
            Math.floor(
              ((workspaceRef.current.clientWidth - 90) / pageWidth) * 100,
            ),
          ),
        ),
      );
  };
  const history = readHistory();
  const margins = pageMargins(page);
  return (
    <div
      className={`office-shell ${window.wordDesktop ? "desktop-windows" : ""} ${readOnly ? "read-mode" : ""}`}
    >
      <style>{`@page {size:${page.size} ${page.orientation};margin:${margins.top}mm ${margins.right}mm ${margins.bottom}mm ${margins.left}mm;}`}</style>
      <header className="topbar">
        <div className="app-brand">
          <span className="brand-mark">
            <FileText size={17} />
          </span>
          <span>
            Word <b>Studio</b>
          </span>
        </div>
        <div className="quick-actions">
          <IconButton label="저장 (Ctrl+S)" onClick={() => save()}>
            <Save size={16} />
          </IconButton>
          <IconButton
            label="실행 취소 (Ctrl+Z)"
            disabled={readOnly || !editor.can().undo()}
            onClick={() => editor.chain().focus().undo().run()}
          >
            <Undo2 size={16} />
          </IconButton>
          <IconButton
            label="다시 실행 (Ctrl+Y)"
            disabled={readOnly || !editor.can().redo()}
            onClick={() => editor.chain().focus().redo().run()}
          >
            <Redo2 size={16} />
          </IconButton>
        </div>
        <div className="document-title">
          <input
            aria-label="문서 제목"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
          />
          <span
            className={
              saved.includes("실패") ? "save-indicator error" : "save-indicator"
            }
          >
            <CheckCircle2 size={12} />
            {saved}
          </span>
        </div>
        <button className="command-trigger" onClick={() => setCommand(true)}>
          <Search size={14} />
          <span>어떤 작업을 원하시나요?</span>
          <kbd>Ctrl Shift K</kbd>
        </button>
        <div className="suite-badge">STUDIO OFFICE</div>
        <IconButton
          label="화면 테마 전환"
          onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
        >
          {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
        </IconButton>
      </header>
      <nav className="tabs" aria-label="리본 메뉴">
        <button
          className={fileMenu ? "selected-tab" : ""}
          onClick={() => setFileMenu(!fileMenu)}
        >
          파일
        </button>
        {[
          "홈",
          "삽입",
          "디자인",
          "레이아웃",
          "검토",
          "보기",
          ...(editor.isActive("table") ? ["표 레이아웃"] : []),
        ].map((t) => (
          <button
            key={t}
            className={tab === t ? "selected-tab" : ""}
            onClick={() => setTab(t)}
          >
            {t}
          </button>
        ))}
        <span className="tab-spacer" />
        <button
          className={commentsOpen ? "tab-action selected-tab" : "tab-action"}
          onClick={() => setCommentsOpen(!commentsOpen)}
        >
          <MessageSquare size={14} />
          메모 {comments.filter((c) => !c.resolved).length || ""}
        </button>
        <Dropdown label={readOnly ? "읽기 모드" : "편집 모드"}>
          <MenuItem onClick={() => setReadOnly(false)}>
            편집 · 문서 수정
          </MenuItem>
          <MenuItem onClick={() => setReadOnly(true)}>
            읽기 · 문서 살펴보기
          </MenuItem>
        </Dropdown>
        <button className="export-top" onClick={() => save("docx")}>
          <FileDown size={14} />
          내보내기
        </button>
      </nav>
      {!collapsed && (
        <div
          className="ribbon"
          role="toolbar"
          aria-label={`${tab} 도구`}
          inert={readOnly && !["보기", "검토"].includes(tab)}
        >
          {tab === "홈" && (
            <>
              <Group label="클립보드">
                <Tool
                  label="붙여넣기"
                  icon={<ClipboardPaste size={26} />}
                  onClick={async () => {
                    try {
                      editor
                        .chain()
                        .focus()
                        .insertContent(
                          (await navigator.clipboard.readText())
                            .split("\n")
                            .map((t) => ({
                              type: "paragraph",
                              content: t ? [{ type: "text", text: t }] : [],
                            })),
                        )
                        .run();
                    } catch {
                      notify("문서에서 Ctrl+V로 붙여넣으세요");
                    }
                  }}
                />
                <div className="tools-stack">
                  <button
                    onClick={() => {
                      editor.commands.focus();
                      document.execCommand("cut");
                    }}
                  >
                    <Scissors />
                    잘라내기
                  </button>
                  <button
                    onClick={() => {
                      editor.commands.focus();
                      document.execCommand("copy");
                      notify("선택한 내용을 복사했습니다");
                    }}
                  >
                    <Copy />
                    복사
                  </button>
                  <button
                    onClick={() =>
                      editor.chain().focus().unsetAllMarks().clearNodes().run()
                    }
                  >
                    <RemoveFormatting />
                    서식 지우기
                  </button>
                </div>
              </Group>
              <Group label="글꼴" className="font-group">
                <div className="tool-rows">
                  <div className="font-row">
                    <select
                      aria-label="글꼴"
                      value={format.fontFamily || "Malgun Gothic"}
                      onChange={(e) =>
                        editor
                          .chain()
                          .focus()
                          .setFontFamily(e.target.value)
                          .run()
                      }
                    >
                      {[
                        "Malgun Gothic",
                        "Arial",
                        "Calibri",
                        "Aptos",
                        "Georgia",
                        "Times New Roman",
                        "Batang",
                        "Gulim",
                        "Courier New",
                      ].map((f) => (
                        <option key={f} value={f}>
                          {f === "Malgun Gothic"
                            ? "맑은 고딕"
                            : f === "Batang"
                              ? "바탕"
                              : f === "Gulim"
                                ? "굴림"
                                : f}
                        </option>
                      ))}
                    </select>
                    <FontSizeInput
                      value={parseFloat(format.fontSize || "11")}
                      onCommit={(n) =>
                        editor.chain().setFontSize(`${n}pt`).run()
                      }
                      onDone={() => editor.commands.focus()}
                    />
                    <IconButton
                      label="글꼴 크기 늘림"
                      onClick={() =>
                        editor
                          .chain()
                          .focus()
                          .setFontSize(
                            `${Math.min(96, parseFloat(format.fontSize || "11") + 1)}pt`,
                          )
                          .run()
                      }
                    >
                      <span>
                        A<sup>+</sup>
                      </span>
                    </IconButton>
                    <IconButton
                      label="글꼴 크기 줄임"
                      onClick={() =>
                        editor
                          .chain()
                          .focus()
                          .setFontSize(
                            `${Math.max(6, parseFloat(format.fontSize || "11") - 1)}pt`,
                          )
                          .run()
                      }
                    >
                      <span>
                        A<sup>−</sup>
                      </span>
                    </IconButton>
                  </div>
                  <div className="button-row">
                    <IconButton
                      label="굵게 (Ctrl+B)"
                      active={editor.isActive("bold")}
                      onClick={() => editor.chain().focus().toggleBold().run()}
                    >
                      <Bold />
                    </IconButton>
                    <IconButton
                      label="기울임꼴 (Ctrl+I)"
                      active={editor.isActive("italic")}
                      onClick={() =>
                        editor.chain().focus().toggleItalic().run()
                      }
                    >
                      <Italic />
                    </IconButton>
                    <IconButton
                      label="밑줄 (Ctrl+U)"
                      active={editor.isActive("underline")}
                      onClick={() =>
                        editor.chain().focus().toggleUnderline().run()
                      }
                    >
                      <Underline />
                    </IconButton>
                    <IconButton
                      label="취소선"
                      active={editor.isActive("strike")}
                      onClick={() =>
                        editor.chain().focus().toggleStrike().run()
                      }
                    >
                      <Strikethrough />
                    </IconButton>
                    <IconButton
                      label="아래 첨자"
                      active={editor.isActive("subscript")}
                      onClick={() =>
                        editor.chain().focus().toggleSubscript().run()
                      }
                    >
                      <Subscript />
                    </IconButton>
                    <IconButton
                      label="위 첨자"
                      active={editor.isActive("superscript")}
                      onClick={() =>
                        editor.chain().focus().toggleSuperscript().run()
                      }
                    >
                      <Superscript />
                    </IconButton>
                    <label className="color-tool" title="텍스트 강조 색">
                      <Highlighter />
                      <input
                        aria-label="텍스트 강조 색"
                        type="color"
                        value="#fff29a"
                        onChange={(e) =>
                          editor
                            .chain()
                            .focus()
                            .setHighlight({ color: e.target.value })
                            .run()
                        }
                      />
                      <i style={{ background: "#fff29a" }} />
                    </label>
                    <label className="color-tool" title="글꼴 색">
                      <Type />
                      <input
                        aria-label="글꼴 색"
                        type="color"
                        value={format.color || "#1d2430"}
                        onChange={(e) =>
                          editor.chain().focus().setColor(e.target.value).run()
                        }
                      />
                      <i style={{ background: format.color || "#db4b4b" }} />
                    </label>
                  </div>
                </div>
              </Group>
              <Group label="단락">
                <div className="tool-rows">
                  <div className="button-row">
                    <IconButton
                      label="글머리 기호"
                      active={editor.isActive("bulletList")}
                      onClick={() =>
                        editor.chain().focus().toggleBulletList().run()
                      }
                    >
                      <List />
                    </IconButton>
                    <IconButton
                      label="번호 매기기"
                      active={editor.isActive("orderedList")}
                      onClick={() =>
                        editor.chain().focus().toggleOrderedList().run()
                      }
                    >
                      <ListOrdered />
                    </IconButton>
                    <IconButton label="내어쓰기" onClick={() => setIndent(-24)}>
                      <IndentDecrease />
                    </IconButton>
                    <IconButton label="들여쓰기" onClick={() => setIndent(24)}>
                      <IndentIncrease />
                    </IconButton>
                    <IconButton
                      label="편집 기호 표시"
                      active={marks}
                      onClick={() => setMarks(!marks)}
                    >
                      <span className="pilcrow">¶</span>
                    </IconButton>
                  </div>
                  <div className="button-row">
                    {[
                      ["left", AlignLeft, "왼쪽 맞춤"],
                      ["center", AlignCenter, "가운데 맞춤"],
                      ["right", AlignRight, "오른쪽 맞춤"],
                      ["justify", AlignJustify, "양쪽 맞춤"],
                    ].map(([align, Icon, label]) => {
                      const I = Icon as typeof AlignLeft;
                      return (
                        <IconButton
                          key={String(align)}
                          label={String(label)}
                          active={editor.isActive({ textAlign: align })}
                          onClick={() =>
                            editor
                              .chain()
                              .focus()
                              .setTextAlign(String(align))
                              .run()
                          }
                        >
                          <I />
                        </IconButton>
                      );
                    })}
                    <Dropdown label="줄 간격">
                      {["1", "1.15", "1.5", "2", "2.5", "3"].map((n) => (
                        <MenuItem
                          key={n}
                          onClick={() =>
                            editor
                              .chain()
                              .focus()
                              .updateAttributes(
                                editor.isActive("heading")
                                  ? "heading"
                                  : "paragraph",
                                { lineSpacing: n, wordLineRule: editor.getAttributes(editor.isActive('heading') ? 'heading' : 'paragraph').wordLineRule ? 'auto' : null },
                              )
                              .run()
                          }
                        >
                          {n}
                        </MenuItem>
                      ))}
                    </Dropdown>
                  </div>
                </div>
              </Group>
              <Group label="스타일" className="styles-group">
                <div className="style-gallery">
                  {[
                    { name: "표준", level: 0 },
                    { name: "제목 1", level: 1 },
                    { name: "제목 2", level: 2 },
                    { name: "제목 3", level: 3 },
                  ].map((s) => (
                    <button
                      key={s.name}
                      className={`style-choice ${s.level === 0 ? (editor.isActive("paragraph") ? "chosen" : "") : editor.isActive("heading", { level: s.level }) ? "chosen" : ""}`}
                      onMouseDown={(e) => e.preventDefault()}
                      onClick={() =>
                        s.level
                          ? editor
                              .chain()
                              .focus()
                              .setHeading({ level: s.level as 1 | 2 | 3 })
                              .run()
                          : editor.chain().focus().setParagraph().run()
                      }
                    >
                      <span
                        style={{
                          fontSize:
                            s.level === 1 ? 22 : s.level === 2 ? 18 : 14,
                          color: s.level ? "var(--style-blue)" : undefined,
                        }}
                      >
                        가나다Aa
                      </span>
                      <small>{s.name}</small>
                    </button>
                  ))}
                </div>
              </Group>
              <Group label="편집">
                <div className="tools-stack">
                  <button
                    onClick={() => {
                      setNavigation(true);
                      setNavTab("검색");
                    }}
                  >
                    <Search />
                    찾기
                  </button>
                  <button
                    onClick={() => {
                      setNavigation(true);
                      setNavTab("바꾸기");
                    }}
                  >
                    <TextCursorInput />
                    바꾸기
                  </button>
                  <button
                    onClick={() => editor.chain().focus().selectAll().run()}
                  >
                    <Check />
                    모두 선택
                  </button>
                </div>
              </Group>
            </>
          )}
          {tab === "삽입" && (
            <>
              <Group label="페이지">
                <Tool
                  label="새 페이지"
                  icon={<FilePlus2 />}
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .insertContent([
                        { type: "pageBreak" },
                        { type: "paragraph" },
                      ])
                      .run()
                  }
                />
                <Tool
                  label="페이지 나누기"
                  icon={<BetweenHorizontalStart />}
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .insertContent({ type: "pageBreak" })
                      .run()
                  }
                />
              </Group>
              <Group label="표">
                <Dropdown label="표" icon={<Table2 size={28} />} wide>
                  <h4>표 삽입</h4>
                  <div className="table-grid">
                    {Array.from({ length: 48 }, (_, i) => (
                      <button
                        key={i}
                        aria-label={`${Math.floor(i / 8) + 1}행 ${(i % 8) + 1}열 표`}
                        className={
                          Math.floor(i / 8) < tableRows && i % 8 < tableCols
                            ? "picked"
                            : ""
                        }
                        onMouseEnter={() => {
                          setTableRows(Math.floor(i / 8) + 1);
                          setTableCols((i % 8) + 1);
                        }}
                        onClick={() =>
                          editor
                            .chain()
                            .focus()
                            .insertTable({
                              rows: Math.floor(i / 8) + 1,
                              cols: (i % 8) + 1,
                              withHeaderRow: true,
                            })
                            .run()
                        }
                      />
                    ))}
                  </div>
                  <p className="menu-hint">
                    {tableRows}행 × {tableCols}열
                  </p>
                </Dropdown>
              </Group>
              <Group label="일러스트레이션">
                <Tool
                  label="그림"
                  icon={<ImageIcon size={28} />}
                  onClick={() => imageInput.current?.click()}
                />
                <Tool
                  label="가로줄"
                  icon={<Minus size={28} />}
                  onClick={() =>
                    editor.chain().focus().setHorizontalRule().run()
                  }
                />
              </Group>
              <Group label="링크">
                <Tool
                  label="링크"
                  icon={<Link size={26} />}
                  onClick={() => {
                    setModalValue(
                      editor.getAttributes("link").href || "https://",
                    );
                    setModal("link");
                  }}
                />
              </Group>
              <Group label="메모">
                <Tool
                  label="새 메모"
                  icon={<MessageSquare size={26} />}
                  onClick={addComment}
                />
              </Group>
              <Group label="머리글 / 바닥글">
                <Tool
                  label="머리글·바닥글"
                  icon={<LayoutTemplate size={26} />}
                  onClick={() => setModal("header")}
                />
                <Tool
                  label="페이지 번호"
                  icon={<FileText size={26} />}
                  active={page.pageNumbers}
                  onClick={() =>
                    setPageValue({ pageNumbers: !page.pageNumbers })
                  }
                />
              </Group>
              <Group label="텍스트">
                <Tool
                  label="날짜"
                  icon={<CalendarDays size={26} />}
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .insertContent(new Date().toLocaleDateString("ko-KR"))
                      .run()
                  }
                />
                <Dropdown label="기호" icon={<span className="symbol">Ω</span>}>
                  <div className="symbol-grid">
                    {[
                      "©",
                      "®",
                      "™",
                      "±",
                      "×",
                      "÷",
                      "°",
                      "→",
                      "✓",
                      "★",
                      "●",
                      "■",
                      "□",
                      "※",
                      "·",
                      "—",
                    ].map((s) => (
                      <button
                        key={s}
                        onClick={() =>
                          editor.chain().focus().insertContent(s).run()
                        }
                      >
                        {s}
                      </button>
                    ))}
                  </div>
                </Dropdown>
              </Group>
            </>
          )}
          {tab === "디자인" && (
            <>
              <Group label="문서 서식">
                <Tool
                  label="기본 서식"
                  icon={<Type size={27} />}
                  onClick={() => {
                    editor.chain().focus().selectAll().unsetAllMarks().run();
                    notify("문자 서식을 초기화했습니다");
                  }}
                />
                <Tool
                  label="클래식"
                  icon={<BookOpen size={27} />}
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .selectAll()
                      .setFontFamily("Georgia")
                      .run()
                  }
                />
                <Tool
                  label="모던"
                  icon={<Paintbrush size={27} />}
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .selectAll()
                      .setFontFamily("Malgun Gothic")
                      .run()
                  }
                />
              </Group>
              <Group label="화면">
                <Tool
                  label="밝은 테마"
                  icon={<Sun />}
                  active={theme === "light"}
                  onClick={() => setTheme("light")}
                />
                <Tool
                  label="어두운 테마"
                  icon={<Moon />}
                  active={theme === "dark"}
                  onClick={() => setTheme("dark")}
                />
                <Tool
                  label="페이지 색 전환"
                  icon={<FileText />}
                  active={darkPaper}
                  onClick={() => setDarkPaper(!darkPaper)}
                />
              </Group>
            </>
          )}
          {tab === "레이아웃" && (
            <>
              <Group label="방향별 여백 · mm"><div className="layout-fields">
                {(['top', 'right', 'bottom', 'left'] as const).map((side, i) => <label key={side}>
                  {['위', '오른쪽', '아래', '왼쪽'][i]}
                  <input aria-label={`${['위', '오른쪽', '아래', '왼쪽'][i]} 여백`} type="number" min="0" max="75" step="0.1" value={Number(margins[side].toFixed(1))}
                    onChange={e => { const value = Number(e.target.value); if (Number.isFinite(value) && value >= 0 && value <= 75) setPageValue({ margins: { ...margins, [side]: value } }); }} />
                </label>)}
              </div></Group>
              <Group label="페이지 설정">
                <Dropdown label="여백" icon={<LayoutTemplate size={28} />}>
                  {[
                    { label: "기본 · 25.4 mm", value: 25.4 },
                    { label: "좁게 · 12.7 mm", value: 12.7 },
                    { label: "보통 · 20 mm", value: 20 },
                    { label: "넓게 · 38.1 mm", value: 38.1 },
                  ].map((m) => (
                    <MenuItem
                      key={m.value}
                      onClick={() => setPageValue({ margin: m.value })}
                    >
                      {page.margin === m.value ? "✓ " : ""}
                      {m.label}
                    </MenuItem>
                  ))}
                </Dropdown>
                <Dropdown
                  label="용지 방향"
                  icon={
                    page.orientation === "portrait" ? (
                      <RectangleVertical size={28} />
                    ) : (
                      <RectangleHorizontal size={28} />
                    )
                  }
                >
                  <MenuItem
                    onClick={() => setPageValue({ orientation: "portrait" })}
                  >
                    세로
                  </MenuItem>
                  <MenuItem
                    onClick={() => setPageValue({ orientation: "landscape" })}
                  >
                    가로
                  </MenuItem>
                </Dropdown>
                <Dropdown label="크기" icon={<FileText size={28} />}>
                  <MenuItem onClick={() => setPageValue({ size: "A4" })}>
                    A4 · 210 × 297 mm
                  </MenuItem>
                  <MenuItem onClick={() => setPageValue({ size: "Letter" })}>
                    Letter · 216 × 279 mm
                  </MenuItem>
                </Dropdown>
                <Tool
                  label="페이지 나누기"
                  icon={<BetweenHorizontalStart />}
                  onClick={() =>
                    editor
                      .chain()
                      .focus()
                      .insertContent({ type: "pageBreak" })
                      .run()
                  }
                />
              </Group>
              <Group label="단락">
                <Tool
                  label="내어쓰기"
                  icon={<IndentDecrease />}
                  onClick={() => setIndent(-24)}
                />
                <Tool
                  label="들여쓰기"
                  icon={<IndentIncrease />}
                  onClick={() => setIndent(24)}
                />
                <div className="layout-fields">
                  <label>
                    줄 간격{" "}
                    <select
                      aria-label="단락 줄 간격"
                      onChange={(e) =>
                        editor
                          .chain()
                          .focus()
                          .updateAttributes(
                            editor.isActive("heading")
                              ? "heading"
                              : "paragraph",
                            { lineSpacing: e.target.value, wordLineRule: editor.getAttributes(editor.isActive('heading') ? 'heading' : 'paragraph').wordLineRule ? 'auto' : null },
                          )
                          .run()
                      }
                      defaultValue="1.65"
                    >
                      {["1", "1.15", "1.5", "1.65", "2", "3"].map((x) => (
                        <option key={x}>{x}</option>
                      ))}
                    </select>
                  </label>
                  <span>
                    여백 {page.margin} mm · {page.size}
                  </span>
                </div>
              </Group>
            </>
          )}
          {tab === "검토" && (
            <>
              <Group label="문서 검사">
                <Tool
                  label="단어 개수"
                  icon={<TextCursorInput size={26} />}
                  onClick={() => setModal("stats")}
                />
              </Group>
              <Group label="메모">
                <Tool
                  label="새 메모"
                  icon={<MessageSquare size={26} />}
                  onClick={addComment}
                />
                <Tool
                  label="메모 표시"
                  icon={<PanelRight size={26} />}
                  active={commentsOpen}
                  onClick={() => setCommentsOpen(!commentsOpen)}
                />
              </Group>
              <Group label="읽기">
                <Tool
                  label={readOnly ? "편집 모드" : "읽기 모드"}
                  icon={<BookOpen size={26} />}
                  active={readOnly}
                  onClick={() => setReadOnly(!readOnly)}
                />
              </Group>
              <div className="ribbon-note">
                메모로 생각을 남기고, 문서를 함께 검토하세요.
                <br />
                <span>이 문서의 메모는 현재 PC에 저장됩니다.</span>
              </div>
            </>
          )}
          {tab === "보기" && (
            <>
              <Group label="출력"><Tool label="인쇄 미리보기" icon={<Printer />} onClick={() => { void preview().catch(e => notify(String(e))); }} /></Group>
              <Group label="표시">
                <Tool
                  label="탐색 창"
                  icon={<PanelLeft size={26} />}
                  active={navigation}
                  onClick={() => setNavigation(!navigation)}
                />
                <Tool
                  label="눈금자"
                  icon={<BetweenHorizontalStart size={26} />}
                  active={ruler}
                  onClick={() => setRuler(!ruler)}
                />
                <Tool
                  label="편집 기호"
                  icon={<span className="symbol">¶</span>}
                  active={marks}
                  onClick={() => setMarks(!marks)}
                />
              </Group>
              <Group label="확대 / 축소">
                <Tool
                  label="100%"
                  icon={<ZoomIn size={26} />}
                  onClick={() => setZoom(100)}
                />
                <Tool
                  label="페이지 너비"
                  icon={<Columns2 size={26} />}
                  onClick={fit}
                />
                <Tool
                  label="한 페이지"
                  icon={<FileText size={26} />}
                  onClick={() =>
                    setZoom(
                      Math.max(
                        40,
                        Math.floor(
                          (((workspaceRef.current?.clientHeight || 800) - 65) /
                            pageHeight) *
                            100,
                        ),
                      ),
                    )
                  }
                />
              </Group>
              <Group label="화면 모드">
                <Tool
                  label="읽기 모드"
                  icon={<BookOpen size={26} />}
                  active={readOnly}
                  onClick={() => setReadOnly(!readOnly)}
                />
                <Tool
                  label="페이지 색 전환"
                  icon={darkPaper ? <Sun size={26} /> : <Moon size={26} />}
                  active={darkPaper}
                  onClick={() => setDarkPaper(!darkPaper)}
                />
              </Group>
              <Group label="Studio">
                <Tool
                  label="앱 정보"
                  icon={<FileText size={26} />}
                  onClick={() => setModal("about")}
                />
              </Group>
            </>
          )}
          {tab === "표 레이아웃" && (
            <>
              <Group label="행과 열">
                <Tool
                  label="위에 삽입"
                  icon={<ArrowUp />}
                  onClick={() => editor.chain().focus().addRowBefore().run()}
                />
                <Tool
                  label="아래에 삽입"
                  icon={<ArrowDown />}
                  onClick={() => editor.chain().focus().addRowAfter().run()}
                />
                <Tool
                  label="왼쪽에 삽입"
                  icon={<IndentDecrease />}
                  onClick={() => editor.chain().focus().addColumnBefore().run()}
                />
                <Tool
                  label="오른쪽에 삽입"
                  icon={<IndentIncrease />}
                  onClick={() => editor.chain().focus().addColumnAfter().run()}
                />
              </Group>
              <Group label="병합">
                <Tool
                  label="셀 병합"
                  icon={<Table2 />}
                  onClick={() => editor.chain().focus().mergeCells().run()}
                />
                <Tool
                  label="셀 분할"
                  icon={<Columns2 />}
                  onClick={() => editor.chain().focus().splitCell().run()}
                />
                <Tool
                  label="머리글 행"
                  icon={<LayoutTemplate />}
                  onClick={() => editor.chain().focus().toggleHeaderRow().run()}
                />
              </Group>
              <Group label="삭제">
                <Tool
                  label="행 삭제"
                  icon={<Minus />}
                  onClick={() => editor.chain().focus().deleteRow().run()}
                />
                <Tool
                  label="열 삭제"
                  icon={<Minus />}
                  onClick={() => editor.chain().focus().deleteColumn().run()}
                />
                <Tool
                  label="표 삭제"
                  icon={<Trash2 />}
                  onClick={() => editor.chain().focus().deleteTable().run()}
                />
              </Group>
            </>
          )}
          <button
            className="collapse-ribbon"
            title="리본 메뉴 접기"
            onClick={() => setCollapsed(true)}
          >
            <ChevronUp size={13} />
          </button>
        </div>
      )}
      {collapsed && (
        <button className="expand-ribbon" onClick={() => setCollapsed(false)}>
          <ChevronDown size={13} />
          리본 메뉴 펼치기
        </button>
      )}
      <main className="workspace-shell">
        {navigation && (
          <aside className="navigation-pane">
            <div className="pane-heading">
              <h3>탐색</h3>
              <IconButton
                label="탐색 창 닫기"
                onClick={() => setNavigation(false)}
              >
                <X size={17} />
              </IconButton>
            </div>
            <div className="nav-tabs">
              {["제목", "검색", "바꾸기"].map((t) => (
                <button
                  className={navTab === t ? "selected" : ""}
                  key={t}
                  onClick={() => setNavTab(t)}
                >
                  {t}
                </button>
              ))}
            </div>
            {navTab === "제목" ? (
              <div className="outline">
                {headings.length ? (
                  headings.map((h) => (
                    <button
                      key={h.pos}
                      style={{ paddingLeft: 16 + (h.level - 1) * 14 }}
                      onClick={() =>
                        editor
                          .chain()
                          .focus()
                          .setTextSelection(h.pos + 1)
                          .scrollIntoView()
                          .run()
                      }
                    >
                      {h.title || "제목"}
                    </button>
                  ))
                ) : (
                  <div className="empty-pane">
                    <Heading1 size={28} />
                    <p>문서의 흐름을 한눈에</p>
                    <span>
                      홈에서 제목 스타일을 적용하면 여기에 표시됩니다.
                    </span>
                  </div>
                )}
              </div>
            ) : (
              <div className="search-pane">
                <div className="search-box">
                  <Search size={15} />
                  <input
                    aria-label="문서에서 찾기"
                    placeholder="문서에서 찾기"
                    value={query}
                    onChange={(e) => {
                      setQuery(e.target.value);
                      setFindIndex(0);
                    }}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") selectMatch(findIndex + 1);
                    }}
                  />
                </div>
                {navTab === "바꾸기" && (
                  <>
                    <input
                      className="replace-input"
                      aria-label="바꿀 내용"
                      placeholder="바꿀 내용"
                      value={replace}
                      onChange={(e) => setReplace(e.target.value)}
                    />
                    <div className="replace-actions" inert={readOnly}>
                      <button onClick={() => replaceMatch()}>바꾸기</button>
                      <button onClick={() => replaceMatch(true)}>
                        모두 바꾸기
                      </button>
                    </div>
                  </>
                )}
                <div className="search-summary">
                  <span>
                    {query ? `${matches.length}개 결과` : "검색어를 입력하세요"}
                  </span>
                  <IconButton
                    label="이전 검색 결과"
                    onClick={() => selectMatch(findIndex - 1)}
                  >
                    <ChevronUp size={15} />
                  </IconButton>
                  <IconButton
                    label="다음 검색 결과"
                    onClick={() => selectMatch(findIndex + 1)}
                  >
                    <ChevronDown size={15} />
                  </IconButton>
                </div>
                {matches.map((m, i) => (
                  <button
                    className={`search-result ${i === findIndex ? "selected" : ""}`}
                    key={m.from}
                    onClick={() => selectMatch(i)}
                  >
                    <small>{i + 1}</small>
                    {m.preview.slice(0, 130)}
                  </button>
                ))}
              </div>
            )}
          </aside>
        )}
        <div
          className={`document-workspace ${darkPaper ? "dark-paper" : ""}`}
          ref={workspaceRef}
          onScroll={() => {
            if (workspaceRef.current)
              setCurrentPage(
                Math.min(
                  pages,
                  Math.max(
                    1,
                    Math.floor(
                      workspaceRef.current.scrollTop /
                        (((pageHeight + 24) * zoom) / 100),
                    ) + 1,
                  ),
                ),
              );
          }}
        >
          {readOnly && (
            <div className="reading-banner">
              <BookOpen size={14} />
              읽기 모드{" "}
              <button onClick={() => setReadOnly(false)}>
                편집으로 돌아가기
              </button>
            </div>
          )}
          <div
            className="zoom-stage"
            style={{
              width: ((pageWidth + 52) * zoom) / 100,
              minHeight: ((pages * (pageHeight + 24) + 65) * zoom) / 100,
            }}
          >
            <div
              className="scaled-document"
              style={{ width: pageWidth, transform: `scale(${zoom / 100})` }}
            >
              {ruler && (
                <div
                  className="ruler"
                  style={{ "--margin": `${margin}px` } as CSSProperties}
                >
                  <span className="ruler-indent left" />
                  <span className="ruler-indent right" />
                  {Array.from(
                    { length: Math.floor(pageWidth / 37.8) },
                    (_, i) => (
                      <span key={i}>{i === 0 ? "" : i}</span>
                    ),
                  )}
                </div>
              )}
              <div
                className={`document-stack ${marks ? "show-marks" : ""}`}
                data-page-height={pageHeight}
                data-margin={margin}
                data-bottom-margin={margins.bottom * 96 / 25.4}
                ref={stackRef}
                style={
                  {
                    width: pageWidth,
                    minHeight: pages * (pageHeight + 24) - 24,
                    "--page-width": `${pageWidth}px`,
                    "--page-height": `${pageHeight}px`,
                    "--page-margin": `${margin}px`,
                    "--page-padding": `${margins.top * 96 / 25.4}px ${margins.right * 96 / 25.4}px ${margins.bottom * 96 / 25.4}px ${margins.left * 96 / 25.4}px`,
                  } as CSSProperties
                }
              >
                <div className="paper-backgrounds" aria-hidden="true">
                  {Array.from({ length: pages }, (_, i) => (
                    <div
                      key={i}
                      className="paper"
                      style={{ top: i * (pageHeight + 24), height: pageHeight }}
                    >
                      <div
                        className="paper-header document-furniture"
                        style={{
                          left: margins.left * 96 / 25.4,
                          right: margins.right * 96 / 25.4,
                          top: page.headerDistance !== undefined ? page.headerDistance * 96 / 25.4 : Math.max(12, margin / 2 - 8),
                        }}
                        dangerouslySetInnerHTML={{ __html: sanitizeImportedHtml(pageFurniture(page, i + 1).headerHtml || `<p>${escapeHtml(pageFurniture(page, i + 1).header)}</p>`).html }}
                      />
                      <div
                        className="paper-footer"
                        style={{
                          left: margins.left * 96 / 25.4,
                          right: margins.right * 96 / 25.4,
                          bottom: page.footerDistance !== undefined ? page.footerDistance * 96 / 25.4 : Math.max(12, margins.bottom * 96 / 25.4 / 2 - 8),
                        }}
                      >
                        <div className="document-furniture" dangerouslySetInnerHTML={{ __html: sanitizeImportedHtml(pageFurniture(page, i + 1).footerHtml || `<p>${escapeHtml(pageFurniture(page, i + 1).footer)}</p>`).html }} />
                        {pageFurniture(page, i + 1).pageNumbers && !/data-page-number/.test((pageFurniture(page, i + 1).headerHtml || '') + (pageFurniture(page, i + 1).footerHtml || '')) && <span>{i + 1}</span>}
                      </div>
                    </div>
                  ))}
                </div>
                <EditorContent editor={editor} />
                <TableContextMenu editor={editor} readOnly={readOnly} notify={notify} />
              </div>
            </div>
          </div>
        </div>
        {commentsOpen && (
          <aside className="comments-pane">
            <div className="pane-heading">
              <h3>메모</h3>
              <IconButton
                label="메모 창 닫기"
                onClick={() => setCommentsOpen(false)}
              >
                <X size={17} />
              </IconButton>
            </div>
            <button className="new-comment" onClick={addComment}>
              <Plus size={15} />새 메모
            </button>
            {comments.length ? (
              comments.map((c) => (
                <article
                  key={c.id}
                  className={`comment-card ${c.resolved ? "resolved" : ""}`}
                >
                  <div className="comment-meta">
                    <span className="avatar">나</span>
                    <strong>나</strong>
                    <time>
                      {new Date(c.createdAt).toLocaleDateString("ko-KR")}
                    </time>
                  </div>
                  {c.quote && <blockquote>{c.quote}</blockquote>}
                  <button onClick={() => {
                    let from = -1, to = -1;
                    editor.state.doc.descendants((node, pos) => {
                      if (node.marks.some(mark => mark.type.name === 'commentAnchor' && mark.attrs.ids.split(' ').includes(c.id))) {
                        if (from < 0) from = pos; to = pos + node.nodeSize;
                      }
                    });
                    if (from >= 0) editor.chain().focus().setTextSelection({ from, to }).scrollIntoView().run();
                    else notify('연결된 본문이 삭제됐거나 문서 전체에 남긴 메모입니다.');
                  }}>본문으로 이동</button>
                  {c.author && <small>{c.author}{c.initials ? ` (${c.initials})` : ''}</small>}
                  {c.html ? <div className="comment-story" dangerouslySetInnerHTML={{ __html: sanitizeImportedHtml(c.html).html }} /> : <p>{c.text}</p>}
                  <button
                    onClick={() =>
                      setComments(
                        comments.map((x) =>
                          x.id === c.id ? { ...x, resolved: !x.resolved } : x,
                        ),
                      )
                    }
                  >
                    {c.resolved ? (
                      <>
                        <Check size={13} />
                        해결됨 · 다시 열기
                      </>
                    ) : (
                      <>
                        <Check size={13} />
                        해결
                      </>
                    )}
                  </button>
                  <button
                    title="메모 삭제"
                    onClick={() => {
                      const tr = editor.state.tr;
                      editor.state.doc.descendants((node, pos) => {
                        const mark = node.marks.find(m => m.type.name === 'commentAnchor');
                        if (!mark) return;
                        const ids = mark.attrs.ids.split(' ').filter((id: string) => id !== c.id);
                        tr.removeMark(pos, pos + node.nodeSize, mark.type);
                        if (ids.length) tr.addMark(pos, pos + node.nodeSize, mark.type.create({ ids: ids.join(' ') }));
                      });
                      editor.view.dispatch(tr);
                      setComments(comments.filter((x) => x.id !== c.id));
                    }}
                  >
                    <Trash2 size={13} />
                  </button>
                </article>
              ))
            ) : (
              <div className="empty-pane">
                <MessageSquare size={30} />
                <p>생각을 남겨 보세요</p>
                <span>문장을 선택하고 새 메모를 눌러 의견을 기록하세요.</span>
              </div>
            )}
          </aside>
        )}
      </main>
      {printPreview && <div className="print-preview-controls" role="toolbar" aria-label="인쇄 미리보기">
        <strong>인쇄 미리보기 · {printPreview.pages}쪽</strong>
        <span className="preview-warning">{printPreview.warnings.join(' ') || '이 페이지 배치로 PDF와 인쇄를 출력합니다.'}</span>
        <button onClick={() => print(true)}>PDF 저장</button>
        <button onClick={() => print()}>인쇄</button>
        <button onClick={closePreview}>편집으로 돌아가기</button>
      </div>}
      <footer className="statusbar">
        <button
          onClick={() => {
            setNavigation(!navigation);
            setNavTab("제목");
          }}
        >
          {currentPage}/{pages} 페이지
        </button>
        <span className="status-divider" />
        <button onClick={() => setModal("stats")}>
          {words.toLocaleString()}개 단어
        </button>
        <span className="status-divider" />
        <span>한국어</span>
        <span className="local-status">
          <span />
          오프라인 · 로컬 문서
        </span>
        <div className="status-spacer" />
        <IconButton
          label="읽기 모드 전환"
          active={readOnly}
          onClick={() => setReadOnly(!readOnly)}
        >
          <BookOpen size={15} />
        </IconButton>
        <IconButton
          label="인쇄 모양"
          active={!readOnly}
          onClick={() => setReadOnly(false)}
        >
          <FileText size={15} />
        </IconButton>
        <span className="status-divider" />
        <IconButton
          label="축소"
          onClick={() => setZoom((z) => Math.max(40, z - 10))}
        >
          <Minus size={14} />
        </IconButton>
        <input
          aria-label="확대 축소"
          type="range"
          min="40"
          max="200"
          step="5"
          value={zoom}
          onChange={(e) => setZoom(Number(e.target.value))}
        />
        <IconButton
          label="확대"
          onClick={() => setZoom((z) => Math.min(200, z + 10))}
        >
          <Plus size={14} />
        </IconButton>
        <button className="zoom-value" onClick={() => setZoom(100)}>
          {zoom}%
        </button>
      </footer>
      <input
        ref={fileInput}
        type="file"
        hidden
        accept=".wordstudio,.docx,.html,.htm,.txt"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void readFile(f);
          e.target.value = "";
        }}
      />
      <input
        ref={imageInput}
        type="file"
        hidden
        accept="image/png,image/jpeg,image/gif,image/webp"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void insertImage(f);
          e.target.value = "";
        }}
      />
      {fileMenu && (
        <div className="file-backdrop" onClick={() => setFileMenu(false)}>
          <div className="backstage" onClick={(e) => e.stopPropagation()}>
            <div className="backstage-nav">
              <button
                className="back-button"
                onClick={() => setFileMenu(false)}
              >
                ← 문서로 돌아가기
              </button>
              <h2>Word Studio</h2>
              {actions.slice(0, 6).map((a) => (
                <button
                  key={a.name}
                  onClick={() => {
                    setFileMenu(false);
                    void a.action();
                  }}
                >
                  {a.name}
                  <small>{a.hint}</small>
                </button>
              ))}
              <button onClick={() => save("html")}>HTML로 내보내기</button>
              <button onClick={() => save("txt")}>텍스트로 내보내기</button>
              <button
                onClick={() => {
                  setFileMenu(false);
                  setModal("about");
                }}
              >
                앱 정보
              </button>
            </div>
            <div className="backstage-content">
              <span className="eyebrow">YOUR LOCAL WORKSPACE</span>
              <h1>문서에 담는 새로운 생각.</h1>
              <p>작성 중인 문서는 현재 PC에 자동으로 저장됩니다.</p>
              <div className="file-cards">
                <button
                  onClick={() => {
                    setFileMenu(false);
                    setModal("new");
                  }}
                >
                  <FilePlus2 size={34} />
                  <strong>새 문서</strong>
                  <span>빈 페이지에서 시작</span>
                </button>
                <button onClick={open}>
                  <FolderOpen size={34} />
                  <strong>파일 열기</strong>
                  <span>DOCX · Word Studio · HTML · TXT</span>
                </button>
              </div>
              <h3>최근 작업한 문서</h3>
              {readPreviousVersions().find(d => d.id === documentId) && <button className="recent-document" onClick={() => {
                const previous = readPreviousVersions().find(d => d.id === documentId);
                if (previous) void loadDoc({ ...previous, id: crypto.randomUUID(), title: `${previous.title} · 복원본` });
              }}>이 문서의 이전 정상 버전을 복사해 열기</button>}
              {history.length ? (
                history.map((d, i) => (
                  <button
                    className="recent-document"
                    key={i}
                    onClick={() => { void loadDoc(d); }}
                  >
                    <FileText size={22} />
                    <span>
                      {d.title}
                      <small>
                        {new Date(d.updatedAt).toLocaleString("ko-KR")}
                      </small>
                    </span>
                  </button>
                ))
              ) : (
                <p className="muted">
                  다른 문서를 열면 여기에 작업 기록이 표시됩니다.
                </p>
              )}
            </div>
          </div>
        </div>
      )}
      {command && (
        <div className="modal-overlay" onClick={() => setCommand(false)}>
          <div className="command-dialog" onClick={(e) => e.stopPropagation()}>
            <div className="command-input">
              <Search size={20} />
              <input
                autoFocus
                placeholder="명령 검색…"
                aria-label="명령 검색"
                value={commandQuery}
                onChange={(e) => setCommandQuery(e.target.value)}
              />
              <kbd>ESC</kbd>
            </div>
            {actions
              .filter((a) => a.name.includes(commandQuery))
              .map((a) => (
                <button
                  key={a.name}
                  onClick={() => {
                    setCommand(false);
                    setCommandQuery("");
                    void a.action();
                  }}
                >
                  {a.name}
                  <kbd>{a.hint}</kbd>
                </button>
              ))}
          </div>
        </div>
      )}
      {modal && (
        <div className="modal-overlay" onMouseDown={() => setModal(null)}>
          <section
            className={`dialog ${modal === "new" ? "template-dialog" : ""}`}
            role="dialog"
            aria-modal="true"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <div className="dialog-heading">
              <h2>
                {
                  {
                    new: "새 문서",
                    link: "링크 삽입",
                    header: "머리글 및 바닥글",
                    comment: "새 메모",
                    stats: "문서 통계",
                    about: "Word Studio",
                  }[modal]
                }
              </h2>
              <IconButton label="대화상자 닫기" onClick={() => setModal(null)}>
                <X size={19} />
              </IconButton>
            </div>
            {modal === "new" && (
              <>
                <p className="muted">
                  현재 문서는 최근 작업 목록에 보관됩니다.
                </p>
                <div className="template-grid">
                  {[
                    {
                      name: "빈 문서",
                      hint: "자유롭게 시작하세요",
                      html: "<p></p>",
                    },
                    {
                      name: "프로젝트 기획안",
                      hint: "목표부터 실행 계획까지",
                      html: "<h1>프로젝트 기획안</h1><p>프로젝트명 · 작성자 · 작성일</p><h2>1. 프로젝트 개요</h2><p>이 프로젝트의 목적과 배경을 작성하세요.</p><h2>2. 목표 및 기대 효과</h2><ul><li><p>달성하고자 하는 목표</p></li><li><p>기대하는 변화</p></li></ul><h2>3. 추진 일정</h2><table><tr><th><p>단계</p></th><th><p>일정</p></th><th><p>담당</p></th></tr><tr><td><p>기획</p></td><td><p></p></td><td><p></p></td></tr><tr><td><p>실행</p></td><td><p></p></td><td><p></p></td></tr></table><h2>4. 다음 단계</h2><p>첫 번째 실행 항목을 작성하세요.</p>",
                    },
                    {
                      name: "회의록",
                      hint: "결정 사항과 다음 할 일",
                      html:
                        "<h1>회의록</h1><p>날짜: " +
                        new Date().toLocaleDateString("ko-KR") +
                        "</p><p>참석자:</p><h2>안건</h2><ol><li><p>첫 번째 안건</p></li></ol><h2>논의 내용</h2><p></p><h2>결정 사항</h2><ul><li><p></p></li></ul><h2>다음 할 일</h2><table><tr><th><p>할 일</p></th><th><p>담당자</p></th><th><p>기한</p></th></tr><tr><td><p></p></td><td><p></p></td><td><p></p></td></tr></table>",
                    },
                  ].map((t, i) => (
                    <button
                      className="template-card"
                      key={t.name}
                      onClick={() =>
                        void loadDoc(blank(i === 0 ? "문서1" : t.name, t.html))
                      }
                    >
                      <div className="template-preview">
                        <div className="mini-page">
                          {i !== 0 && (
                            <>
                              <b />
                              <i />
                              <i />
                              <i />
                              <strong />
                              <i />
                              <i />
                              <strong />
                              <i />
                            </>
                          )}
                        </div>
                      </div>
                      <strong>{t.name}</strong>
                      <span>{t.hint}</span>
                    </button>
                  ))}
                </div>
              </>
            )}
            {modal === "link" && (
              <>
                <label className="dialog-field">
                  주소
                  <input
                    autoFocus
                    aria-label="링크 주소"
                    value={modalValue}
                    onChange={(e) => setModalValue(e.target.value)}
                    placeholder="https://example.com"
                  />
                </label>
                <div className="dialog-actions">
                  <button
                    onClick={() => {
                      editor.chain().focus().unsetLink().run();
                      setModal(null);
                    }}
                  >
                    링크 제거
                  </button>
                  <button
                    className="primary"
                    onClick={() => {
                      if (!/^(https?:\/\/|mailto:|tel:)/i.test(modalValue)) {
                        notify(
                          "http, https, mailto 또는 tel 주소를 입력하세요",
                        );
                        return;
                      }
                      if (editor.state.selection.empty)
                        editor
                          .chain()
                          .focus()
                          .insertContent({
                            type: "text",
                            text: modalValue,
                            marks: [
                              { type: "link", attrs: { href: modalValue } },
                            ],
                          })
                          .run();
                      else
                        editor
                          .chain()
                          .focus()
                          .setLink({ href: modalValue })
                          .run();
                      setModal(null);
                    }}
                  >
                    삽입
                  </button>
                </div>
              </>
            )}
            {modal === "header" && (
              <>
                <label className="dialog-field">
                  머리글
                  <input
                    aria-label="머리글"
                    value={page.header}
                    onChange={(e) => setPageValue({ header: e.target.value })}
                  />
                </label>
                <label className="dialog-field">
                  바닥글
                  <input
                    aria-label="바닥글"
                    value={page.footer}
                    onChange={(e) => setPageValue({ footer: e.target.value })}
                  />
                </label>
                <label className="checkbox-label">
                  <input
                    type="checkbox"
                    checked={page.pageNumbers}
                    onChange={(e) =>
                      setPageValue({ pageNumbers: e.target.checked })
                    }
                  />
                  페이지 번호 표시
                </label>
                {(['first', 'even'] as const).map(kind => {
                  const first = kind === 'first', enabled = first ? page.differentFirstPage : page.differentOddEven;
                  const label = first ? '첫 페이지' : '짝수 페이지';
                  return <div key={kind}>
                    <label className="checkbox-label"><input type="checkbox" checked={!!enabled}
                      onChange={e => setPageValue(first ? { differentFirstPage: e.target.checked } : { differentOddEven: e.target.checked })} />
                      {first ? '첫 페이지 다르게 지정' : '홀수·짝수 페이지 다르게 지정'}
                    </label>
                    {enabled && <>
                      <label className="dialog-field">{label} 머리글<input aria-label={`${label} 머리글`} value={page[`${kind}Header`] ?? ''}
                        onChange={e => setPageValue({ [`${kind}Header`]: e.target.value })} /></label>
                      <label className="dialog-field">{label} 바닥글<input aria-label={`${label} 바닥글`} value={page[`${kind}Footer`] ?? ''}
                        onChange={e => setPageValue({ [`${kind}Footer`]: e.target.value })} /></label>
                      <label className="checkbox-label"><input type="checkbox" checked={page[`${kind}PageNumbers`] ?? false}
                        onChange={e => setPageValue({ [`${kind}PageNumbers`]: e.target.checked })} />{label} 페이지 번호 표시</label>
                    </>}
                  </div>;
                })}
                <div className="dialog-actions">
                  <button className="primary" onClick={() => setModal(null)}>
                    완료
                  </button>
                </div>
              </>
            )}
            {modal === "comment" && (
              <>
                {commentQuote && (
                  <blockquote className="selected-quote">
                    {commentQuote}
                  </blockquote>
                )}
                <textarea
                  autoFocus
                  aria-label="메모 내용"
                  placeholder="의견이나 아이디어를 남겨 보세요…"
                  value={modalValue}
                  onChange={(e) => setModalValue(e.target.value)}
                />
                <div className="dialog-actions">
                  <button onClick={() => setModal(null)}>취소</button>
                  <button
                    className="primary"
                    disabled={!modalValue.trim()}
                    onClick={() => {
                      const id = crypto.randomUUID();
                      const range = commentRange.current;
                      if (range && range.from < range.to) {
                        const tr = editor.state.tr;
                        editor.state.doc.nodesBetween(range.from, range.to, (node, pos) => {
                          if (!node.isText) return;
                          const prior = node.marks.find(m => m.type.name === 'commentAnchor')?.attrs.ids || '';
                          tr.addMark(Math.max(pos, range.from), Math.min(pos + node.nodeSize, range.to), editor.schema.marks.commentAnchor.create({ ids: [...new Set([...prior.split(' ').filter(Boolean), id])].join(' ') }));
                        });
                        editor.view.dispatch(tr);
                      }
                      setComments([
                        ...comments,
                        {
                          id,
                          quote: commentQuote,
                          text: modalValue.trim(),
                          createdAt: new Date().toISOString(),
                          resolved: false,
                        },
                      ]);
                      setCommentsOpen(true);
                      setModal(null);
                    }}
                  >
                    메모 추가
                  </button>
                </div>
              </>
            )}
            {modal === "stats" && (
              <div className="statistics">
                {[
                  ["페이지 (편집 화면)", pages],
                  ["단어", words],
                  ["문자 (공백 포함)", text.length],
                  ["문자 (공백 제외)", text.replace(/\s/g, "").length],
                  ["단락", editor.state.doc.childCount],
                ].map(([label, n]) => (
                  <div key={label}>
                    <span>{label}</span>
                    <strong>{Number(n).toLocaleString()}</strong>
                  </div>
                ))}
              </div>
            )}
            {modal === "about" && (
              <div className="about-content">
                <span className="about-logo">
                  <FileText size={40} />
                </span>
                <h3>
                  Word Studio <small>1.0</small>
                </h3>
                <p>생각을 담는 로컬 문서 편집기.</p>
                <p className="muted">
                  Slide Studio와 같은 Studio Office 디자인을 사용하는 독립
                  앱입니다. 문서와 메모는 현재 PC에 저장됩니다.
                </p>
                <p className="muted">
                  DOCX의 본문, 기본 서식, 표와 이미지를 지원합니다. 매크로, 변경
                  내용 추적, 복잡한 구역·도형 및 원본 페이지 배치는 완전히
                  재현되지 않을 수 있습니다.
                </p>
                <p className="muted">
                  Microsoft Word와 제휴하지 않은 독립 구현입니다.
                </p>
              </div>
            )}
          </section>
        </div>
      )}
      {failedTransition && <SaveFailureDialog title={title}
        onCancel={() => { setFailedTransition(null); transitioning.current = false; }}
        onDiscard={() => activateDocument(failedTransition)}
        onSave={async () => {
          const contents = documentContents(currentRef.current);
          if (!await save() || contents !== documentContents(currentRef.current)) return false;
          activateDocument(failedTransition);
          return true;
        }} />}
      {toast && (
        <div className="toast" role="status">
          <CheckCircle2 size={17} />
          <span>{toast}</span>
          <button aria-label="알림 닫기" onClick={() => setToast("")}>
            <X size={14} />
          </button>
        </div>
      )}
      {busy && <div className="busy-indicator">문서를 처리하고 있습니다…</div>}
    </div>
  );
}
