import {imageFloating} from './image-floating';
import {drawingGroup,type DrawingGroup} from './drawing-group';
import { Node, Extension } from '@tiptap/core';
import { Plugin, TextSelection } from '@tiptap/pm/state';
import { documentScale } from './document-scale';

export type TextBoxFrame = {
  group?: string; drawingGroup?: DrawingGroup;
  width: number; height: number; x: number; y: number;
  position: 'floating' | 'inline'; wrap?: 'none' | 'square'; noWrap?: boolean; autoFit?: boolean; horizontalAlign?: 'left'|'center'|'right'; horizontalReference?: 'page'|'margin'|'column'; horizontalOffset?: number; vertical: 'page' | 'paragraph';
  padding: [number, number, number, number];
  wrapMargins?: [number, number, number, number];
  align: 'top' | 'center' | 'bottom'; shape: 'rect' | 'homePlate' | 'roundRect';
  fill?: string; gradient?: { position: number; color: string }[];
  border?: string; borderWidth?: number; behind?: boolean;
};
const color = (value: unknown): value is string => typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value);
const number = (value: unknown, low: number, high: number): value is number => typeof value === 'number' && Number.isFinite(value) && value >= low && value <= high;
/** Only bounded geometry and resolved colors cross the document/HTML boundary. */
export function textBoxFrame(value: unknown): TextBoxFrame | undefined {
  if (typeof value !== 'string' || value.length > 10000) return;
  try {
    const v = JSON.parse(value);
    if(v?.drawingGroup!==undefined && !drawingGroup(v.drawingGroup))return;
    if(v?.group!==undefined && (typeof v.group!=='string' || !/^[a-z0-9_]{1,64}$/i.test(v.group)))return;
    if (!v || typeof v !== 'object' || !number(v.width, .1, 100000) || !number(v.height, .1, 100000) ||
      !number(v.x, -100000, 100000) || !number(v.y, -100000, 100000) || !['floating','inline'].includes(v.position) ||
      !['page','paragraph'].includes(v.vertical) || !Array.isArray(v.padding) || v.padding.length !== 4 ||
      !v.padding.every((n: unknown) => number(n, 0, 10000)) || !['top','center','bottom'].includes(v.align) ||
      !['rect','homePlate','roundRect'].includes(v.shape)) return;
    if(v.wrapMargins!==undefined && (!Array.isArray(v.wrapMargins) || v.wrapMargins.length!==4 || !v.wrapMargins.every((n:unknown)=>number(n,0,10000))))return;
    if(v.horizontalReference!==undefined && !['page','margin','column'].includes(v.horizontalReference) || v.horizontalOffset!==undefined && !number(v.horizontalOffset,-100000,100000))return;
    if(v.autoFit!==undefined && typeof v.autoFit!=='boolean' || v.horizontalAlign!==undefined && !['left','center','right'].includes(v.horizontalAlign))return;
    if (v.wrap !== undefined && !['none','square'].includes(v.wrap) || v.noWrap !== undefined && typeof v.noWrap !== 'boolean') return;
    if (v.fill !== undefined && !color(v.fill) || v.border !== undefined && !color(v.border) ||
      v.borderWidth !== undefined && !number(v.borderWidth, 0, 1000) || v.behind !== undefined && typeof v.behind !== 'boolean') return;
    if (v.gradient !== undefined && (!Array.isArray(v.gradient) || !v.gradient.length || v.gradient.length > 16 ||
      !v.gradient.every((s: any) => s && number(s.position, 0, 1) && color(s.color)))) return;
    return { ...(v.group?{group:v.group}:{}), ...(v.drawingGroup?{drawingGroup:drawingGroup(v.drawingGroup)}:{}), width:v.width, height:v.height, x:v.x, y:v.y, position:v.position, vertical:v.vertical,
      padding:v.padding, ...(v.wrapMargins?{wrapMargins:v.wrapMargins}:{}), align:v.align, shape:v.shape, ...(v.horizontalReference ? {horizontalReference:v.horizontalReference} : {}), ...(v.horizontalOffset!==undefined ? {horizontalOffset:v.horizontalOffset} : {}), ...(v.autoFit!==undefined ? {autoFit:v.autoFit} : {}), ...(v.horizontalAlign ? {horizontalAlign:v.horizontalAlign} : {}), ...(v.wrap ? {wrap:v.wrap} : {}), ...(v.noWrap!==undefined ? {noWrap:v.noWrap} : {}), ...(v.fill ? { fill:v.fill } : {}),
      ...(v.gradient ? { gradient:v.gradient.map((s: any) => ({ position:s.position, color:s.color })) } : {}),
      ...(v.border ? { border:v.border, borderWidth:v.borderWidth ?? 0 } : {}), ...(v.behind !== undefined ? { behind:v.behind } : {}) };
  } catch { return; }
}
function contentStyle(f: TextBoxFrame) {
  const background = f.gradient ? `linear-gradient(180deg,${f.gradient.map(s => `${s.color} ${s.position*100}%`).join(',')})` : f.fill || 'transparent';
  return `box-sizing:border-box;width:${f.autoFit && f.noWrap ? 'max-content' : f.width+'px'};min-height:${f.autoFit?0:f.height}px;padding:${f.padding.map(n => `${n}px`).join(' ')};background:${f.shape==='homePlate'?'transparent':background};${f.shape==='homePlate'?`--textbox-shape-fill:${background};--textbox-arrow-inset:${Math.min(f.height/2,f.width/2)}px;isolation:isolate;`:''}color:#000000;${f.noWrap?'white-space:nowrap;':''}${f.align!=='top'?`display:flex;flex-direction:column;justify-content:${f.align==='center'?'center':'flex-end'};`:''}${f.border ? `border:${f.borderWidth || 0}px solid ${f.border};` : ''}${f.shape==='roundRect' ? 'border-radius:12px;' : ''}`;
}
function applyFrame(dom: HTMLElement, content: HTMLElement, frame: TextBoxFrame) {
  dom.className = 'document-textbox'; dom.dataset.documentTextbox = JSON.stringify(frame);
  dom.dataset.textboxPosition = frame.position;
  dom.title=`텍스트 상자 · ${/Mac|iPhone|iPad/.test(navigator.platform)?'⌘':'Ctrl+'}Enter로 상자 밖에 입력`;
  content.dataset.textboxContent = 'true'; content.className = 'document-textbox-content';
  content.style.cssText = contentStyle(frame);
}
export const DocumentTextBox = Node.create({
  name:'documentTextBox', priority:110, group:'block', content:'block+', defining:true, isolating:true,
  addAttributes: () => ({ frame:{ default:'', parseHTML: el => { const f=textBoxFrame(el.getAttribute('data-document-textbox')); return f ? JSON.stringify(f) : ''; }, renderHTML: a => ({ 'data-document-textbox':a.frame }) } }),
  parseHTML: () => [{ tag:'div[data-document-textbox]', getAttrs: el => !!textBoxFrame((el as HTMLElement).getAttribute('data-document-textbox')) && null, contentElement:'[data-textbox-content]' }],
  renderHTML: ({ node }) => {
    const f=textBoxFrame(node.attrs.frame);
    return ['div', { class:'document-textbox','data-document-textbox':f ? JSON.stringify(f) : '', 'data-textbox-position':f?.position || 'inline' },
      ['div', { class:'document-textbox-content','data-textbox-content':'true',style:f ? contentStyle(f) : '' },0]];
  },
  addNodeView: () => ({ node }) => {
    const dom=document.createElement('div'),contentDOM=document.createElement('div');dom.append(contentDOM);
    const update=(next: typeof node) => { if(next.type.name!=='documentTextBox')return false;const f=textBoxFrame(next.attrs.frame);if(!f)return false;applyFrame(dom,contentDOM,f);return true; };
    update(node);return { dom,contentDOM,update,ignoreMutation:m => m.type==='attributes' && (m.target===dom || m.target===contentDOM) };
  },
  addKeyboardShortcuts() {
    return { 'Mod-End': () => {
      const {state,view}=this.editor; const {$from}=state.selection;
      for(let depth=$from.depth;depth>0;depth--) if($from.node(depth).type.name==='documentTextBox') {
        view.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve($from.end(depth)),-1)).scrollIntoView());return true;
      }
      let position=-1,hasBoxes=false;
      state.doc.descendants((node,pos)=>{if(node.type.name==='documentTextBox'){hasBoxes=true;return false;}if(node.isTextblock)position=pos+node.nodeSize-1;});
      if(!hasBoxes || position<0)return false;
      view.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(position),-1)).scrollIntoView());return true;
    }, 'Mod-Enter': () => {
      const { state,view }=this.editor; const { $from }=state.selection;
      for(let depth=$from.depth;depth>0;depth--) if($from.node(depth).type.name==='documentTextBox') {
        let after=$from.after(depth),next=state.doc.nodeAt(after);
        const group=textBoxFrame($from.node(depth).attrs.frame)?.group,tr=state.tr;
        // All adjacent stories in a Word group belong to one drawing anchor.
        // Exiting its first story must not insert body text between members.
        while(group && next?.type.name==='documentTextBox' && textBoxFrame(next.attrs.frame)?.group===group) {
          after+=next.nodeSize;next=state.doc.nodeAt(after);
        }
        if(!next?.isTextblock)tr.insert(after,state.schema.nodes.paragraph.create());
        tr.setSelection(TextSelection.near(tr.doc.resolve(after+1)));view.dispatch(tr.scrollIntoView());return true;
      }
      return false;
    } };
  },
});

/** Resolve the anchor against the same paper in editor and print. The outer
 * zero-height frame owns the story; typing never merges its runs into body. */
export function layoutTextBoxes(root: HTMLElement) {
  for (const frame of root.querySelectorAll<HTMLElement>('.document-textbox[data-document-textbox]')) {
    const f=textBoxFrame(frame.dataset.documentTextbox); if(!f || f.position!=='floating')continue;
    const sheet=frame.closest<HTMLElement>('.print-page,.document-section-content,.document-stack');if(!sheet)continue;
    const scale=documentScale(frame),paper=sheet.getBoundingClientRect(),rect=frame.getBoundingClientRect();if(!(scale>0))continue;
    const height=Number(sheet.dataset.pageHeight) || paper.height/scale;
    const carrier=f.drawingGroup && [...root.querySelectorAll<HTMLImageElement>('img[data-image-floating]')].find(img=>imageFloating(img.dataset.imageFloating)?.drawingGroup?.id===f.drawingGroup!.id)?.closest('p,h1,h2,h3,h4,h5,h6');
    const natural=((carrier?.getBoundingClientRect().top??rect.top)-paper.top)/scale, page=sheet.classList.contains('print-page') ? 0 : Math.max(0,Math.floor((natural+.5)/(height+24)));
    const actualWidth=frame.firstElementChild!.getBoundingClientRect().width/scale;
    const adjustment=f.autoFit && f.noWrap ? (f.width-actualWidth)*(f.horizontalAlign==='center'?.5:f.horizontalAlign==='right'?1:0):0;
    frame.style.setProperty('--textbox-left',`${f.x+adjustment-(rect.left-paper.left)/scale}px`);
    frame.style.setProperty('--textbox-top',`${f.vertical==='paragraph' ? f.y+(carrier?(carrier.getBoundingClientRect().top-rect.top)/scale:0) : page*(height+24)+f.y-(rect.top-paper.top)/scale}px`);
  }
}
function syncPaintedSelection(view: import('@tiptap/pm/view').EditorView) {
  if(view.composing)return;
  const sel=view.dom.ownerDocument.getSelection();if(!sel?.anchorNode || !sel.focusNode || !view.dom.contains(sel.anchorNode) || !view.dom.contains(sel.focusNode))return;
  const parent=(n:globalThis.Node)=>n instanceof HTMLElement?n:n.parentElement;
  const affected=(n:globalThis.Node)=>{const p=parent(n);return p?.closest('.document-textbox-content') || p?.closest('p,h1,h2,h3,h4,h5,h6')?.querySelector('[data-flow-widget]');};
  if(!affected(sel.anchorNode) && !affected(sel.focusNode))return;
  try{
    const anchor=view.posAtDOM(sel.anchorNode,sel.anchorOffset),head=view.posAtDOM(sel.focusNode,sel.focusOffset);
    if(anchor!==view.state.selection.anchor || head!==view.state.selection.head)view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc,anchor,head)));
  }catch{/* During a composition/mutation the native selection may be transient. */}
}
export const TextBoxLayout = Extension.create({ name:'textBoxLayout', addProseMirrorPlugins: () => [new Plugin({
  props:{handleDOMEvents:{
    // Native hit testing sees the painted textbox, while a zero-height anchor
    // can make ProseMirror's block hit test pick the underlying body paragraph.
    // Flow widgets can similarly shift native hit tests in wrapped body text.
    // Adopt the actual DOM selection before commands; leave dragging/IME native.
    mouseup(view){syncPaintedSelection(view);return false;},
    keydown(view){syncPaintedSelection(view);return false;},
  }},
  view(view) {
  let frame=0;const update=()=>{cancelAnimationFrame(frame);frame=requestAnimationFrame(()=>{if(!view.composing)layoutTextBoxes(view.dom);});};
  const observer=new ResizeObserver(update);observer.observe(view.dom);document.fonts.addEventListener('loadingdone',update);update();
  return { update,destroy(){cancelAnimationFrame(frame);observer.disconnect();document.fonts.removeEventListener('loadingdone',update);} };
} })] });
