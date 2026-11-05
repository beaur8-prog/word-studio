import {Node} from '@tiptap/core';
import {parseDrawingGroup} from './drawing-group';

/** An inline Word group owns one inline extent. Its editable image and story
 * nodes remain separate; this empty atom reserves only the original flow box. */
export const DocumentDrawingAnchor=Node.create({
 name:'documentDrawingAnchor',group:'inline',inline:true,atom:true,selectable:false,
 addAttributes:()=>({frame:{default:'',parseHTML:el=>{const f=parseDrawingGroup(el.getAttribute('data-drawing-group-anchor'));return f?.position==='inline'?JSON.stringify(f):'';}}}),
 parseHTML:()=>[{tag:'span[data-drawing-group-anchor]',getAttrs:el=>parseDrawingGroup((el as HTMLElement).getAttribute('data-drawing-group-anchor'))?.position==='inline'?null:false}],
 renderHTML:({node})=>{const g=parseDrawingGroup(node.attrs.frame);return ['span',{'data-drawing-group-anchor':g?JSON.stringify(g):'',contenteditable:'false','aria-hidden':'true',style:g?`display:inline-block;width:${g.width}px;height:${g.height}px;vertical-align:baseline;line-height:0`:''}];},
});
