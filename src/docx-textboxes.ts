import { normalizeDrawingGroups } from './docx-drawing-groups';
import type JSZip from 'jszip';
import { W, all, child, ancestor } from './docx-format';
import { storyHtml } from './docx-stories';
import { pageMargins, paperDimensions, type PageSettings } from './types';
import { textBoxFrame, type TextBoxFrame } from './textbox-frame';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WPS='http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const EMU=9525;
const descendants=(root:Element|Document,name:string)=>Array.from(root.getElementsByTagNameNS('*',name));
const number=(el:Element|null|undefined,key:string,fallback=0)=>{const n=Number(el?.getAttribute(key)??fallback);return Number.isFinite(n)?n:fallback;};
const first=(el:Element|null|undefined,name:string)=>el && Array.from(el.children).find(c=>c.localName===name);
const hex=(v:string|undefined)=>v && /^#?[\da-f]{6}$/i.test(v) ? '#'+v.replace('#','') : undefined;
type PreparedBox={ marker:string; frames:{ frame:TextBoxFrame; html:string }[] };

async function colors(zip:JSZip) {
 const part=zip.file('word/theme/theme1.xml'),theme=part ? new DOMParser().parseFromString(await part.async('string'),'application/xml') : undefined;
 const scheme=theme && descendants(theme,'clrScheme')[0];const aliases:Record<string,string>={bg1:'lt1',bg2:'lt2',tx1:'dk1',tx2:'dk2'};
 function color(el:Element|undefined,placeholder?:string):string|undefined {
  if(!el)return;let value:string|undefined;
  if(el.localName==='srgbClr')value=hex(el.getAttribute('val')||'');
  else if(el.localName==='sysClr')value=hex(el.getAttribute('lastClr')||'');
  else if(el.localName==='schemeClr') {
   const key=el.getAttribute('val')||'';
   if(key==='phClr')value=placeholder;
   else {const item=scheme && first(scheme,aliases[key]||key);value=item ? color(item.firstElementChild||undefined) : key==='lt1' ? '#ffffff' : key==='dk1' ? '#000000' : undefined;}
  }
  if(!value)return;
  let rgb=value.slice(1).match(/../g)!.map(n=>parseInt(n,16)/255);
  for(const transform of el.children) {
   const n=number(transform,'val')/100000;
   if(transform.localName==='tint')rgb=rgb.map(c=>c+(1-c)*n);
   if(transform.localName==='shade')rgb=rgb.map(c=>c*n);
   if(transform.localName==='lumMod')rgb=rgb.map(c=>c*n);
   if(transform.localName==='lumOff')rgb=rgb.map(c=>c+n);
  }
  return '#'+rgb.map(c=>Math.round(Math.max(0,Math.min(1,c))*255).toString(16).padStart(2,'0')).join('');
 }
 function fill(shape:Element) {
  const props=first(shape,'spPr');let fill=props && Array.from(props.children).find(x=>['noFill','solidFill','gradFill'].includes(x.localName));let placeholder:string|undefined;
  if(!fill) {
   const ref=first(first(shape,'style'),'fillRef');placeholder=ref ? color(ref.firstElementChild||undefined) : undefined;
   const index=number(ref,'idx');const list=theme && descendants(theme,index>=1001 ? 'bgFillStyleLst' : 'fillStyleLst')[0];
   fill=list?.children[index>=1001 ? index-1001 : index-1];
  }
  if(fill?.localName==='noFill')return {};
  if(fill?.localName==='gradFill') {
   const gradient=descendants(fill,'gs').slice(0,16).map(s=>({position:Math.max(0,Math.min(1,number(s,'pos')/100000)),color:color(s.firstElementChild||undefined,placeholder)})).filter((s):s is {position:number;color:string}=>!!s.color);
   if(gradient.length)return {gradient};
  }
  const value=fill?.localName==='solidFill' ? color(fill.firstElementChild||undefined,placeholder) : placeholder;
  return value ? {fill:value} : {};
 }
 return {fill,color};
}
// Word can put a percentage anchor in Choice and a pixel fallback beside it.
// Resolve only direct compatibility wrappers; do not pick a child's anchor.
function anchorPosition(anchor:Element|null,name:string):Element|undefined {
 const direct=first(anchor,name);if(direct)return direct;
 for(const wrapper of anchor?.children||[])if(wrapper.localName==='AlternateContent') {
  for(const branch of wrapper.children) {
   const candidate=first(branch,name);if(!candidate)continue;
   if(first(candidate,'posOffset') || first(candidate,'align') || first(candidate,'pctPosHOffset') || first(candidate,'pctPosVOffset'))return candidate;
  }
 }
}
function modernFrame(box:Element,page:PageSettings,palette:Awaited<ReturnType<typeof colors>>):TextBoxFrame {
 const shape=ancestor(box,'wsp')!,anchor=ancestor(shape,'anchor')||ancestor(shape,'inline');
 const transform=first(first(shape,'spPr'),'xfrm'),offset=first(transform,'off'),size=first(transform,'ext');
 let x=number(offset,'x'),y=number(offset,'y'),width=number(size,'cx',number(anchor ? first(anchor,'extent') : undefined,'cx',3657600)),height=number(size,'cy',number(anchor ? first(anchor,'extent') : undefined,'cy',914400));
 let parent=shape.parentElement,grouped=false;
 while(parent && parent!==anchor) {
  if(parent.localName==='wgp') {grouped=true;
   const t=first(first(parent,'grpSpPr'),'xfrm'),off=first(t,'off'),ext=first(t,'ext'),childOff=first(t,'chOff'),childExt=first(t,'chExt');
   const sx=number(ext,'cx',1)/number(childExt,'cx',number(ext,'cx',1)),sy=number(ext,'cy',1)/number(childExt,'cy',number(ext,'cy',1));
   x=number(off,'x')+(x-number(childOff,'x'))*sx;y=number(off,'y')+(y-number(childOff,'y'))*sy;width*=sx;height*=sy;
  }
  parent=parent.parentElement;
 }
 const margins=pageMargins(page),dimensions=paperDimensions(page).map(n=>n*96/25.4);
 const position=(name:string,vertical:boolean)=>{
  const p=anchorPosition(anchor,name),relative=p?.getAttribute('relativeFrom')||'page',axis=vertical?1:0;
  const start=(vertical?margins.top:margins.left)*96/25.4,end=(vertical?margins.bottom:margins.right)*96/25.4;
  const isMargin=['margin','column'].includes(relative),base=isMargin?start:0,available=dimensions[axis]-(isMargin?start+end:0);
  const own=anchor ? number(first(anchor,'extent'),vertical?'cy':'cx')/EMU : (vertical?height:width)/EMU;
  const align=first(p,'align')?.textContent,pct=first(p,vertical?'pctPosVOffset':'pctPosHOffset');
  return {offset:pct ? base+available*Number(pct.textContent)/100000 : first(p,'posOffset') ? base+Number(first(p,'posOffset')!.textContent)/EMU : align==='center'?base+(available-own)/2 : ['right','bottom','outside'].includes(align||'')?base+available-own:base,relative};
 };
 const hp=position('positionH',false),vp=position('positionV',true),properties=first(shape,'bodyPr');
 const line=first(first(shape,'spPr'),'ln'),lineColor=first(line,'solidFill'),border=palette.color(lineColor?.firstElementChild||undefined);
 const hAlign=first(anchorPosition(anchor,'positionH'),'align')?.textContent||'';
 const hReference=grouped?'page':['margin','column'].includes(hp.relative)?hp.relative as 'margin'|'column':'page';
 const frame:TextBoxFrame={width:width/EMU,height:height/EMU,x:hp.offset+x/EMU,y:vp.offset+y/EMU,position:anchor?.localName==='inline'?'inline':'floating',vertical:['paragraph','line'].includes(vp.relative)?'paragraph':'page',
  wrapMargins:[number(anchor,'distT')/EMU,number(anchor,'distR')/EMU,number(anchor,'distB')/EMU,number(anchor,'distL')/EMU],
  padding:[number(properties,'tIns',45720)/EMU,number(properties,'rIns',91440)/EMU,number(properties,'bIns',45720)/EMU,number(properties,'lIns',91440)/EMU],
  align:properties?.getAttribute('anchor')==='ctr'?'center':properties?.getAttribute('anchor')==='b'?'bottom':'top',
  shape:['homePlate','roundRect'].includes(first(first(shape,'spPr'),'prstGeom')?.getAttribute('prst')||'') ? first(first(shape,'spPr'),'prstGeom')!.getAttribute('prst') as 'homePlate'|'roundRect':'rect',
  autoFit:!!first(properties,'spAutoFit'),horizontalReference:hReference,horizontalOffset:hp.offset+x/EMU-(hReference==='page'?0:margins.left*96/25.4),horizontalAlign:!grouped && ['center','right'].includes(hAlign)?hAlign as 'center'|'right':'left',
  wrap:anchor && Array.from(anchor.children).some(c=>['wrapSquare','wrapTight','wrapThrough','wrapTopAndBottom'].includes(c.localName))?'square':'none',noWrap:properties?.getAttribute('wrap')==='none',
  ...palette.fill(shape),...(border ? {border,borderWidth:number(line,'w')/EMU}:{}),behind:anchor?.getAttribute('behindDoc')==='1'};
 return textBoxFrame(JSON.stringify(frame)) || {...frame,x:0,y:0,width:400,height:100,position:'inline',padding:[5,10,5,10]};
}
function legacyFrame(box:Element,page:PageSettings):TextBoxFrame {
 const shape=ancestor(box,'shape')||ancestor(box,'rect')||ancestor(box,'roundrect')||ancestor(box,'textbox')?.parentElement;
 if(!shape)throw Error('텍스트상자 도형을 찾을 수 없습니다.');
 const styles=new Map((shape.getAttribute('style')||'').split(';').map(p=>p.split(':',2).map(s=>s.trim()) as [string,string]));
 const pixels=(v:string|undefined,fallback=0)=>{const m=v?.match(/^(-?[\d.]+)(pt|px|in|cm|mm)?$/);return m ? Number(m[1])*({pt:96/72,in:96,cm:96/2.54,mm:96/25.4,px:1}[m[2]||'px']||1):fallback;};
 const insets=(ancestor(box,'textbox')?.getAttribute('inset')||'7.2pt,3.6pt,7.2pt,3.6pt').split(',').map(s=>pixels(s.trim()));
 const margin=pageMargins(page),fill=hex(shape.getAttribute('fillcolor')?.split(' ')[0]);
 const result:TextBoxFrame={width:pixels(styles.get('width'),400),height:pixels(styles.get('height'),100),x:pixels(styles.get('margin-left'))+(styles.get('mso-position-horizontal-relative')==='page'?0:margin.left*96/25.4),y:pixels(styles.get('margin-top'))+(styles.get('mso-position-vertical-relative')==='page'?0:margin.top*96/25.4),position:styles.get('position')==='absolute'?'floating':'inline',vertical:'page',padding:[insets[1]||0,insets[2]||0,insets[3]||0,insets[0]||0],align:'top',shape:'rect',...(fill && shape.getAttribute('filled')!=='f'?{fill}:{})};
 return textBoxFrame(JSON.stringify(result)) || {...result,x:0,y:0,width:400,height:100,padding:[5,10,5,10]};
}

/** Isolate each selected textbox story before Mammoth can flatten/drop it.
 * Modern Choice wins once; the equivalent legacy fallback is never imported twice. */
export async function prepareTextBoxes(zip:JSZip,body:Document,pages:PageSettings[],depth=0) {
 if(depth>12)throw Error('중첩 텍스트상자 구조가 너무 깊습니다.');
 normalizeDrawingGroups(body,pages);
 const palette=await colors(zip),records:PreparedBox[]=[];
 const candidates=all(body,'txbxContent').filter(box=>{
  const fallback=ancestor(box,'Fallback');if(!fallback)return true;
  const alternate=fallback.parentElement;
  return !alternate || !Array.from(alternate.children).some(c=>c.localName==='Choice' && all(c,'txbxContent').length);
 });
 const grouped=new Map<Element,Element[]>();
 for(const box of candidates) {
  let container=ancestor(box,'AlternateContent')||ancestor(box,'drawing')||ancestor(box,'pict');
  if(!container || ancestor(box,'txbxContent'))continue;
  const boxes=grouped.get(container)||[];boxes.push(box);grouped.set(container,boxes);
 }
 let section=0;const pageFor=new Map<Element,PageSettings>();
 for(const paragraph of all(body,'p')) {pageFor.set(paragraph,pages[Math.min(section,pages.length-1)]||pages[0]);if(first(first(paragraph,'pPr'),'sectPr') && !ancestor(paragraph,'txbxContent'))section++;}
 for(const [container,boxes] of grouped) {
  const page=pageFor.get(ancestor(container,'p')!)||pages[0];const frames=[];
  for(const box of boxes) {
   const frame=ancestor(box,'wsp') ? modernFrame(box,page,palette):legacyFrame(box,page);
   const raw=ancestor(box,'drawing')?.getAttribute('data-ws-drawing-group');
   if(raw){frame.drawingGroup=JSON.parse(raw);frame.group=frame.drawingGroup!.id;}
   else if(boxes.length>1 && ancestor(box,'wgp'))frame.group='textbox_group_'+records.length;
   frames.push({frame,html:await storyHtml(zip,'word/document.xml',box,page,depth+1)});
  }
  const marker='ws_textbox_'+crypto.randomUUID().replaceAll('-','');
  const bookmark=body.createElementNS(W,'w:bookmarkStart');bookmark.setAttributeNS(W,'w:id',String(800000+records.length));bookmark.setAttributeNS(W,'w:name',marker);
  const end=body.createElementNS(W,'w:bookmarkEnd');end.setAttributeNS(W,'w:id',String(800000+records.length));
  container.before(bookmark,end);container.remove();records.push({marker,frames});
 }
 return {count:records.reduce((n,r)=>n+r.frames.length,0),decorate(html:string) {
  const template=document.createElement('template');template.innerHTML=html;const last=new Map<Element,Element>();
  for(const record of records) {
   const marker=Array.from(template.content.querySelectorAll('[data-docx-bookmark]')).find(e=>e.getAttribute('data-docx-bookmark')===record.marker);
   if(!marker)continue;
   const paragraph=marker.closest('p,h1,h2,h3,h4,h5,h6,li');let after=paragraph ? last.get(paragraph)||paragraph:marker;
   for(const item of record.frames) {
    const box=document.createElement('div'),content=document.createElement('div');box.dataset.documentTextbox=JSON.stringify(item.frame);content.dataset.textboxContent='true';content.innerHTML=item.html;box.append(content);after.after(box);after=box;
   }
   if(paragraph)last.set(paragraph,after);marker.remove();
  }
  return template.innerHTML;
 }};
}

export type TextBoxExport={ id:string; frame:TextBoxFrame };
/** Replace generated marker paragraphs with a real DrawingML text story. All
 * child paragraphs/tables have already gone through the normal DOCX exporter. */
export async function writeTextBoxes(zip:JSZip,records:TextBoxExport[]) {
 if(!records.length)return;
 for(const name of Object.keys(zip.files).filter(n=>/^word\/.*\.xml$/.test(n))) {
  const source=await zip.file(name)!.async('string');if(!records.some(r=>source.includes(r.id)))continue;
  const doc=new DOMParser().parseFromString(source,'application/xml');
  const written:{record:TextBoxExport;paragraph:Element}[]=[];
  for(const [index,record] of records.entries()) {
   const mark=all(doc,'bookmarkStart').find(e=>e.getAttributeNS(W,'name')===record.id+'_start');if(!mark)continue;
   const start=ancestor(mark,'p')!,end=all(doc,'bookmarkStart').find(e=>e.getAttributeNS(W,'name')===record.id+'_end');if(!end)throw Error('텍스트상자 끝을 저장할 수 없습니다.');
   const finish=ancestor(end,'p')!,children:Element[]=[];let next=start.nextElementSibling;
   while(next && next!==finish) {const following=next.nextElementSibling;children.push(next);next=following;}
   if(next!==finish)throw Error('텍스트상자 구조가 올바르지 않습니다.');
   const f=record.frame,emu=(v:number)=>String(Math.round(v*EMU));
   const gradient=f.gradient ? `<a:gradFill><a:gsLst>${f.gradient.map(s=>`<a:gs pos="${Math.round(s.position*100000)}"><a:srgbClr val="${s.color.slice(1)}"/></a:gs>`).join('')}</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>`:f.fill?`<a:solidFill><a:srgbClr val="${f.fill.slice(1)}"/></a:solidFill>`:'<a:noFill/>';
   const xml=`<w:p xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr><w:r><w:drawing><wp:anchor distT="${emu(f.wrapMargins?.[0]||0)}" distB="${emu(f.wrapMargins?.[2]||0)}" distL="${emu(f.wrapMargins?.[3]||0)}" distR="${emu(f.wrapMargins?.[1]||0)}" simplePos="0" relativeHeight="${100000+index}" behindDoc="${f.behind?1:0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${f.horizontalReference||'page'}">${f.horizontalReference && f.horizontalAlign && f.horizontalAlign!=='left'?`<wp:align>${f.horizontalAlign}</wp:align>`:`<wp:posOffset>${emu(f.horizontalOffset??f.x)}</wp:posOffset>`}</wp:positionH><wp:positionV relativeFrom="${f.vertical}"><wp:posOffset>${emu(f.y)}</wp:posOffset></wp:positionV><wp:extent cx="${emu(f.width)}" cy="${emu(f.height)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${f.wrap==='square'?'<wp:wrapSquare wrapText="bothSides"/>':'<wp:wrapNone/>'}<wp:docPr id="${100000+index}" name="Text Box ${index+1}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(f.width)}" cy="${emu(f.height)}"/></a:xfrm><a:prstGeom prst="${f.shape}"><a:avLst/></a:prstGeom>${gradient}<a:ln w="${emu(f.borderWidth||0)}">${f.border?`<a:solidFill><a:srgbClr val="${f.border.slice(1)}"/></a:solidFill>`:'<a:noFill/>'}</a:ln></wps:spPr><wps:txbx><w:txbxContent/></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="${f.noWrap?'none':'square'}" lIns="${emu(f.padding[3])}" tIns="${emu(f.padding[0])}" rIns="${emu(f.padding[1])}" bIns="${emu(f.padding[2])}" anchor="${f.align==='center'?'ctr':f.align==='bottom'?'b':'t'}">${f.autoFit?'<a:spAutoFit/>':'<a:noAutofit/>'}</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
   const parsed=new DOMParser().parseFromString(xml,'application/xml');if(parsed.getElementsByTagName('parsererror').length)throw Error('텍스트상자 XML을 만들 수 없습니다.');
   const p=doc.importNode(parsed.documentElement,true),target=all(p,'txbxContent')[0];target.append(...children);
   if(f.position==='inline'){const anchor=p.getElementsByTagNameNS(WP,'anchor')[0],inline=doc.createElementNS(WP,'wp:inline');for(const c of [...anchor.children])if(['extent','effectExtent','docPr','cNvGraphicFramePr','graphic'].includes(c.localName))inline.append(c);anchor.replaceWith(inline);}
   if(f.drawingGroup)p.getElementsByTagNameNS(WP,'anchor')[0].setAttribute('data-ws-drawing-group',JSON.stringify(f.drawingGroup));
   start.replaceWith(p);finish.remove();written.push({record,paragraph:p});
  }
  // A Word group is one anchor. Splitting its children into separate anchors
  // makes Word's web editor give each large child another page of its own.
  const groups=new Map<string,typeof written>();
  for(const item of written)if(item.record.frame.group && !item.record.frame.drawingGroup){const items=groups.get(item.record.frame.group)||[];items.push(item);groups.set(item.record.frame.group,items);}
  for(const members of groups.values()) {
    if(members.length<2 || members.some((m,i)=>m.record.frame.position!=='floating' || m.record.frame.vertical!==members[0].record.frame.vertical || i>0 && members[i-1].paragraph.nextElementSibling!==m.paragraph))continue;
    const x=Math.min(...members.map(m=>m.record.frame.x)),y=Math.min(...members.map(m=>m.record.frame.y));
    const width=Math.max(...members.map(m=>m.record.frame.x+m.record.frame.width))-x,height=Math.max(...members.map(m=>m.record.frame.y+m.record.frame.height))-y;
    const firstParagraph=members[0].paragraph,anchor=firstParagraph.getElementsByTagNameNS(WP,'anchor')[0],graphic=anchor.getElementsByTagNameNS(A,'graphicData')[0];
    const element=(ns:string,name:string,attributes:Record<string,string>={})=>{const el=doc.createElementNS(ns,name);for(const[k,v]of Object.entries(attributes))el.setAttribute(k,v);return el;};
    const emu=(v:number)=>String(Math.round(v*EMU));
    const group=element(WPG,'wpg:wgp'),props=element(WPG,'wpg:grpSpPr'),transform=element(A,'a:xfrm');
    transform.append(element(A,'a:off',{x:'0',y:'0'}),element(A,'a:ext',{cx:emu(width),cy:emu(height)}),element(A,'a:chOff',{x:'0',y:'0'}),element(A,'a:chExt',{cx:emu(width),cy:emu(height)}));
    props.append(transform);group.append(element(WPG,'wpg:cNvGrpSpPr'),props);
    for(const {record,paragraph} of members){const shape=paragraph.getElementsByTagNameNS(WPS,'wsp')[0],off=first(shape,'spPr') && first(first(shape,'spPr'),'xfrm') && first(first(first(shape,'spPr'),'xfrm'),'off');if(off){off.setAttribute('x',emu(record.frame.x-x));off.setAttribute('y',emu(record.frame.y-y));}group.append(shape);if(paragraph!==members[0].paragraph)paragraph.remove();}
    graphic.replaceChildren(group);graphic.setAttribute('uri',WPG);
    const h=first(anchor,'positionH')!,v=first(anchor,'positionV')!;h.setAttribute('relativeFrom','page');h.replaceChildren(element(WP,'wp:posOffset'));h.firstElementChild!.textContent=emu(x);v.replaceChildren(element(WP,'wp:posOffset'));v.firstElementChild!.textContent=emu(y);
    const extent=first(anchor,'extent')!;extent.setAttribute('cx',emu(width));extent.setAttribute('cy',emu(height));
    first(anchor,'docPr')!.setAttribute('name','Text Box Group');
  }
  zip.file(name,new XMLSerializer().serializeToString(doc));
 }
}
