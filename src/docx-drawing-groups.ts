import type JSZip from 'jszip';
import {W,all,ancestor} from './docx-format';
import {drawingFloating} from './image-floating';
import {drawingGroup,type DrawingGroup} from './drawing-group';
import {pageMargins,paperDimensions,type PageSettings} from './types';
const A='http://schemas.openxmlformats.org/drawingml/2006/main',WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',WPG='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',WPS='http://schemas.microsoft.com/office/word/2010/wordprocessingShape',PIC='http://schemas.openxmlformats.org/drawingml/2006/picture';
const EMU=9525,first=(e:Element|null|undefined,n:string)=>e && [...e.children].find(c=>c.localName===n);
const number=(e:Element|null|undefined,k:string,f=0)=>{const n=Number(e?.getAttribute(k)??f);return Number.isFinite(n)?n:f;};
const emu=(n:number)=>String(Math.round(n*EMU));
function position(anchor:Element,name:string,vertical:boolean,page:PageSettings) {
 let p=first(anchor,name);
 if(!p){outer:for(const wrapper of anchor.children)if(wrapper.localName==='AlternateContent')for(const branch of wrapper.children){const c=first(branch,name);if(c){p=c;break outer;}}}
 const margins=pageMargins(page),length=paperDimensions(page)[vertical?1:0]*96/25.4,relative=p?.getAttribute('relativeFrom')||'page';
 const start=(vertical?margins.top:margins.left)*96/25.4,end=(vertical?margins.bottom:margins.right)*96/25.4,isMargin=['margin','column'].includes(relative),base=isMargin?start:0,available=length-(isMargin?start+end:0);
 const size=number(first(anchor,'extent'),vertical?'cy':'cx')/EMU,align=first(p,'align')?.textContent,pct=first(p,vertical?'pctPosVOffset':'pctPosHOffset'),offset=first(p,'posOffset');
 return {relative,value:pct?base+available*Number(pct.textContent)/100000:offset?base+Number(offset.textContent)/EMU:align==='center'?base+(available-size)/2:['right','bottom','outside'].includes(align||'')?base+available-size:base};
}
function shapeGeometry(shape:Element,group:Element) {
 const t=first(first(shape,'spPr'),'xfrm'),off=first(t,'off'),ext=first(t,'ext');
 let x=number(off,'x'),y=number(off,'y'),width=number(ext,'cx'),height=number(ext,'cy'),parent=shape.parentElement;
 while(parent){
  if(parent.localName==='wgp' || parent.localName==='grpSp'){
   const p=first(first(parent,'grpSpPr'),'xfrm'),o=first(p,'off'),e=first(p,'ext'),co=first(p,'chOff'),ce=first(p,'chExt');
   const sx=number(e,'cx',1)/number(ce,'cx',number(e,'cx',1)),sy=number(e,'cy',1)/number(ce,'cy',number(e,'cy',1));
   x=number(o,'x')+(x-number(co,'x'))*sx;y=number(o,'y')+(y-number(co,'y'))*sy;width*=sx;height*=sy;
  }
  if(parent===group)break;parent=parent.parentElement;
 }
 return{x:x/EMU,y:y/EMU,width:width/EMU,height:height/EMU};
}
/** Mammoth skips pictures in wpg:wgp. Extract supported members before the
 * textbox pass removes their shared drawing, retaining one parent wrap/anchor. */
export function normalizeDrawingGroups(body:Document,pages:PageSettings[]) {
 const groups=new Map<Element,Element[]>();
 const collect=(root:Element,members:Element[],depth=0)=>{
  if(depth>24)throw Error('중첩 그림 그룹이 너무 깊습니다.');
  for(const c of root.children){
   if(c.namespaceURI===PIC && c.localName==='pic' || c.namespaceURI===WPS && c.localName==='wsp')members.push(c);
   else if(c.namespaceURI===WPG && c.localName==='grpSp')collect(c,members,depth+1);
  }
 };
 for(const group of [...body.getElementsByTagNameNS(WPG,'wgp')]){
  if(ancestor(group,'txbxContent') || ancestor(group,'Fallback'))continue;
  const members:Element[]=[];collect(group,members);
  // A picture in a textbox's own story is not a sibling group member. Keep
  // floating box-only groups on their existing path; nested stories handle themselves.
  const drawing=ancestor(group,'drawing');
  if(members.some(m=>m.namespaceURI===PIC) || drawing && first(drawing,'inline') && members.some(m=>m.namespaceURI===WPS))groups.set(group,members);
 }
 if(!groups.size)return;
 let section=0;const pageFor=new Map<Element,PageSettings>();
 for(const p of all(body,'p')){pageFor.set(p,pages[Math.min(section,pages.length-1)]||pages[0]);if(first(first(p,'pPr'),'sectPr') && !ancestor(p,'txbxContent'))section++;}
 for(const [group,members]of groups) {
  if(!group.isConnected)continue;
  const drawing=ancestor(group,'drawing'),anchor=drawing && (first(drawing,'anchor')||first(drawing,'inline'));if(!drawing || !anchor)continue;
  const page=pageFor.get(ancestor(drawing,'p')!)||pages[0];if(!page)continue;
  const h=position(anchor,'positionH',false,page),v=position(anchor,'positionV',true,page),floating=drawingFloating(drawing);
  const bounds=drawingGroup({...(anchor.localName==='inline'?{position:'inline'}:{}),id:'drawing_group_'+crypto.randomUUID().replaceAll('-',''),order:0,x:h.value,y:v.value,width:number(first(anchor,'extent'),'cx')/EMU,height:number(first(anchor,'extent'),'cy')/EMU,
   vertical:['paragraph','line'].includes(v.relative)?'paragraph':'page',wrap:floating?.wrap||{type:0},margins:[number(anchor,'distT')/EMU,number(anchor,'distR')/EMU,number(anchor,'distB')/EMU,number(anchor,'distL')/EMU],behind:anchor.getAttribute('behindDoc')==='1',...(floating?.polygon?{polygon:floating.polygon}:{})});
  if(!bounds)continue;
  const replacement=body.createDocumentFragment();
  for(const [order,member]of members.entries()){
   const geometry=shapeGeometry(member,group);if(!(geometry.width>0 && geometry.height>0))continue;
   const copy=drawing.cloneNode(false) as Element,a=anchor.localName==='inline'?body.createElementNS(WP,'wp:anchor'):anchor.cloneNode(false) as Element;copy.append(a);
   if(anchor.localName==='inline'){for(const[k,v]of Object.entries({simplePos:'0',relativeHeight:'0',behindDoc:'0',locked:'0',layoutInCell:'1',allowOverlap:'1'}))a.setAttribute(k,v);const simple=body.createElementNS(WP,'wp:simplePos');simple.setAttribute('x','0');simple.setAttribute('y','0');a.append(simple);}
   // Clone only anchor properties, never the entire group per member. Large
   // groups must remain linear in memory rather than duplicating every image.
   for(const c of anchor.children){if(c.localName!=='graphic')a.append(c.cloneNode(true));else{const g=c.cloneNode(false) as Element,d=first(c,'graphicData')!;g.append(d.cloneNode(false));a.append(g);}}
   const graphic=first(first(a,'graphic'),'graphicData')!;
   const shape=member.cloneNode(true) as Element;graphic.replaceChildren(shape);graphic.setAttribute('uri',member.namespaceURI===PIC?PIC:WPS);
   const transform=first(first(shape,'spPr'),'xfrm'),off=first(transform,'off'),ext=first(transform,'ext');
   if(off){off.setAttribute('x','0');off.setAttribute('y','0');}if(ext){ext.setAttribute('cx',emu(geometry.width));ext.setAttribute('cy',emu(geometry.height));}
   const element=(n:string)=>body.createElementNS(WP,'wp:'+n);
   for(const [name,value,relative]of [['positionH',bounds.x+geometry.x,'page'],['positionV',bounds.y+geometry.y,bounds.vertical]] as const){
    const p=element(name),o=element('posOffset');p.setAttribute('relativeFrom',relative);o.textContent=emu(value);p.append(o);const previous=first(a,name);if(previous)previous.replaceWith(p);else a.insertBefore(p,first(a,'extent')!);
   }
   const extent=first(a,'extent')!;extent.setAttribute('cx',emu(geometry.width));extent.setAttribute('cy',emu(geometry.height));
   for(const c of [...a.children])if(c.localName.startsWith('wrap'))c.remove();a.insertBefore(element('wrapNone'),first(a,'docPr')!);
   for(const side of ['T','R','B','L'])a.setAttribute('dist'+side,'0');
   const properties=first(a,'docPr')!,description=first(first(shape,'nvPicPr'),'cNvPr');
   if(description){properties.setAttribute('name',description.getAttribute('name')||'Group picture');properties.setAttribute('descr',description.getAttribute('descr')||description.getAttribute('title')||'');}
   properties.setAttribute('id',String(900000+order));
   copy.setAttribute('data-ws-drawing-group',JSON.stringify({...bounds,order}));replacement.append(copy);
  }
  if(!replacement.childNodes.length)continue;
  const alternate=ancestor(drawing,'AlternateContent');
  // Remove the selected drawing's legacy copy once, before either importer.
  if(alternate && [...alternate.children].some(c=>c.localName==='Choice' && c.contains(drawing)) && [...alternate.getElementsByTagNameNS(W,'drawing')].filter(d=>!ancestor(d,'Fallback')).length===1)alternate.replaceWith(replacement);
  else drawing.replaceWith(replacement);
 }
}

/** Reassemble flattened editable members into one standard Word anchor. */
export async function writeDrawingGroups(zip:JSZip) {
 for(const name of Object.keys(zip.files).filter(n=>/^word\/.*\.xml$/.test(n))){
  const source=await zip.file(name)!.async('string');if(!source.includes('data-ws-drawing-group'))continue;
  const doc=new DOMParser().parseFromString(source,'application/xml'),groups=new Map<string,{anchor:Element;group:DrawingGroup}[]>();
  for(const anchor of [...doc.getElementsByTagNameNS(WP,'anchor')]){
   const raw=anchor.getAttribute('data-ws-drawing-group');if(!raw)continue;anchor.removeAttribute('data-ws-drawing-group');const group=drawingGroup(JSON.parse(raw));if(!group)throw Error('그림 그룹의 저장 좌표가 올바르지 않습니다.');
   const items=groups.get(group.id)||[];items.push({anchor,group});groups.set(group.id,items);
  }
  const element=(ns:string,name:string,attrs:Record<string,string>={})=>{const e=doc.createElementNS(ns,name);for(const[k,v]of Object.entries(attrs))e.setAttribute(k,v);return e;};
  for(const members of groups.values()){
   const f=members[0].group,base=members[0].anchor,paragraph=ancestor(base,'p')!,parent=paragraph.parentElement;
   if(members.some(m=>ancestor(m.anchor,'p')?.parentElement!==parent))throw Error('그림 그룹의 구성 요소가 서로 다른 영역에 있어 저장할 수 없습니다.');
   const group=element(WPG,'wpg:wgp'),props=element(WPG,'wpg:grpSpPr'),transform=element(A,'a:xfrm');
   transform.append(element(A,'a:off',{x:'0',y:'0'}),element(A,'a:ext',{cx:emu(f.width),cy:emu(f.height)}),element(A,'a:chOff',{x:'0',y:'0'}),element(A,'a:chExt',{cx:emu(f.width),cy:emu(f.height)}));props.append(transform);group.append(element(WPG,'wpg:cNvGrpSpPr'),props);
   for(const {anchor}of [...members].sort((a,b)=>a.group.order-b.group.order)){
    const data=first(first(anchor,'graphic'),'graphicData')!,shape=data.firstElementChild!;
    if(shape.localName==='pic'){const properties=first(anchor,'docPr'),picture=first(first(shape,'nvPicPr'),'cNvPr');if(properties && picture)for(const key of ['name','descr','title']){const value=properties.getAttribute(key);if(value!==null)picture.setAttribute(key,value);}}
    const x=Number(first(first(anchor,'positionH'),'posOffset')?.textContent)/EMU-f.x,y=Number(first(first(anchor,'positionV'),'posOffset')?.textContent)/EMU-f.y;
    const t=first(first(shape,'spPr'),'xfrm'),off=first(t,'off');if(!off)throw Error('그림 그룹의 구성 요소를 저장할 수 없습니다.');off.setAttribute('x',emu(x));off.setAttribute('y',emu(y));group.append(shape);
    if(anchor!==base){const p=ancestor(anchor,'p')!,drawing=ancestor(anchor,'drawing')!,run=drawing.parentElement;drawing.remove();if(run?.localName==='r' && [...run.children].every(c=>c.localName==='rPr'))run.remove();if(shape.localName==='wsp' && [...p.children].every(c=>c.localName==='pPr'))p.remove();}
   }
   const data=first(first(base,'graphic'),'graphicData')!;data.replaceChildren(group);data.setAttribute('uri',WPG);
   for(const [n,value,relative]of [['positionH',f.x,'page'],['positionV',f.y,f.vertical]] as const){const p=first(base,n)!;p.setAttribute('relativeFrom',relative);const off=element(WP,'wp:posOffset');off.textContent=emu(value);p.replaceChildren(off);}
   const ext=first(base,'extent')!;ext.setAttribute('cx',emu(f.width));ext.setAttribute('cy',emu(f.height));first(base,'docPr')!.setAttribute('name','Drawing Group');base.setAttribute('behindDoc',f.behind?'1':'0');
   for(const[k,v]of [['T',f.margins[0]],['R',f.margins[1]],['B',f.margins[2]],['L',f.margins[3]]] as const)base.setAttribute('dist'+k,emu(v));
   for(const c of [...base.children])if(c.localName.startsWith('wrap'))c.remove();
   const kind=['wrapNone','wrapSquare','wrapTight','wrapTopAndBottom','wrapThrough'][f.wrap.type],wrap=element(WP,'wp:'+kind,f.wrap.type===1 || f.wrap.type===2 || f.wrap.type===4?{wrapText:f.wrap.side||'bothSides'}:{});
   if(f.wrap.type===2 || f.wrap.type===4){const polygon=element(WP,'wp:wrapPolygon',{edited:'0'});for(const[i,p]of (f.polygon||[[0,0],[1,0],[1,1],[0,1],[0,0]]).entries())polygon.append(element(WP,i?'wp:lineTo':'wp:start',{x:String(Math.round(p[0]*21600)),y:String(Math.round(p[1]*21600))}));wrap.append(polygon);}
   base.insertBefore(wrap,first(base,'docPr')!);
   if(f.position==='inline'){const inline=element(WP,'wp:inline');for(const c of [...base.children])if(['extent','effectExtent','docPr','cNvGraphicFramePr','graphic'].includes(c.localName))inline.append(c);base.replaceWith(inline);}
  }
  zip.file(name,new XMLSerializer().serializeToString(doc));
 }
}
