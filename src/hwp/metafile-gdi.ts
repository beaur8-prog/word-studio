type Rect = [number, number, number, number];
type Brush = string|null|{ color: string; definition: string }|undefined;
type State = { brush: Brush; align: number; clip: Rect|null };
type Paint = { kind: 'line'|'text'; text?: string; align?: number }|{ kind: 'fill'; rect: Rect; color: string };
const intersect = (a: Rect, b: Rect): Rect => [Math.max(a[0],b[0]),Math.max(a[1],b[1]),Math.min(a[2],b[2]),Math.min(a[3],b[3])];
const xmlText = (s: string) => s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, v: string) =>
  v[0] === '#' ? String.fromCodePoint(parseInt(v.slice(v[1].toLowerCase()==='x'?2:1),v[1].toLowerCase()==='x'?16:10)) : ({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"}[v.toLowerCase()]!));

/** Restore two GDI operations omitted by the native SVG bridge: TA_TOP text
 * reference points and solid-brush PATCOPY cell backgrounds. Keep native line
 * geometry, clipping and the outer device-to-SVG matrix. Apply atomically only
 * to a verified plain GDI stream whose every paint matches one native group.
 * Unknown records/transforms never receive guessed coordinates or colours.
 * MS-EMF2.2.5/2.3.1.2; Win32 SetTextAlign defaults to TA_LEFT|TA_TOP. */
export function restoreMetafileGdi(svg: string, bytes: Uint8Array): string {
  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const paints: Paint[] = [], brushes = new Map<number,Brush>();
    let state: State = { brush: 'rgb(255,255,255)', align: 0, clip: null };
    const stack: State[] = [], patterns: string[] = [];
    let patternPixels = 0;
    const u = (p: number) => view.getUint32(p,true), i = (p: number) => view.getInt32(p,true);
    const rect = (p: number): Rect => [i(p),i(p+4),i(p+8),i(p+12)];
    const color = (n: number) => `rgb(${n&255},${(n>>>8)&255},${(n>>>16)&255})`;
    let offset = 0, records = 0;
    while (offset + 8 <= bytes.length) {
      if (++records > 20000) return svg;
      const type = u(offset), size = u(offset+4), at = offset;
      if (size < 8 || size % 4 || offset + size > bytes.length) return svg;
      const need = (length: number) => { if(size<length) throw Error('Truncated EMF record'); };
      switch(type) {
        case 1: case 14: case 18: case 20: case 24: case 25: case 27: case 70: break;
        case 10: need(16); if(i(at+8)||i(at+12)) return svg; break;
        case 17: need(12); if(u(at+8)!==1) return svg; break;
        case 22: need(12); state.align=u(at+8); break;
        case 33:
          if(stack.length>=256) return svg;
          stack.push({...state,clip:state.clip&&[...state.clip]}); break;
        case 34:
          need(12); if(i(at+8)!==-1||!stack.length) return svg;
          state=stack.pop()!; break;
        case 30: { need(24); const r=rect(at+8); state.clip=state.clip?intersect(state.clip,r):r; break; }
        case 75:
          need(16); if(u(at+8)!==0||u(at+12)!==5) return svg;
          state.clip=null; break;
        case 39:
          need(24); if(u(at+12)>1) return svg;
          brushes.set(u(at+8),u(at+12)===1?null:color(u(at+16))); break;
        case 94: {
          need(32); const id=u(at+8),usage=u(at+12),bmi=u(at+16),info=u(at+20),bits=u(at+24),length=u(at+28);
          brushes.set(id,undefined);
          if(usage||bmi<32||info<40||bmi+info>size||bits<32||bits+length>size) break;
          const w=i(at+bmi+4),rawHeight=i(at+bmi+8),h=Math.abs(rawHeight),compression=u(at+bmi+16);
          if(w<1||w>16||h<1||h>16||view.getUint16(at+bmi+12,true)!==1||view.getUint16(at+bmi+14,true)!==32||length<w*h*4) break;
          if(compression!==0&&(compression!==3||info<52||u(at+bmi+40)!==0xff0000||u(at+bmi+44)!==0xff00||u(at+bmi+48)!==0xff)) break;
          if((patternPixels+=w*h)>4096) return svg;
          const key=`studio-emf-pattern-${at}`;let pixels='';
          for(let y=0;y<h;y++)for(let x=0;x<w;x++) {
            const p=at+bits+((rawHeight>0?h-1-y:y)*w+x)*4;
            pixels+=`<rect x="${x}" y="${y}" width="1" height="1" fill="rgb(${bytes[p+2]},${bytes[p+1]},${bytes[p]})"/>`;
          }
          const definition=`<pattern id="${key}" patternUnits="userSpaceOnUse" width="${w}" height="${h}">${pixels}</pattern>`;
          brushes.set(id,{color:`url(#${key})`,definition}); patterns.push(definition); break;
        }
        case 37: {
          need(12); const id=u(at+8);
          if(brushes.has(id)) state.brush=brushes.get(id);
          else if(id>=0x80000000&&id<=0x80000005) state.brush=[
            'rgb(255,255,255)','rgb(192,192,192)','rgb(128,128,128)','rgb(64,64,64)','rgb(0,0,0)',null,
          ][id-0x80000000];
          else if(id===0x80000012) return svg; // DC_BRUSH colour is outside this subset.
          break;
        }
        case 38: case 40: need(12); brushes.delete(u(at+8)); break;
        case 82:
          need(104); brushes.delete(u(at+8));
          if(i(at+20)||i(at+24)||view.getUint16(at+40,true)===64) return svg; // Rotated/vertical font.
          break;
        case 54: need(16); paints.push({kind:'line'}); break;
        case 84: {
          need(76); const count=u(at+44), start=u(at+48), options=u(at+52);
          if(!count) break;
          if(options&0x10||start<76||start+count*2>size) return svg; // Glyph indices are not Unicode.
          const text=new TextDecoder('utf-16le').decode(bytes.subarray(at+start,at+start+count*2));
          if(text) paints.push({kind:'text',text,align:state.align});
          break;
        }
        case 76: {
          need(100); if(u(at+40)!==0x00f00021||u(at+88)||u(at+96)||state.brush===undefined) return svg;
          if(state.brush!==null) {
            const x=i(at+24),y=i(at+28),w=i(at+32),h=i(at+36);
            let r: Rect=[Math.min(x,x+w),Math.min(y,y+h),Math.max(x,x+w),Math.max(y,y+h)];
            if(state.clip) r=intersect(r,state.clip);
            if(r[2]>r[0]&&r[3]>r[1]) paints.push({kind:'fill',rect:r,color:typeof state.brush==='string'?state.brush:state.brush.color});
          }
          break;
        }
        default: return svg;
      }
      offset += size;
    }
    if(offset!==bytes.length||!paints.length) return svg;
    const opening=svg.match(/^<svg\b[^>]*><g\b[^>]*>/), ending='</g></svg>';
    if(!opening||!svg.endsWith(ending)) return svg;
    const inner=svg.slice(opening[0].length,-ending.length), children: string[]=[];
    let depth=0,start=-1;
    for(const match of inner.matchAll(/<\/?[a-zA-Z][^>]*>/g)) {
      const tag=match[0],closing=tag.startsWith('</'),self=tag.endsWith('/>');
      if(depth===0&&!closing) start=match.index!;
      if(closing) depth--; else if(!self) depth++;
      if(depth<0) return svg;
      if(depth===0&&start>=0) { children.push(inner.slice(start,match.index!+tag.length));start=-1; }
    }
    if(depth||start>=0) return svg;
    const groups: {prefix:string;group:string}[]=[];let prefix='';
    for(const child of children) {
      if(/^<(?:g|text|line)\b/.test(child)) { groups.push({prefix,group:child});prefix=''; }
      else if(/^<clipPath\b/.test(child)) prefix+=child;
      else return svg;
    }
    if(groups.length!==paints.filter(p=>p.kind!=='fill').length) return svg;
    let index=0, result='';
    for(const paint of paints) {
      if(paint.kind==='fill') {
        const [x,y,right,bottom]=paint.rect;
        result+=`<rect x="${x}" y="${y}" width="${right-x}" height="${bottom-y}" fill="${paint.color}" stroke="none"/>`;
      } else {
        const item=groups[index++];let group=item.group;
        if(paint.kind==='line') { if(!/<line\b/.test(group)||/<text\b/.test(group)) return svg; }
        else {
          const text=group.match(/<text\b[^>]*>([^<]*)<\/text>/);
          if(!text||xmlText(text[1])!==paint.text) return svg;
          if((paint.align!&24)===0&&!/\bdominant-baseline=/.test(group))
            group=group.replace('<text ','<text dominant-baseline="text-before-edge" ');
        }
        result+=item.prefix+group;
      }
    }
    const complete=opening[0]+(patterns.length?'<defs>'+patterns.join('')+'</defs>':'')+result+prefix+ending;
    return complete.length>4*1024*1024?svg:complete;
  } catch { return svg; }
}
