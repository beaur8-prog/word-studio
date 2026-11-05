/** Bounded, resolved group geometry; no source XML crosses the model boundary. */
export type DrawingGroup = {
 id:string; order:number; x:number; y:number; width:number; height:number;
 vertical:'page'|'paragraph'; wrap:{type:0|1|2|3|4;side?:'bothSides'|'left'|'right'|'largest'};
 margins:[number,number,number,number]; behind:boolean; polygon?:[number,number][];
};
export function drawingGroup(value:unknown):DrawingGroup|undefined {
 if(!value || typeof value!=='object')return;const v=value as DrawingGroup;
 const bounded=(n:unknown,low:number,high:number)=>typeof n==='number' && Number.isFinite(n) && n>=low && n<=high;
 if(typeof v.id!=='string' || !/^[a-z0-9_]{1,64}$/i.test(v.id) || !Number.isInteger(v.order) || !bounded(v.order,0,10000) ||
  !bounded(v.x,-100000,100000) || !bounded(v.y,-100000,100000) || !bounded(v.width,.1,100000) || !bounded(v.height,.1,100000) ||
  !['page','paragraph'].includes(v.vertical) || ![0,1,2,3,4].includes(v.wrap?.type) || typeof v.behind!=='boolean' ||
  !Array.isArray(v.margins) || v.margins.length!==4 || !v.margins.every(n=>bounded(n,0,10000)))return;
 if(v.polygon!==undefined && (!Array.isArray(v.polygon) || v.polygon.length<3 || v.polygon.length>256 || !v.polygon.every(p=>Array.isArray(p) && p.length===2 && p.every(n=>bounded(n,-4,4)))))return;
 return {id:v.id,order:v.order,x:v.x,y:v.y,width:v.width,height:v.height,vertical:v.vertical,
  wrap:{type:v.wrap.type,...(['bothSides','left','right','largest'].includes(v.wrap.side||'')?{side:v.wrap.side}:{})},margins:[...v.margins],behind:v.behind,...(v.polygon?{polygon:v.polygon.map(p=>[...p] as [number,number])}:{})};
}
