const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url');
let pending;
module.exports=()=>pending??=(async()=>{
 const dir=path.dirname(require.resolve('@rhwp/core')),{default:init,HwpDocument}=await import(pathToFileURL(dir+'/rhwp.js'));
 await init({module_or_path:await fs.readFile(dir+'/rhwp_bg.wasm')});const samples={};
 for(const placement of['body','cell','nested']){
  const text=(placement==='nested'?'😀':'')+'앞  본문 뒤  끝';
  const d=HwpDocument.createEmpty();
  try{d.createBlankDocument();const p=`<p style="text-align:left;line-height:26px;margin:0"><span style="font-family:Arial;font-size:15pt">${text}</span></p>`;
   const inner=`<table style="width:400px"><tr><td style="padding:0">${p}</td></tr></table>`;
   d.pasteHtml(0,0,0,'<p>LEAD</p>'+(placement==='body'?p:placement==='cell'?inner:`<table><tr><td>${inner}</td></tr></table>`));
   for(const format of['hwp','hwpx']){const bytes=format==='hwp'?d.exportHwp():d.exportHwpx();const source=new HwpDocument(bytes);
    try{const runs=JSON.parse(source.getPageTextLayout(0)).runs.filter(r=>r.text?.includes('\u2007'));const widths=runs.flatMap(r=>[...r.text].flatMap((c,i)=>c==='\u2007'?[{font:r.fontSize,width:r.charX[i+1]-r.charX[i]}]:[]));
     if(widths.length!==4||widths.some(w=>Math.abs(w.width-w.font/4)>.11))throw Error('Native fwSpace must paint at a quarter em: '+JSON.stringify(widths));
     samples[`${format}-${placement}`]={bytes,text,spaces:4,widths};
    }finally{source.free()}}
  }finally{d.free()}
 }
 return samples;
})();
