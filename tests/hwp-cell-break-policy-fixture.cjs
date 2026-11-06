const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url'),{createRequire}=require('node:module'),ts=require('typescript'),{unzipSync,zipSync,strFromU8,strToU8}=require('fflate');
let pending;
function load(file,cache=new Map()){
 if(cache.has(file))return cache.get(file).exports;
 const m={exports:{}},req=createRequire(file);cache.set(file,m);
 new Function('require','module','exports',ts.transpileModule(require('node:fs').readFileSync(file,'utf8'),{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText)(n=>n.startsWith('.')&&!path.extname(n)?load(path.resolve(path.dirname(file),n+'.ts'),cache):req(n),m,m.exports);return m.exports;
}
module.exports=()=>pending??=(async()=>{
 const core=path.dirname(require.resolve('@rhwp/core')),{default:init,HwpDocument}=await import(pathToFileURL(core+'/rhwp.js'));await init({module_or_path:await fs.readFile(core+'/rhwp_bg.wasm')});
 const {editableHwp}=load(path.join(process.env.HWP_CELL_BREAK_SOURCE_ROOT||path.resolve(__dirname,'..'),'src/hwp/conversion.ts')),result={};
 for(const nested of [false,true])for(const inline of [false,true])for(const policy of inline?[2]:[0,1,2]){
  const d=HwpDocument.createEmpty();
  try{
   d.createBlankDocument();
   const lines=label=>Array.from({length:24},(_,i)=>`<p style="font-family:Arial;font-size:10pt;line-height:20px;margin:0">${label} ${i}</p>`).join('');
   const inner='<table style="width:400px"><tr><td style="width:200px;height:520px;padding:0">'+lines('LEFT')+'</td><td style="width:200px;height:520px;padding:0">'+lines('RIGHT')+'</td></tr></table>';
   d.pasteHtml(0,0,0,'<p>LEAD</p>'+(nested?'<table style="width:420px"><tr><td style="padding:0">'+inner+'</td></tr></table>':inner)+'<p>END</p>');
   const c=JSON.parse(d.getControls()).find(c=>c.list===0&&c.ctrlId==='tbl');
   d.setTableProperties(0,c.para,c.controlIndex,JSON.stringify({pageBreak:nested?2:policy,repeatHeader:false,treatAsChar:nested?false:inline,textWrap:'TopAndBottom',vertRelTo:'Para',vertOffset:0}));
   for(let cell=0;cell<(nested?1:2);cell++)d.setCellProperties(0,c.para,c.controlIndex,cell,JSON.stringify({verticalAlign:0}));
   const parts=unzipSync(d.exportHwpx());let xml=strFromU8(parts['Contents/section0.xml']);
   if(nested){let ordinal=0;xml=xml.replace(/<hp:tbl\b[^>]*>/g,tag=>ordinal++===1?tag.replace(/pageBreak="[^"]*"/,'pageBreak="'+['NONE','TABLE','CELL'][policy]+'"'):tag);if(ordinal!==2)throw Error('Fixture must contain exactly two source tables');let table=-1;xml=xml.replace(/<hp:tbl\b[^>]*>|<hp:pos\b[^>]*>/g,tag=>{if(tag.startsWith('<hp:tbl')){table++;return tag;}return table===1?tag.replace(/treatAsChar="[^"]*"/,'treatAsChar="'+(inline?'1':'0')+'"'):tag;});parts['Contents/section0.xml']=strToU8(xml);}
   const hwpx=zipSync(parts,{level:1}),canonical=new HwpDocument(hwpx);
   try {
   for(const format of ['hwp','hwpx']){
    const bytes=format==='hwp'?canonical.exportHwp():hwpx,s=new HwpDocument(bytes);
    try{const t=JSON.parse(s.getControls()).find(c=>c.list===0&&c.ctrlId==='tbl'),boxes=JSON.parse(s.getTableCellBboxes(0,t.para,t.controlIndex));
     if(s.pageCount()!==1||new Set(boxes.map(b=>b.pageIndex)).size!==1)throw Error('The source fixture must fit one native page before an edit');
     result[`${format}-${policy}-${nested?'nested':'ordinary'}${inline?'-inline':''}`]={bytes,editable:editableHwp(s),sourcePages:s.pageCount(),sourceCellPages:new Set(boxes.map(b=>b.pageIndex)).size,policy:JSON.parse(s.getTableProperties(0,t.para,t.controlIndex)).pageBreak,xmlPolicy:[...xml.matchAll(/<hp:tbl\b[^>]*>/g)][nested?1:0][0].match(/pageBreak="([^"]*)"/)[1],nested,inline};
    }finally{s.free();}
   }
   }finally{canonical.free();}
  }finally{d.free();}
 }
 return result;
})();
