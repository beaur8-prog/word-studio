const fs=require('node:fs/promises'),fss=require('node:fs'),path=require('node:path'),{createRequire}=require('node:module'),{pathToFileURL}=require('node:url');
const ts=require('typescript'),{unzipSync,zipSync,strFromU8,strToU8}=require('fflate');
let pending;
function load(file,cache=new Map()) {
 if(cache.has(file))return cache.get(file).exports;
 const m={exports:{}},native=createRequire(file);cache.set(file,m);
 new Function('require','module','exports',ts.transpileModule(fss.readFileSync(file,'utf8'),{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText)(name=>name.startsWith('.')&&!path.extname(name)?load(path.resolve(path.dirname(file),name+'.ts'),cache):native(name),m,m.exports);
 return m.exports;
}
module.exports=()=>pending??=(async()=>{
 const dir=path.dirname(require.resolve('@rhwp/core')),{default:init,HwpDocument}=await import(pathToFileURL(dir+'/rhwp.js'));
 await init({module_or_path:await fs.readFile(dir+'/rhwp_bg.wasm')});
 const root=process.env.HWP_RANGE_BOOKMARK_SOURCE_ROOT||path.resolve(__dirname,'..'),{editableHwp}=load(root+'/src/hwp/conversion.ts'),samples={};
 for(const nested of [false,true]) {
  const d=HwpDocument.createEmpty();
  try {
   d.createBlankDocument();
   const cell='<table><tr><td><p>CELL PRE BOOKMARK_CELL CELL TAIL</p></td></tr></table>';
   d.pasteHtml(0,0,0,'<p>BODY PRE BOOKMARK_BODY BODY TAIL</p>'+(nested?'<table><tr><td>'+cell+'</td></tr></table>':cell)+'<p>END</p>');
   const files=unzipSync(d.exportHwpx());let xml=strFromU8(files['Contents/section0.xml']);
   for(const[label,id,name]of[['BODY',1001,'정부 범위 BODY'],['CELL',1002,'정부 범위 CELL']]) {
    const token='BOOKMARK_'+label;let count=0;
    xml=xml.replace(/<hp:t>([^<]*)<\/hp:t>/g,(tag,text)=>{
     if(!text.includes(token))return tag;count++;
     const at=text.indexOf(token);
     return `<hp:t>${text.slice(0,at)}</hp:t><hp:ctrl><hp:fieldBegin id="${id}" type="BOOKMARK" name="${name}" editable="0" dirty="1"><hp:parameters cnt="1" name=""><hp:integerParam name="Prop">2</hp:integerParam></hp:parameters></hp:fieldBegin></hp:ctrl><hp:t>${token}</hp:t><hp:ctrl><hp:fieldEnd beginIDRef="${id}"/></hp:ctrl><hp:t>${text.slice(at+token.length)}</hp:t>`;
    });
    if(count!==1)throw Error('Range bookmark fixture text is not unique');
   }
   files['Contents/section0.xml']=strToU8(xml);const hwpx=zipSync(files,{level:1}),canonical=new HwpDocument(hwpx);
   try {
    for(const format of ['hwp','hwpx']) {
     const bytes=format==='hwp'?canonical.exportHwp():hwpx,id=format+'-'+(nested?'nested':'ordinary'),source=new HwpDocument(bytes);
     try{samples[id]={bytes,format,nested,sourcePages:source.pageCount(),editable:editableHwp(source)};}finally{source.free();}
    }
   }finally{canonical.free();}
   if(!nested)for(const mode of ['unknown','invalid-name']) {
    const parts={...files};const altered=mode==='unknown'?xml.replace(/type="BOOKMARK"/g,'type="UNKNOWN"'):xml.replace(/name="정부 범위 (?:BODY|CELL)"/g,'name="'+'X'.repeat(1001)+'"');
    parts['Contents/section0.xml']=strToU8(altered);const bytes=zipSync(parts,{level:1}),source=new HwpDocument(bytes);
    try{samples[mode]={bytes,format:'hwpx',negative:true,sourcePages:source.pageCount(),editable:editableHwp(source)};}finally{source.free();}
   }
  }finally{d.free();}
 }
 return samples;
})();
