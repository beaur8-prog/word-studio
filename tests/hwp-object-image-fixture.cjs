const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url'),{createRequire}=require('node:module'),{unzipSync,zipSync,strFromU8,strToU8}=require('fflate'),ts=require('typescript');
let ready;
function load(file,cache=new Map()){if(cache.has(file))return cache.get(file).exports;const m={exports:{}};cache.set(file,m);const req=createRequire(file);new Function('require','module','exports',ts.transpileModule(require('node:fs').readFileSync(file,'utf8'),{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText)(name=>name.startsWith('.')&&!path.extname(name)?load(path.resolve(path.dirname(file),name+'.ts'),cache):req(name),m,m.exports);return m.exports;}
module.exports=async function fixture(options={}){
 ready??=(async()=>{const dir=path.dirname(require.resolve('@rhwp/core')),{default:init,HwpDocument}=await import(pathToFileURL(dir+'/rhwp.js'));await init({module_or_path:await fs.readFile(dir+'/rhwp_bg.wasm')});return HwpDocument;})();
 const HwpDocument=await ready,sharp=require('sharp'),pixels=Buffer.alloc(200*80*4);
 for(let y=0;y<80;y++)for(let x=0;x<200;x++){const at=(y*200+x)*4;pixels[at+(x<100?0:2)]=255;pixels[at+3]=255;}
 const png=await sharp(pixels,{raw:{width:200,height:80,channels:4}}).png().toBuffer(),d=HwpDocument.createEmpty(),para=options.pureImage?1:0;let canonical;
 try{
  d.createBlankDocument();d.insertText(0,0,0,'BODY OUTSIDE GROUP');
  if(options.pureImage)d.splitParagraph(0,0,d.getParagraphLength(0,0));
  const title=JSON.parse(d.createShapeControl(JSON.stringify({sectionIdx:0,paraIdx:para,charOffset:0,shapeType:'textbox',width:18000,height:3600,treatAsChar:false,textWrap:'InFrontOfText'})));
  d.insertTextInCell(0,title.paraIdx,title.controlIdx,0,0,0,options.pureImage?'     ':'TITLE 한글\t7');
  d.setShapeProperties(0,title.paraIdx,title.controlIdx,JSON.stringify({horzRelTo:'Paper',vertRelTo:'Paper',horzOffset:9000,vertOffset:7200,borderColor:0,borderWidth:options.pureImage?0:75,fillType:'solid',fillBgColor:0xffffff}));
  d.insertPicture(0,para,0,'[]',png,12000,4800,200,80,'png','RED BLUE LOGO');
  const picture=JSON.parse(d.getControls()).find(c=>c.list===0&&c.userDesc==='그림');
  if(!picture)throw Error('Missing native picture');
  const pp=JSON.parse(d.getPictureProperties(0,picture.para,picture.controlIndex));
  d.setPictureProperties(0,picture.para,picture.controlIndex,JSON.stringify({treatAsChar:false,textWrap:'InFrontOfText',horzRelTo:'Paper',vertRelTo:'Paper',horzOffset:9000,vertOffset:11400,width:6000,height:4800,cropRight:pp.originalWidth/2}));
  const targets=JSON.parse(d.getControls()).filter(c=>c.list===0&&c.ctrlId==='gso').map(c=>({paraIdx:c.para,controlIdx:c.controlIndex}));
  const group=JSON.parse(d.groupShapes(JSON.stringify({sectionIdx:0,targets})));if(!group.ok)throw Error('Could not group image and textbox');
  d.setShapeProperties(0,group.paraIdx,group.controlIdx,JSON.stringify({treatAsChar:true}));d.reflowLinesegs();
  const parts=unzipSync(d.exportHwpx());
  if(options.pureImage){
   let xml=strFromU8(parts['Contents/section0.xml']).replace(/<hp:lineShape\b[^>]*>/g,tag=>tag.replace(/style="[^"]*"/,'style="NONE"'));
   const {elementRanges}=load(path.join(__dirname,'../src/hwp/markup-ranges.ts'));
   const frame=JSON.parse(d.getPageControlLayout(0)).controls.find(c=>c.stableIndex?.[1]===para&&c.stableIndex?.[2]===group.controlIdx&&['group','shape'].includes(c.type));
   const old=JSON.parse(d.getParagraphLineMetrics(0,para,'[]')).lines[0];
   if(!frame||!(frame.h>0))throw Error('Missing source group frame');
   // Editing APIs leave this synthetic object's old 10pt host cache intact.
   // Define an explicit one-object fixture line from its measured own frame,
   // rather than testing top alignment with an invalid 13px/120px cache.
   const units=n=>Math.round(n*75),cache=`<hp:linesegarray><hp:lineseg textpos="0" vertpos="${units(old.verticalPosition)}" vertsize="${units(frame.h)}" textheight="${units(frame.h)}" baseline="${units(frame.h*.85)}" spacing="${units(old.spacing)}" horzpos="0" horzsize="${units(old.width)}" flags="393216"/></hp:linesegarray>`;
   const host=elementRanges(xml,'hp:p').find(p=>p.inner.includes('<hp:container'));
   if(!host)throw Error('Missing fixture group host');
   const own=elementRanges(host.inner,'hp:linesegarray').at(-1);
   if(!own||own.start<host.inner.lastIndexOf('</hp:container>'))throw Error('Missing own fixture line cache');
   const inside=host.inner.slice(0,own.start)+cache+host.inner.slice(own.end);
   xml=xml.slice(0,host.start)+host.open+inside+'</hp:p>'+xml.slice(host.end);parts['Contents/section0.xml']=strToU8(xml);
  }
  if(options.pageHide)parts['Contents/section0.xml']=strToU8(strFromU8(parts['Contents/section0.xml']).replace('</hp:container>','</hp:container><hp:ctrl><hp:pageHiding hideHeader="1" hideFooter="1" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="1"/></hp:ctrl>'));
  let header=strFromU8(parts['Contents/header.xml']);
  const tab='<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"><hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"><hh:tabItem pos="14400" type="RIGHT" leader="DASH"/></hp:case><hp:default><hh:tabItem pos="28800" type="RIGHT" leader="DASH"/></hp:default></hp:switch></hh:tabPr>';
  const before=header;header=header.replace(/<hh:tabPr id="0"[^>]*\/>/,tab);if(header===before)throw Error('Missing fixture tab definition');parts['Contents/header.xml']=strToU8(header);canonical=zipSync(parts,{level:1});
 }finally{d.free();}
 const {editableHwp}=load(path.join(__dirname,'../src/hwp/conversion.ts')),result={};
 for(const format of ['hwp','hwpx']){
  let source=new HwpDocument(canonical),bytes;
  try{bytes=format==='hwp'?source.exportHwp():canonical;}finally{source.free();}
  source=new HwpDocument(bytes);
  try{const tree=JSON.parse(source.getPageLayerTreeWithProfile(0,'print',true,true)),ops=[];const walk=n=>{for(const o of n.ops||[])ops.push(o);n.children?.forEach(walk);if(n.child)walk(n.child);};walk(tree.root);
   const groupHeight=JSON.parse(source.getParagraphLineMetrics(0,para,'[]')).lines?.[0]?.height;
   result[format]={bytes,editable:editableHwp(source),ops,groupHeight,sourceSvg:source.renderPageSvgWithProfile(0,'print')};
  }finally{source.free();}
 }
 return result;
};
