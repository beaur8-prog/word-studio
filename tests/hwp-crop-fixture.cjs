const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url');
module.exports=async function(directory){
  const core=path.dirname(require.resolve('@rhwp/core'));
  const{default:init,HwpDocument:C}=await import(pathToFileURL(path.join(core,'rhwp.js')));await init({module_or_path:await fs.readFile(path.join(core,'rhwp_bg.wasm'))});
  const sharp=require('sharp'),pixels=Buffer.alloc(200*80*4);
  for(let y=0;y<80;y++)for(let x=0;x<200;x++){const at=(y*200+x)*4;pixels[at+(x<100?0:2)]=255;pixels[at+3]=255;}
  const png=await sharp(pixels,{raw:{width:200,height:80,channels:4}}).png().toBuffer();
  const d=C.createEmpty();try{
    d.createBlankDocument();d.pasteHtml(0,0,0,`<p><img src="data:image/png;base64,${png.toString('base64')}" width="200" height="80"></p><table><tr><td style="width:400px"><p>LEAD</p><p></p><p style="text-align:center"><img src="data:image/png;base64,${png.toString('base64')}" width="200" height="80"></p></td></tr></table><p>END</p>`);
    for(const c of JSON.parse(d.getControls()).filter(c=>c.list===0)){
      if(c.ctrlId==='gso'){const props=JSON.parse(d.getPictureProperties(0,c.para,c.controlIndex));d.setPictureProperties(0,c.para,c.controlIndex,JSON.stringify({width:7500,cropRight:props.originalWidth/2}));}
      if(c.ctrlId==='tbl'){d.applyParaFormatInCell(0,c.para,c.controlIndex,0,2,JSON.stringify({alignment:'center'}));const cellPath=JSON.stringify([{controlIndex:c.controlIndex,cellIndex:0,cellParaIndex:2}]),props=JSON.parse(d.getCellPicturePropertiesByPath(0,c.para,cellPath,0));d.setCellPicturePropertiesByPath(0,c.para,cellPath,0,JSON.stringify({width:7500,cropRight:props.originalWidth/2}));}
    }
    const files={};for(const f of ['hwp','hwpx']){files[f]=path.join(directory,'crop.'+f);await fs.writeFile(files[f],f==='hwp'?d.exportHwp():d.exportHwpx())}return files;
  }finally{d.free()}
};
