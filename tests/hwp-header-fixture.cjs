const fs=require('node:fs/promises'),path=require('node:path'),{pathToFileURL}=require('node:url'),JSZip=require('jszip');
module.exports=async function(directory){
  const core=path.dirname(require.resolve('@rhwp/core'));
  const {default:init,HwpDocument}=await import(pathToFileURL(path.join(core,'rhwp.js')));
  await init({module_or_path:await fs.readFile(path.join(core,'rhwp_bg.wasm'))});
  const d=HwpDocument.createEmpty();
  try{
    d.createBlankDocument(); d.pasteHtml(0,0,0,'<p>PAGE ONE</p><p>PAGE TWO</p><p>PAGE THREE</p>');
    for(let i=0;i<d.getParagraphCount(0);i++) if(/TWO|THREE/.test(d.getTextRange(0,i,0,100))) d.applyParaFormat(0,i,JSON.stringify({pageBreakBefore:true}));
    const shape=JSON.parse(d.createShapeControl(JSON.stringify({sectionIdx:0,paraIdx:0,charOffset:0,shapeType:'textbox',width:18000,height:2400,treatAsChar:false,textWrap:'InFrontOfText'})));
    d.insertTextInCell(0,shape.paraIdx,shape.controlIdx,0,0,0,'HEADER TITLE ');
    d.setShapeProperties(0,shape.paraIdx,shape.controlIdx,JSON.stringify({horzRelTo:'Paper',vertRelTo:'Paper',horzOffset:6000,vertOffset:1500,borderWidth:75,borderColor:0,fillType:'none'}));
    const zip=await JSZip.loadAsync(d.exportHwpx()); let xml=await zip.file('Contents/section0.xml').async('string');
    const rectangle=xml.match(/<hp:rect\b[\s\S]*?<\/hp:rect>/)[0]; xml=xml.replace(rectangle,'');
    const header=(parity,id)=>`<hp:ctrl><hp:header id="${id}" applyPageType="${parity}"><hp:subList id="0" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0"><hp:p id="${id+1}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${rectangle.replace(/\bid="\d+"/,`id="${id+2}"`).replace('<hp:t>HEADER TITLE </hp:t>',`<hp:t>${parity} HEADER </hp:t><hp:ctrl><hp:autoNum num="37" numType="PAGE"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supScript="0"/></hp:autoNum></hp:ctrl>`)}</hp:run></hp:p></hp:subList></hp:header></hp:ctrl>`;
    xml=xml.replace('<hp:t>PAGE ONE</hp:t>',header('EVEN',200000)+header('ODD',200100)+'<hp:ctrl><hp:pageHiding hideHeader="1" hideFooter="0" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0"/></hp:ctrl><hp:t>PAGE ONE</hp:t>');
    zip.file('Contents/section0.xml',xml);const source=new HwpDocument(await zip.generateAsync({type:'uint8array'}));
    try{const files={};for(const f of ['hwp','hwpx']){files[f]=path.join(directory,'headers.'+f);await fs.writeFile(files[f],f==='hwp'?source.exportHwp():source.exportHwpx())}return files;}finally{source.free()}
  }finally{d.free()}
};
