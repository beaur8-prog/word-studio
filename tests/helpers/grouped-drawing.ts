import JSZip from 'jszip';

export const GROUP_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
export type GroupScenario = 'mixed'|'pictures'|'nested'|'repeated'|'fallback'|'paragraph'|'inline'|'inline-box';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',A='http://schemas.openxmlformats.org/drawingml/2006/main',WPG='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',WPS='http://schemas.microsoft.com/office/word/2010/wordprocessingShape',PIC='http://schemas.openxmlformats.org/drawingml/2006/picture',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const emu=(n:number)=>Math.round(n*9525);
const transform=(x:number,y:number,w:number,h:number)=>`<a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>`;
const picture=(id:number,x:number,y=10,w=80,h=80)=>`<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Group picture ${id}" descr="GROUP PICTURE ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="groupPicture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>${transform(x,y,w,h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`;
const box=`<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>${transform(0,0,240,80)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="205080"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>EDITABLE GROUP STORY</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr wrap="square" lIns="91440" tIns="0" rIns="91440" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp>`;
export async function groupedDrawing(scenario:GroupScenario) {
 const zip=new JSZip();
 zip.file('[Content_Types].xml',`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`);
 zip.file('_rels/.rels',`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
 zip.file('word/_rels/document.xml.rels',`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="groupPicture" Type="${R}/image" Target="media/group-picture.png"/></Relationships>`);
 zip.file('word/styles.xml',`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:color w:val="000000"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`);
 const nested=`<wpg:grpSp><wpg:cNvPr id="20" name="Nested group"/><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="${emu(120)}" y="0"/><a:ext cx="${emu(240)}" cy="${emu(100)}"/><a:chOff x="${emu(10)}" y="${emu(20)}"/><a:chExt cx="${emu(120)}" cy="${emu(50)}"/></a:xfrm></wpg:grpSpPr>${picture(14,20,25,40,40)}</wpg:grpSp>`;
 const children=(scenario==='pictures'?'':box)+(scenario==='inline-box'?'':scenario==='nested'?nested:picture(14,280))+(scenario==='repeated'?picture(15,380):'');
 const group=`<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(480)}" cy="${emu(100)}"/><a:chOff x="0" y="0"/><a:chExt cx="${emu(480)}" cy="${emu(100)}"/></a:xfrm></wpg:grpSpPr>${children}</wpg:wgp>`;
 let drawing=`<w:drawing><wp:anchor distT="0" distB="0" distL="${emu(8)}" distR="${emu(8)}" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>${emu(20)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="${scenario==='paragraph'?'paragraph':'margin'}"><wp:posOffset>${emu(60)}</wp:posOffset></wp:positionV><wp:extent cx="${emu(480)}" cy="${emu(100)}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="10" name="Mixed drawing group"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${WPG}">${group}</a:graphicData></a:graphic></wp:anchor></w:drawing>`;
 if(scenario.startsWith('inline')){const extent=drawing.slice(drawing.indexOf('<wp:extent'),drawing.indexOf('<wp:wrapSquare')),content=drawing.slice(drawing.indexOf('<wp:docPr'),drawing.indexOf('</wp:anchor>'));drawing='<w:drawing><wp:inline>'+extent+content+'</wp:inline></w:drawing>';}
 if(scenario==='fallback')drawing=`<mc:AlternateContent><mc:Choice Requires="wpg">${drawing}</mc:Choice><mc:Fallback><w:pict><v:shape id="legacyGroupPicture" style="width:80px;height:80px"><v:imagedata r:id="groupPicture"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`;
 zip.file('word/document.xml',`<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wpg="${WPG}" xmlns:wps="${WPS}" xmlns:pic="${PIC}" xmlns:r="${R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="wpg wps"><w:body><w:p><w:r><w:t>BODY BEFORE</w:t></w:r></w:p><w:p><w:r>${scenario.startsWith('inline')?'<w:t xml:space="preserve">INLINE BEFORE </w:t>':''}${drawing}${scenario.startsWith('inline')?'<w:t xml:space="preserve"> INLINE AFTER</w:t>':''}</w:r></w:p><w:p><w:r><w:t>BODY AFTER</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
 zip.file('word/media/group-picture.png',Buffer.from(GROUP_PNG,'base64'));
 return zip.generateAsync({type:'nodebuffer'});
}

export async function headerGroupedDrawing(twoPages=false){
 const zip=await JSZip.loadAsync(await groupedDrawing('mixed'));let xml=await zip.file('word/document.xml')!.async('string');const drawing=xml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/)![0];
 const root=xml.slice(0,xml.indexOf('<w:body>')).replace('<w:document','<w:hdr');zip.file('word/header1.xml',root+'<w:p><w:r>'+drawing+'</w:r></w:p></w:hdr>');
 xml=xml.replace(drawing,'').replace('<w:sectPr>','<w:sectPr><w:headerReference w:type="default" r:id="header"/>');if(twoPages)xml=xml.replace('BODY AFTER','SECOND PAGE').replace('<w:t>SECOND PAGE</w:t>','<w:br w:type="page"/><w:t>SECOND PAGE</w:t>');zip.file('word/document.xml',xml);
 zip.file('word/_rels/header1.xml.rels',(await zip.file('word/_rels/document.xml.rels')!.async('string')).replace(/<Relationship\b[^>]*\bId="styles"[^>]*\/>/,''));
 zip.file('word/_rels/document.xml.rels',(await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>','<Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'));
 zip.file('[Content_Types].xml',(await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>','<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'));
 return zip.generateAsync({type:'nodebuffer'});
}

export async function boxGroupWithNestedDrawing(){
 const zip=await JSZip.loadAsync(await groupedDrawing('mixed'));let xml=await zip.file('word/document.xml')!.async('string');
 const drawing=xml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/)![0],shape=xml.match(/<wps:wsp>[\s\S]*?<\/wps:wsp>/)![0];
 const outer=shape.replace('EDITABLE GROUP STORY','OUTER GROUP STORY').replace('</w:txbxContent>',`<w:p><w:r>${drawing.replace('<wp:docPr id="10"','<wp:docPr id="11"')}</w:r></w:p></w:txbxContent>`),second=shape.replace('EDITABLE GROUP STORY','SECOND OUTER STORY').replace('<a:off x="0" y="0"/>','<a:off x="2286000" y="0"/>');
 const start=xml.indexOf(shape),picEnd=xml.indexOf('</pic:pic>',start)+10;xml=xml.slice(0,start)+outer+second+xml.slice(picEnd);zip.file('word/document.xml',xml);return zip.generateAsync({type:'nodebuffer'});
}

/** A separate group member and an inline cell picture use distinct image parts. */
export const TABLE_GROUP_INNER_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVQImWNgYPj/HwADAgH/xCAAOgAAAABJRU5ErkJggg==';
export async function tableGroupedDrawing(scenario:'mixed'|'inline') {
 const zip=await JSZip.loadAsync(await groupedDrawing(scenario));
 const paragraph=(t:string)=>`<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr><w:t>${t}</w:t></w:r></w:p>`;
 const innerPicture=picture(35,0,0,12,12).replace('groupPicture','innerPicture').replaceAll('GROUP PICTURE 35','INNER TABLE PICTURE');
 const image=`<w:drawing><wp:inline><wp:extent cx="${emu(12)}" cy="${emu(12)}"/><wp:docPr id="35" name="Inner table picture" descr="INNER TABLE PICTURE"/><a:graphic><a:graphicData uri="${PIC}">${innerPicture}</a:graphicData></a:graphic></wp:inline></w:drawing>`;
 const margins='<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>';
 const cell=(t:string,extra='')=>`<w:tc><w:tcPr><w:tcW w:w="${extra.includes('gridSpan')?3000:1500}" w:type="dxa"/>${extra}${margins}</w:tcPr>${paragraph(t)}</w:tc>`;
 const table=`<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${cell('MERGED HEADER','<w:gridSpan w:val="2"/>')}</w:tr><w:tr>${cell('ROWSPAN CELL','<w:vMerge w:val="restart"/>')}${cell('UPPER CELL').replace('</w:p>',`<w:r>${image}</w:r></w:p>`)}</w:tr><w:tr>${cell('','<w:vMerge/>')}${cell('TAIL CELL')}</w:tr></w:tbl>${paragraph('BOX TAIL')}`;
 const xml=(await zip.file('word/document.xml')!.async('string')).replaceAll(`cy="${emu(100)}"`,`cy="${emu(160)}"`).replace(`<a:ext cx="${emu(240)}" cy="${emu(80)}"/>`,`<a:ext cx="${emu(240)}" cy="${emu(140)}"/>`);zip.file('word/document.xml',xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/,`<w:txbxContent>${table}</w:txbxContent>`));
 zip.file('word/_rels/document.xml.rels',(await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>',`<Relationship Id="innerPicture" Type="${R}/image" Target="media/inner-picture.png"/></Relationships>`));
 zip.file('word/media/inner-picture.png',Buffer.from(TABLE_GROUP_INNER_PNG,'base64'));return zip.generateAsync({type:'nodebuffer'});
}
