import JSZip from 'jszip';

const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationship = 'http://schemas.openxmlformats.org/package/2006/relationships';
const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const wsh = 'urn:word-studio:hwp-cell-lines:1';
const compact = '1 Percent 13.3333 21.3333 13.3 13.3';
const terminal = '1 Percent 13.3333 21.3 13.3 21.3';
const widths = [1095, 1275, 2070, 5175];
export const headLabels = ['HEAD-A', 'HEAD-B', 'HEAD-C', 'HEAD-D'];
export const floatLabel = 'QA FLOAT FENCE';
export const localLabel = 'FLOAT-LOCAL-PROBE';
export const remoteLabel = 'FLOAT-REMOTE-PROBE';
export const rowLabels = ['ROW-01', 'ROW-02', 'ROW-03', 'ROW-04 WIDE-PART-A', 'ROW-05', 'ROW-06'];
// The same self-authored solid red PNG used by the separate drawing cycle spec.
const red = 'iVBORw0KGgoAAAANSUhEUgAAAEQAAABWCAIAAAD5b6d1AAAAfUlEQVR4nO3PAQkAIADAMCPZwe7mMoZXBg/wjT3XN43rBzAwbwVTDaYaTDWYajDVYKrBVIOpBlMNphpMNZhqMNVgqsFUg6kGUw2mGkw1mGow1WCqwVSDqQZTDaYaTDWYajDVYKrBVIOpBlMNphpMNZhqMNVgqsFUg6n2FeYAv7/vrqlGY2MAAAAASUVORK5CYII=';
type Scenario = 'absent' | 'remote' | 'local' | 'outside-body';

function floatingDrawing(top = 940) {
  const [x, y, width, height] = [88, top, 68, 72].map(n => n * 9525);
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${height}"
    behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>
    <wp:extent cx="${width}" cy="${height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapSquare wrapText="bothSides" distT="0" distB="0" distL="0" distR="0"/>
    <wp:docPr id="66" name="${floatLabel}" descr="${floatLabel}"/>
    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
    <a:graphic><a:graphicData uri="${pic}"><pic:pic>
      <pic:nvPicPr><pic:cNvPr id="66" name="figure-fence.png"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="rId66"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic>
  </wp:anchor></w:drawing></w:r>`;
}

export async function fenceFixture(scenario: Scenario) {
  const textRun = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const paragraph = (text: string, line = 200, policy = '', suffix = '') => {
    const value = JSON.stringify({ policy, font: 13.3333, line }).replaceAll('"', '&quot;');
    const attributes = policy ? ` xmlns:mc="${mc}" xmlns:wsh="${wsh}" mc:Ignorable="wsh" mc:PreserveAttributes="wsh:cellLine" wsh:cellLine="${value}"` : '';
    const run = text ? `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r>` : '';
    return `<w:p${attributes}><w:pPr><w:keepNext w:val="off"/><w:keepLines w:val="off"/><w:widowControl w:val="off"/>
      <w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="exact"/><w:jc w:val="left"/></w:pPr>${run}${suffix}</w:p>`;
  };
  const marker = (text: string, suffix = '') => `<w:p><w:pPr><w:spacing w:before="0" w:after="0"
    w:line="200" w:lineRule="exact"/></w:pPr>${textRun(text)}${suffix}</w:p>`;
  const cell = (column: number, content: string, merge = '') => `<w:tc><w:tcPr>
    <w:tcW w:w="${widths[column]}" w:type="dxa"/>${merge ? `<w:vMerge w:val="${merge}"/>` : ''}<w:vAlign w:val="center"/>
    </w:tcPr>${content}</w:tc>`;
  const row = (cells: string[], header = false) => `<w:tr><w:trPr><w:cantSplit/>${header ? '<w:tblHeader/>' : ''}
    <w:trHeight w:val="${header ? 314 : 372}" w:hRule="atLeast"/></w:trPr>${cells.join('')}</w:tr>`;
  const header = row(headLabels.map((text, i) => cell(i, paragraph(text, 200, compact))), true);
  const body = rowLabels.map((label, i) => row([
    cell(0, i === 0 ? paragraph('TOP-A', 320) + paragraph('BOT-A', 200, compact) :
      paragraph(i === 5 ? 'NOTE-A' : ''), i === 0 ? 'restart' : i < 5 ? 'continue' : ''),
    cell(1, i === 0 ? paragraph('OWN-B MARK-B', 320, terminal) : paragraph(i >= 4 ? 'NOTE-B' : ''),
      i === 0 ? 'restart' : i < 4 ? 'continue' : ''),
    cell(2, i === 3 ? paragraph(label, 320, terminal) : paragraph(label)),
    cell(3, paragraph(`CHECK-0${i + 1}`)),
  ])).join('');
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map(side => `<w:${side} w:val="single" w:sz="3" w:color="000000"/>`).join('');
  const padding = ['top', 'left', 'bottom', 'right'].map(side =>
    `<w:${side} w:w="${side === 'top' || side === 'bottom' ? 57 : 113}" w:type="dxa"/>`).join('');
  const table = `<w:tbl><w:tblPr><w:tblW w:w="9615" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders>
    <w:tblLayout w:type="fixed"/><w:tblCellMar>${padding}</w:tblCellMar></w:tblPr>
    <w:tblGrid>${widths.map(n => `<w:gridCol w:w="${n}"/>`).join('')}</w:tblGrid>${header}${body}</w:tbl>`;
  const prefix = [1, 2, 3, 4].map(i => paragraph(`PREFIX-0${i}`, i === 4 ? 2820 : 3150, '',
    i === 1 ? textRun(localLabel) + (scenario === 'local' ? floatingDrawing() : '') : '')).join('');
  const authoredBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const late = marker(remoteLabel, scenario === 'remote' ? floatingDrawing() : scenario === 'outside-body' ? floatingDrawing(1030) : '') + marker('AFTER-FLOAT');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    ${scenario === 'absent' ? '' : '<Default Extension="png" ContentType="image/png"/>'}
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="${packageRelationship}">
    <Relationship Id="rId1" Type="${relationship}/officeDocument" Target="word/document.xml"/></Relationships>`);
  if (scenario !== 'absent') {
    zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="${packageRelationship}">
      <Relationship Id="rId66" Type="${relationship}/image" Target="media/fence.png"/></Relationships>`);
    zip.file('word/media/fence.png', red, { base64: true });
  }
  zip.file('word/document.xml', `<w:document xmlns:w="${word}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:pic="${pic}" xmlns:r="${relationship}"><w:body>
    ${prefix}${table}${authoredBreak}${late}${paragraph('TAIL')}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1134" w:bottom="1134"
      w:left="1134" w:header="850" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}
