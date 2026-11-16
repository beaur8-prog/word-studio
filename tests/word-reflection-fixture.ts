import JSZip from 'jszip';

export const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const word14Namespace = 'http://schemas.microsoft.com/office/word/2010/wordml';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRelationship = 'http://schemas.openxmlformats.org/package/2006/relationships';
const compatibility = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const unrelated = 'urn:word-studio:qa:unrelated-reflection';

const sourceAttributes = {
  blurRad: '7620', stA: '45000', stPos: '0', endA: '300', endPos: '36000',
  dist: '0', dir: '5400000', fadeDir: '5400000', sx: '100000', sy: '-90000',
  kx: '0', ky: '0', algn: 'bl',
};
// All positive attributes are scalar. Deliberately vary the property values
// of the run and paragraph mark to catch accidental sharing between them.
export const reflectionAttributes = {
  directRun: sourceAttributes,
  paragraphMark: {
    ...sourceAttributes, blurRad: '13970', stA: '31000', endA: '0', endPos: '45000',
    dist: '1105', sy: '-100000',
  },
  effectOnlyRun: { ...sourceAttributes, blurRad: '3175', stA: '45000', dist: '1270' },
  emptyTitle: { ...sourceAttributes, blurRad: '13970', endPos: '32000', dist: '2540' },
} as const;

export const reflectionLabels = {
  direct: 'REFLECTED-RUN',
  plainSibling: ' PLAIN-SIBLING',
  effectOnly: 'EFFECT-ONLY',
  unrelated: 'CONTROL-PARAGRAPH',
  unsafe: 'UNSUPPORTED-EFFECT-TEXT',
} as const;

export type ReflectionFixture = 'supported' | 'wrong-namespace' | 'unknown-child' | 'numeric-excess';

function reflection(attributes: Readonly<Record<string, string>>, prefix = 'w14', children = '') {
  const values = Object.entries(attributes).map(([name, value]) => `w14:${name}="${value}"`).join(' ');
  return children
    ? `<${prefix}:reflection ${values}>${children}</${prefix}:reflection>`
    : `<${prefix}:reflection ${values}/>`;
}

function run(text: string, properties = '') {
  return `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}
const paragraph = (runs = '', properties = '') => `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}${runs}</w:p>`;

function supportedBody() {
  const { direct, plainSibling, effectOnly, unrelated } = reflectionLabels;
  return [
    paragraph(
      run(direct, '<w:b/><w:sz w:val="24"/>' + reflection(reflectionAttributes.directRun)) + run(plainSibling),
      '<w:rPr><w:i/>' + reflection(reflectionAttributes.paragraphMark) + '</w:rPr>',
    ),
    paragraph(run(effectOnly, reflection(reflectionAttributes.effectOnlyRun))),
    paragraph('', '<w:pStyle w:val="Title"/><w:rPr>' + reflection(reflectionAttributes.emptyTitle) + '</w:rPr>'),
    paragraph(run(unrelated)),
  ].join('');
}

function unsupportedBody(scenario: Exclude<ReflectionFixture, 'supported'>) {
  const effect =
    scenario === 'wrong-namespace' ? reflection(sourceAttributes, 'other') :
    scenario === 'unknown-child' ? reflection(sourceAttributes, 'w14', '<w14:qaUnsupportedChild/>') :
    reflection({ ...sourceAttributes, blurRad: '999999999999999999999' });
  return paragraph(run(reflectionLabels.unsafe, effect)) + paragraph(run(reflectionLabels.unrelated));
}

export async function reflectionFixture(scenario: ReflectionFixture = 'supported') {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="${packageRelationship}">
    <Relationship Id="rId1" Type="${relationship}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="${packageRelationship}">
    <Relationship Id="rStyles" Type="${relationship}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('word/styles.xml', `<w:styles xmlns:w="${wordNamespace}">
    <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>
      <w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>
  </w:styles>`);
  const body = scenario === 'supported' ? supportedBody() : unsupportedBody(scenario);
  zip.file('word/document.xml', `<w:document xmlns:w="${wordNamespace}" xmlns:w14="${word14Namespace}"
    xmlns:mc="${compatibility}" xmlns:other="${unrelated}" mc:Ignorable="w14 other"><w:body>${body}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}
