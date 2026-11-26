const fs = require('node:fs/promises'), path = require('node:path'), { pathToFileURL } = require('node:url');
module.exports = async function createFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  await fs.mkdir(directory, { recursive: true }); const files = {};
  for (const policy of [0, 2]) {
    const doc = HwpDocument.createEmpty();
    try {
      doc.createBlankDocument();
      doc.pasteHtml(0, 0, 0, '<p>LEAD</p><table style="width:300px"><tr><td rowspan="2" style="width:100px;height:320px"><p>MERGED ONCE</p></td><td style="width:200px;height:160px"><p>ROW ONE</p></td></tr><tr><td style="width:200px;height:160px"><p>ROW TWO</p></td></tr></table>');
      const c = JSON.parse(doc.getControls()).find(c => c.list === 0 && c.ctrlId === 'tbl');
      doc.setTableProperties(0, c.para, c.controlIndex, JSON.stringify({ pageBreak: policy, repeatHeader: false, treatAsChar: false, textWrap: 'TopAndBottom', vertRelTo: 'Para', vertOffset: 0 }));
      for (let cell = 0; cell < 3; cell++) doc.setCellProperties(0, c.para, c.controlIndex, cell, JSON.stringify({ verticalAlign: 0 }));
      for (const format of ['hwp', 'hwpx']) {
        const bytes = format === 'hwp' ? doc.exportHwp() : doc.exportHwpx(), original = new HwpDocument(bytes);
        try {
          const boxes = JSON.parse(original.getTableCellBboxes(0, c.para, c.controlIndex));
          if (original.pageCount() !== 1 || new Set(boxes.filter(b => b.cellIdx === 0).map(b => b.pageIndex)).size !== 1)
            throw Error('Fixture must start with its merged cell entirely on one source page');
        } finally { original.free(); }
        files[`${format}-${policy}`] = path.join(directory, `merge-${policy}.${format}`);
        await fs.writeFile(files[`${format}-${policy}`], bytes);
      }
    } finally { doc.free(); }
  }
  return files;
};

module.exports.createInsideRowFixture = async function createInsideRowFixture(directory) {
  const core = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(core, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm')) });
  await fs.mkdir(directory, { recursive: true });
  const doc = HwpDocument.createEmpty();
  try {
    doc.createBlankDocument();
    const paper = JSON.parse(doc.getPageDef(0));
    doc.setPageDef(0, JSON.stringify({ ...paper, width: 45000, height: 37500, marginTop: 3000, marginBottom: 3000,
      marginLeft: 3000, marginRight: 3000, marginHeader: 0, marginFooter: 0 }));
    const para = text => `<p style="font-family:Arial;font-size:11pt;line-height:150%;margin:0">${text}</p>`;
    const heights = [32, 34, 35, 42, 148, 41, 34, 36];
    const rows = heights.map((_, index) => {
      const first = index < 2 || index > 6 ? `<td style="width:124px">${para('LEFT ' + index)}</td>` :
        index === 2 ? `<td rowspan="5" style="width:124px">${para('MERGED OWNER')}</td>` : '';
      const content = index === 4 ? Array.from({ length: 6 }, (_, i) => para('TARGET RIGHT ' + (i + 1))).join('') : para('RIGHT ' + index);
      return `<tr>${first}<td style="width:296px">${content}</td></tr>`;
    }).join('');
    const lead = Array.from({ length: 9 }, (_, i) => para('LEAD ' + (i + 1))).join('');
    doc.pasteHtml(0, 0, 0, lead + '<table style="width:420px">' + rows + '</table>' + para('AFTER TABLE'));
    const table = JSON.parse(doc.getControls()).find(control => control.list === 0 && control.ctrlId === 'tbl');
    const address = [0, table.para, table.controlIndex];
    doc.setTableProperties(...address, JSON.stringify({ pageBreak: 2, treatAsChar: false, repeatHeader: false,
      textWrap: 'TopAndBottom', vertRelTo: 'Para', vertOffset: 0 }));
    const { cellCount } = JSON.parse(doc.getTableDimensions(...address));
    for (let i = 0; i < cellCount; i++) {
      const { row, rowSpan } = JSON.parse(doc.getCellInfo(...address, i));
      const height = heights.slice(row, row + rowSpan).reduce((sum, height) => sum + height, 0) * 75;
      doc.setCellProperties(...address, i, JSON.stringify({ height, paddingTop: 75, paddingBottom: 75, applyInnerMargin: true, verticalAlign: 0 }));
    }
    doc.reflowLinesegs();
    const files = {};
    for (const [format, bytes] of [['hwp', doc.exportHwp()], ['hwpx', doc.exportHwpx()]]) {
      const source = new HwpDocument(bytes);
      try {
        const pages = Array.from({ length: source.pageCount() }, (_, index) =>
          JSON.parse(source.getPageTextLayout(index)).runs.flatMap(run => [...run.text.matchAll(/TARGET RIGHT (\d+)/g)].map(m => Number(m[1]))));
        if (JSON.stringify(pages) !== '[[1,2,3,4],[5,6]]') throw Error('Source fixture needs a split inside the merged row');
      } finally { source.free(); }
      const file = path.join(directory, 'inside-merge.' + format);
      await fs.writeFile(file, bytes);
      files[format] = file;
    }
    return files;
  } finally { doc.free(); }
};

const continuationParagraph = text => '<p style="font-family:Arial;font-size:11pt;line-height:150%;margin:0">' + text + '</p>';
module.exports.createLongMergeFixture = async function createLongMergeFixture(directory, { count, rowHeight, sourceEnds }) {
  const heights = [34, 34, rowHeight, 34, 34];
  const core = path.dirname(require.resolve('@rhwp/core'));
  const {default:init, HwpDocument} = await import(pathToFileURL(path.join(core, 'rhwp.js')));
  await init({module_or_path: await fs.readFile(path.join(core, 'rhwp_bg.wasm'))});
  const doc = HwpDocument.createEmpty();
  const files = {};
  await fs.mkdir(directory, { recursive: true });
  try {
    doc.createBlankDocument();
    doc.setPageDef(0, JSON.stringify({...JSON.parse(doc.getPageDef(0)), width:45000, height:37500,
      marginTop:3000, marginBottom:3000, marginLeft:3000, marginRight:3000, marginHeader:0, marginFooter:0}));
    const rows = heights.map((_, i) => {
      const left = i === 1 ? '<td rowspan="2" style="width:124px">' + continuationParagraph('OWNER 1') + continuationParagraph('OWNER 2') + '</td>'
        : i === 2 ? '' : '<td style="width:124px">' + continuationParagraph('LEFT ' + i) + '</td>';
      const right = i === 2 ? Array.from({length:count}, (_, k) => continuationParagraph('TARGET RIGHT ' + (k + 1))).join('')
        : continuationParagraph('RIGHT ' + i);
      return '<tr>' + left + '<td style="width:296px">' + right + '</td></tr>';
    }).join('');
    const lead = Array.from({length:4}, (_, i) => continuationParagraph('LEAD ' + (i + 1))).join('');
    doc.pasteHtml(0, 0, 0, lead + '<table style="width:420px">' + rows + '</table>' + continuationParagraph('AFTER TABLE'));
    const control = JSON.parse(doc.getControls()).find(item => item.list === 0 && item.ctrlId === 'tbl');
    const address = [0, control.para, control.controlIndex];
    doc.setTableProperties(...address, JSON.stringify({pageBreak:2, treatAsChar:false, repeatHeader:false,
      textWrap:'TopAndBottom', vertRelTo:'Para', vertOffset:0}));
    const {cellCount} = JSON.parse(doc.getTableDimensions(...address));
    for (let i = 0; i < cellCount; i++) {
      const {row,rowSpan} = JSON.parse(doc.getCellInfo(...address, i));
      const height = heights.slice(row,row + rowSpan).reduce((total,h) => total + h, 0) * 75;
      doc.setCellProperties(...address, i, JSON.stringify({height, paddingTop:75, paddingBottom:75,
        applyInnerMargin:true, verticalAlign:0}));
    }
    doc.reflowLinesegs();
    for (const [extension, bytes] of [['hwp',doc.exportHwp()], ['hwpx',doc.exportHwpx()]]) {
      const source = new HwpDocument(bytes);
      try {
        const pages = Array.from({length:source.pageCount()}, (_, i) => {
          const numbers = JSON.parse(source.getPageTextLayout(i)).runs.flatMap(item =>
            [...item.text.matchAll(/TARGET RIGHT (\d+)/g)].map(match => +match[1]));
          return numbers.at(-1);
        });
        if (JSON.stringify(pages) !== JSON.stringify(sourceEnds)) throw Error('Source partition changed for ' + extension);
      } finally {source.free();}
      files[extension] = path.join(directory, `long-merge-${count}.${extension}`);
      await fs.writeFile(files[extension], bytes);
    }
    return files;
  } finally {doc.free();}
};
