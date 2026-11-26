'use strict';

const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');

function elements(xml, name) {
  const result = [], stack = [];
  for (const match of xml.matchAll(new RegExp('<\\/?' + name + '\\b[^>]*>', 'g'))) {
    const tag = match[0];
    if (tag.startsWith('</')) {
      const open = stack.pop();
      if (open && !stack.length) {
        const from = open.index + open[0].length;
        result.push({ from, content: xml.slice(from, match.index) });
      }
    } else if (!tag.endsWith('/>')) stack.push(match);
  }
  return result;
}

const controls = (doc, page) => JSON.parse(doc.getPageControlLayout(page)).controls.filter(control => !control.headerFooter);
const isTable = control => control.type === 'table' && !control.cellPath?.length;
const dimensions = control => [control.x,control.y,control.w,control.h];
const placement = (doc, control) => JSON.parse(doc.getTableProperties(control.secIdx,control.paraIdx,control.controlIdx));
const visibleText = doc => {
  const owners = new Map();
  for (let page = 0; page < doc.pageCount(); page++) {
    for (const run of JSON.parse(doc.getPageTextLayout(page)).runs) {
      if (run.headerFooter || !run.text.trim()) continue;
      const key = JSON.stringify([run.secIdx,run.parentParaIdx,run.paraIdx,run.cellPath]);
      owners.set(key, (owners.get(key) || '') + run.text);
    }
  }
  return JSON.stringify([...owners].sort());
};

module.exports = (Core, original, saved) => {
  const source = controls(original,0).find(c => isTable(c) && c.paraIdx === 8);
  const target = controls(saved,1).find(c => isTable(c) && c.paraIdx === 10);
  if (!source || !target) return { missingWitness: true };
  const sourceInsets = placement(original,source), generatedInsets = placement(saved,target);
  const archive = unzipSync(saved.exportHwpx()), file = 'Contents/section' + target.secIdx + '.xml';
  const xml = strFromU8(archive[file]), owner = elements(xml,'hp:p')[target.paraIdx], table = elements(owner.content,'hp:tbl')[0];
  const margin = /<hp:outMargin\b[^>]*\/>/.exec(table.content);
  if (!margin) return { missingMargin: true };
  const at = owner.from + table.from + margin.index, prior = visibleText(saved), trials = [];
  const pack = fields => fields.reduce((tag,key) => tag.replace(new RegExp('\\b' + key + '="[^"]*"'), key + '="' +
    sourceInsets['outer' + key[0].toUpperCase() + key.slice(1)] + '"'), margin[0]);
  for (const [name, fields, erasePageBreak] of [
    ['bottom',['bottom'],false],['topAndBottom',['top','bottom'],false],['allOuter',['top','bottom','left','right'],false],
    ['noGeneratedBreak',[],true],['breakAndOuter',['top','bottom','left','right'],true],
  ]) {
    const tag = pack(fields);
    const doc = new Core(zipSync({ ...archive, [file]:strToU8(xml.slice(0,at)+tag+xml.slice(at+margin[0].length)) },{level:1}));
    try {
      const changed = erasePageBreak && JSON.parse(doc.applyParaFormat(target.secIdx,target.paraIdx,JSON.stringify({pageBreakBefore:false}))).ok;
      const onFirst = controls(doc,0).find(c => isTable(c) && c.paraIdx === target.paraIdx);
      trials.push({name,changed,pages:doc.pageCount(),samePaintedText:visibleText(doc)===prior,onFirst:onFirst&&dimensions(onFirst)});
    } finally { doc.free(); }
  }
  const sideKeys=['outerTop','outerBottom','outerLeft','outerRight'];
  const originalHost = JSON.parse(original.getParaPropertiesAt(source.secIdx,source.paraIdx));
  const generatedHost = JSON.parse(saved.getParaPropertiesAt(target.secIdx,target.paraIdx));
  const before = Array.from({length:target.paraIdx + 1},(_,p)=>{
    const properties=JSON.parse(saved.getParaPropertiesAt(target.secIdx,p));
    return [p,saved.getParagraphLength(target.secIdx,p),properties.pageBreakBefore,properties.lineSpacing,
      JSON.parse(saved.getCharPropertiesAt(target.secIdx,p,0)).fontSize];
  });
  return { reference:dimensions(source), generated:dimensions(target), sourceInsets:sideKeys.map(k=>sourceInsets[k]),
    generatedInsets:sideKeys.map(k=>generatedInsets[k]), sourceBreak:originalHost.pageBreakBefore, generatedBreak:generatedHost.pageBreakBefore,
    before,trials };
};
