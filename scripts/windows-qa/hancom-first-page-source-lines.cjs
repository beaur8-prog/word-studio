'use strict';

const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
const pageAgreement = require('./hancom-page-agreement.cjs');
const unit = /<hp:switch>[\s\S]*?<hh:margin>[\s\S]*?<\/hp:switch>|<hh:margin>[\s\S]*?<\/hh:margin><hh:lineSpacing\b[^>]*\/>/;
const cache = /<hp:linesegarray\b[^>]*>[\s\S]*?<\/hp:linesegarray>/;

function ranges(source,name) {
  const result = [], stack = [];
  for (const { 0:tag, index } of source.matchAll(new RegExp('<\\/?'+name+'\\b[^>]*>', 'g'))) {
    if (tag.startsWith('</')) {
      const open = stack.pop();
      if (open && !stack.length) result.push({ ...open, end:index+tag.length, inner:source.slice(open.start+open.tag.length,index) });
    } else if (!tag.endsWith('/>')) stack.push({ tag, start:index });
  }
  return result;
}
function id(tag, attribute='id') { return tag?.match(new RegExp('\\b'+attribute+'="(\\d+)"'))?.[1]; }
function indexText(doc) {
  const result = new Map();
  for (let p=0; p<doc.getParagraphCount(0); p++) {
    const text = doc.getTextRange(0,p,0,0xffffffff);
    if (text.trim()) result.set(text, result.has(text) ? null : p);
  }
  return result;
}

function read(doc) {
  const parts = unzipSync(doc.exportHwpx());
  const xml = strFromU8(parts['Contents/section0.xml']);
  const head = strFromU8(parts['Contents/header.xml']);
  const paras = ranges(xml,'hp:p');
  const styles = new Map(ranges(head,'hh:paraPr').map(r=>[id(r.tag),r]));
  const first = JSON.parse(doc.getPageInfo(0));
  const wide = JSON.parse(doc.getPageControlLayout(0)).controls.filter(c =>
    c.secIdx===0 && c.type==='table' && !c.headerFooter && !c.cellPath?.length &&
    c.w >= (first.width-first.marginLeft-first.marginRight) * .95);
  if (wide.length !== 1) throw Error('Need one complete paper-width table');
  const table = wide[0], owner = paras[table.paraIdx];
  const tables = ranges(owner.inner,'hp:tbl').filter(t =>
    Math.abs(Number(t.inner.match(/^<hp:sz\b[^>]*\bwidth="(\d+)"/)?.[1])/75-table.w)<.5);
  if (tables.length !== 1) throw Error('Cannot uniquely map table structure');
  const [node] = tables;
  const shape = node.inner.match(/^<hp:sz\b[^>]*\/>(<hp:pos\b[^>]*\/>)(<hp:outMargin\b[^>]*\/>)/);
  if (!shape) throw Error('Missing table placement');
  const placement = shape[1]+shape[2];
  const start = owner.start+owner.tag.length+node.start+node.tag.length+shape[0].length-placement.length;
  return {parts,xml,head,paras,styles,first,table,start,placement};
}

function paint(doc) {
  const result = new Map();
  for (let page=0; page<Math.min(doc.pageCount(),4); page++) {
    const {columns} = JSON.parse(doc.getPageInfo(page));
    for (const run of JSON.parse(doc.getPageTextLayout(page)).runs) {
      if (run.secIdx || run.headerFooter || run.cellPath?.length || !run.text.trim()) continue;
      const lane = columns.findIndex(c=>run.x>=c.x-1&&run.x<c.x+c.width+1);
      const lines = result.get(run.paraIdx) ?? new Map();
      const key = [page,lane,run.y].join(':');
      lines.set(key,{page,lane,text:(lines.get(key)?.text??'')+run.text.replace(/\s/gu,'')});
      result.set(run.paraIdx,lines);
    }
  }
  return result;
}

module.exports = function trial(Core, source, output) {
  const original = read(source), saved = read(output);
  const titleAlreadyAuthored = saved.placement.includes('treatAsChar="0"') &&
    Math.abs(saved.table.x-original.table.x)<.16 && Math.abs(saved.table.y-original.table.y)<.16;
  if (original.first.columns.length !== 2 || saved.first.columns.length !== 2 ||
      Math.abs(original.table.w-saved.table.w)>.5 || Math.abs(original.table.h-saved.table.h)>1 ||
      !original.placement.includes('treatAsChar="0"') ||
      !titleAlreadyAuthored && !saved.placement.includes('treatAsChar="1"'))
    throw Error('Unexpected title carrier geometry');
  const from = indexText(source), to = indexText(output);
  const originalPaint = paint(source);
  const proof = [];
  for (const [text,index] of from) {
    const next = to.get(text), originalOwner = original.paras[index], savedOwner = saved.paras[next];
    if (index===null || !Number.isInteger(next) || !originalOwner || !savedOwner ||
        originalPaint.get(index)?.values().next().value?.page !== 0 || text.length<8 || text.length>200 ||
        /[^\p{L}\p{N}\p{P}\p{S}\p{Zs}\t]/u.test(text) || /<hp:(?:ctrl|pic|tbl|shape)\b/.test(originalOwner.inner)) continue;
    const before = original.styles.get(id(originalOwner.tag,'paraPrIDRef'));
    const after = saved.styles.get(id(savedOwner.tag,'paraPrIDRef'));
    const line = originalOwner.inner.match(cache)?.[0];
    const authoredUnit = before?.inner.match(unit)?.[0], generatedUnit = after?.inner.match(unit)?.[0];
    if (line && authoredUnit && generatedUnit && JSON.parse(source.getParagraphLineMetrics(0,index,'[]')).controlCount===0)
      proof.push({index,next,owner:savedOwner,after,line,authoredUnit,generatedUnit,tab:text.includes('\t')});
  }
  if (!proof.some(row=>row.tab)) throw Error('Missing an unchanged authored tab row');
  const count = (doc, title) => {
    const painted = paint(doc);
    let starts=0, exactLineBreaks=0;
    for (const row of proof) {
      const a = [...(originalPaint.get(row.index)?.values()??[])];
      const b = [...(painted.get(row.next)?.values()??[])];
      starts += Number(a[0]?.page===b[0]?.page && a[0]?.lane===b[0]?.lane);
      exactLineBreaks += Number(a.length===b.length && a.every((line,i)=>line.text===b[i].text));
    }
    const unchanged = [...to].every(([text,index])=>index===null||doc.getTextRange(0,index,0,0xffffffff)===text);
    if (!unchanged) throw Error('Changed paragraph payload');
    const {body,cell} = pageAgreement(source,doc);
    return { title,pages:doc.pageCount(),starts,exactLineBreaks,body,cell };
  };
  const results = [count(output, 'ordinary output')];
  for (const scope of ['title only','tab rows','all exact first-page rows']) {
    if (scope==='title only' && titleAlreadyAuthored) continue;
    const edits = titleAlreadyAuthored ? [] : [[saved.start,saved.start+saved.placement.length,original.placement]];
    const additions = [];
    let nextStyle = 1+Math.max(...[...saved.styles.keys()].map(Number));
    for (const row of proof) {
      if (scope==='title only'||scope==='tab rows'&&!row.tab) continue;
      const style = nextStyle++;
      additions.push(row.after.tag.replace(/\bid="\d+"/,'id="'+style+'"')+
        row.after.inner.replace(row.generatedUnit,row.authoredUnit)+'</hh:paraPr>');
      edits.push([row.owner.start,row.owner.end,row.owner.tag.replace(/\bparaPrIDRef="\d+"/,'paraPrIDRef="'+style+'"')+
        row.owner.inner.replace(cache,'')+row.line+'</hp:p>']);
    }
    let xml=saved.xml, head=saved.head;
    for (const [start,end,edited] of edits.sort((a,b)=>b[0]-a[0])) xml=xml.slice(0,start)+edited+xml.slice(end);
    if (additions.length) head=head.replace(/(<hh:paraProperties\b[^>]*itemCnt=")(\d+)(")/,
      (_,begin,old,end)=>begin+(Number(old)+additions.length)+end).replace('</hh:paraProperties>',additions.join('')+'</hh:paraProperties>');
    const changed = new Core(zipSync({...saved.parts,'Contents/header.xml':strToU8(head),
      'Contents/section0.xml':strToU8(xml)},{level:1}));
    try { results.push({...count(changed,scope), ownerEdits:additions.length}); }
    finally { changed.free(); }
  }
  return { originalPages:source.pageCount(),fixed:proof.length,titleAlreadyAuthored,results };
};
