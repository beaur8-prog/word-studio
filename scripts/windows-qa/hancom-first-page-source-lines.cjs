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

// Preserve the authored character offsets while matching the quarter-em
// representation used by the current DOCX/HTML conversion path.
function sourceFixedSpaceKey(owner, literal) {
  if (!owner?.inner.includes('<hp:fwSpace')) return literal;
  const controls = { fwSpace:'\u2007', nbSpace:'\u00a0', tab:'\t', lineBreak:'\n' };
  const entities = { amp:'&', lt:'<', gt:'>', quot:'"', apos:"'" };
  let reconstructed = '', projected = '';
  for (const run of owner.inner.matchAll(/<hp:t(?:\s[^>]*)?>([\s\S]*?)<\/hp:t>/g)) {
    for (const token of run[1].matchAll(/<[^>]+>|[^<]+/g)) {
      const tag = /^<hp:(fwSpace|nbSpace|tab|lineBreak)\b[^>]*\/>$/.exec(token[0]);
      if (!tag && token[0].startsWith('<')) return literal;
      if (tag) {
        const name = tag[1];
        reconstructed += controls[name];
        projected += name === 'fwSpace' ? '\u2005' : controls[name];
      } else {
        const value = token[0].replace(/&(#\d+|#x[\da-f]+|amp|lt|gt|quot|apos);/gi, (original, name) => {
          if (entities[name]) return entities[name];
          const codePoint = name.startsWith('#x') ? parseInt(name.slice(2),16) : parseInt(name.slice(1),10);
          return codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : original;
        });
        reconstructed += value;
        projected += value;
      }
    }
  }
  return reconstructed === literal && projected.length === literal.length ? projected : literal;
}

function indexText(doc, structure) {
  const result = new Map();
  for (let p=0; p<doc.getParagraphCount(0); p++) {
    const literal = doc.getTextRange(0,p,0,0xffffffff);
    const text = structure ? sourceFixedSpaceKey(structure.paras[p],literal) : literal;
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

function paint(doc, limit = 4) {
  const result = new Map();
  for (let page=0; page<Math.min(doc.pageCount(),limit); page++) {
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
  const expand = process.env.WINDOWS_HWP_FIRST_PAGE_EXPAND === '1';
  const includeLong = expand && process.env.WINDOWS_HWP_FIRST_PAGE_INCLUDE_LONG === '1';
  const matchFixedSpaces = includeLong && process.env.WINDOWS_HWP_FIRST_PAGE_PROJECT_FWSPACE === '1';
  const traceLanes = matchFixedSpaces && process.env.WINDOWS_HWP_FIRST_PAGE_TRACE_LANES === '1';
  const from = indexText(source, matchFixedSpaces ? original : undefined), to = indexText(output);
  const limit = expand ? 40 : 4;
  const originalPaint = paint(source,limit);
  const proof = [];
  for (const [text,index] of from) {
    const next = to.get(text), originalOwner = original.paras[index], savedOwner = saved.paras[next];
    const start = originalPaint.get(index)?.values().next().value;
    const page = start?.page;
    if (index===null || !Number.isInteger(next) || !originalOwner || !savedOwner ||
        page===undefined || !expand && page!==0 || text.length<8 || !includeLong && text.length>200 ||
        /[^\p{L}\p{N}\p{P}\p{S}\p{Zs}\t]/u.test(text) || /<hp:(?:ctrl|pic|tbl|shape)\b/.test(originalOwner.inner)) continue;
    const before = original.styles.get(id(originalOwner.tag,'paraPrIDRef'));
    const after = saved.styles.get(id(savedOwner.tag,'paraPrIDRef'));
    const line = originalOwner.inner.match(cache)?.[0];
    const authoredUnit = before?.inner.match(unit)?.[0], generatedUnit = after?.inner.match(unit)?.[0];
    if (line && authoredUnit && generatedUnit && JSON.parse(source.getParagraphLineMetrics(0,index,'[]')).controlCount===0)
      proof.push({index,next,page,lane:start.lane,length:text.length,projected:matchFixedSpaces &&
        source.getTextRange(0,index,0,0xffffffff)!==text,owner:savedOwner,after,line,authoredUnit,generatedUnit,tab:text.includes('\t')});
  }
  if (!proof.some(row=>row.tab)) throw Error('Missing an unchanged authored tab row');
  const count = (doc, title) => {
    const painted = paint(doc,limit);
    let starts=0, exactLineBreaks=0, firstPageStarts=0, firstPageExactLineBreaks=0;
    const pageTrace = {};
    for (const row of proof) {
      const a = [...(originalPaint.get(row.index)?.values()??[])];
      const b = [...(painted.get(row.next)?.values()??[])];
      const sameStart = a[0]?.page===b[0]?.page && a[0]?.lane===b[0]?.lane;
      const sameLines = a.length===b.length && a.every((line,i)=>line.text===b[i].text);
      starts += Number(sameStart);
      exactLineBreaks += Number(sameLines);
      if (row.page===0) { firstPageStarts += Number(sameStart); firstPageExactLineBreaks += Number(sameLines); }
      if (traceLanes) {
        const report = pageTrace[row.page+1] ??= { fixed:0, sameStart:0, sameLines:0, emittedPages:{} };
        const destination = b[0] ? b[0].page+1 : 'unmapped';
        report.fixed++;
        report.sameStart += Number(sameStart);
        report.sameLines += Number(sameLines);
        report.emittedPages[destination] = (report.emittedPages[destination] ?? 0) + 1;
        if (!sameStart && !report.earliestShift) report.earliestShift = {
          sourceOwner: row.index, savedOwner: row.next, sourceLane: a[0]?.lane,
          savedLane: b[0]?.lane, startPage: destination, literalLength: row.length,
          lineCounts: [a.length,b.length], correspondsByFixedSpace: row.projected,
        };
      }
    }
    const unchanged = [...to].every(([text,index])=>index===null||doc.getTextRange(0,index,0,0xffffffff)===text);
    if (!unchanged) throw Error('Changed paragraph payload');
    const {body,cell} = pageAgreement(source,doc);
    return { title,pages:doc.pageCount(),starts,exactLineBreaks,firstPageStarts,firstPageExactLineBreaks,body,cell,...(traceLanes ? { pageTrace } : {}) };
  };
  const results = [count(output, 'ordinary output')];

  function shiftedRightBoundary(doc) {
    const painted = paint(doc,limit);
    const row = proof.find(owner => {
      const a = [...(originalPaint.get(owner.index)?.values() ?? [])];
      const b = [...(painted.get(owner.next)?.values() ?? [])];
      return owner.page >= 2 && a[0]?.lane === 1 && b[0]?.lane === 0 && b[0]?.page === owner.page+1 &&
        a.length === b.length && a.every((line,i) => line.text === b[i].text);
    });
    if (!row) return null;
    const savedStructure = read(doc);
    const surrounding = (document,structure,center) => {
      const owners = [];
      for (let p=Math.max(0,center-2); p<=Math.min(center+2,document.getParagraphCount(0)-1); p++) {
        const location = (() => {
          try { const value = JSON.parse(document.getPageOfPosition(0,p)); return value.ok ? value.page+1 : null; }
          catch (error) { if (typeof error === 'string' && /^렌더링 오류: 문단 .*이 페이지에 없습니다$/.test(error)) return null; throw error; }
        })();
        const options = JSON.parse(document.getParaPropertiesAt(0,p));
        const { stored, controlCount, lines } = JSON.parse(document.getParagraphLineMetrics(0,p,'[]'));
        owners.push({ p, location, controls:controlCount, stored, lineCount:lines.length,
          linePositions:lines.slice(0,3).map(line=>[line.verticalPosition,line.height,line.spacing]),
          units:document.getParagraphLength(0,p), pageFormatBreak:options.pageBreakBefore,
          rawBreak:id(structure.paras[p]?.tag,'pageBreak'), rawColumnBreak:id(structure.paras[p]?.tag,'columnBreak'),
          spacing:[options.spacingBefore,options.spacingAfter,options.lineSpacing],
        });
      }
      return owners;
    };
    const pageDetails = (document,index) => {
      const paper = JSON.parse(document.getPageInfo(index));
      const controls = JSON.parse(document.getPageControlLayout(index)).controls
        .filter(control=>!control.headerFooter).map(control=>[control.type,control.paraIdx,control.x,control.y,control.w,control.h]);
      const lines = JSON.parse(document.getPageTextLayout(index)).runs
        .filter(run=>run.secIdx===0&&!run.cellPath?.length&&!run.headerFooter&&run.text.trim())
        .reduce((result,run)=>{
          const lane = paper.columns.findIndex(column=>run.x>=column.x-1&&run.x<column.x+column.width+1);
          const band = result[lane] ??= { top:run.y,bottom:run.y+run.h,owners:new Set() };
          band.top=Math.min(band.top,run.y); band.bottom=Math.max(band.bottom,run.y+run.h); band.owners.add(run.paraIdx);
          return result;
        },{});
      return {page:index+1,columns:paper.columns,body:[paper.marginTop+paper.marginHeader,paper.footerArea.y],
        controls,lanes:Object.fromEntries(Object.entries(lines).map(([lane,{top,bottom,owners}])=>[lane,[top,bottom,owners.size]]))};
    };
    const prior = surrounding(source,original,row.index), after = surrounding(doc,savedStructure,row.next);
    const authored = prior.find(owner=>owner.p===row.index), converted = after.find(owner=>owner.p===row.next);
    const measured = {
      sourceOwner:row.index,savedOwner:row.next, source:prior, saved:after,
      physical:[pageDetails(source,row.page),pageDetails(doc,row.page),pageDetails(doc,row.page+1)], trials:[],
    };
    if (authored && converted && !authored.pageFormatBreak && converted.pageFormatBreak) {
      const changed = new Core(doc.exportHwpx());
      try {
        const control = count(changed,'reopen diagnostic control');
        const ok = JSON.parse(changed.applyParaFormat(0,row.next,JSON.stringify({pageBreakBefore:false}))).ok;
        if (ok) {
          const result = count(changed,'clear only this generated break');
          measured.trials.push({control:[control.pages,control.body.sameStartPage,control.body.sameLineBreaks],
            result:[result.pages,result.body.sameStartPage,result.body.sameLineBreaks], first:result.firstPageStarts});
        }
      } finally { changed.free(); }
    }
    return measured;
  }

  const literal = row => !row.projected;
  const onFirst = row => literal(row) && row.page===0;
  const scopes = [
    ['title only',()=>false], ['first-page tab rows',row=>onFirst(row)&&row.tab], ['all exact first-page rows',onFirst],
    ...(expand ? [['first two source pages',row=>literal(row) && row.page<=1], ['first three source pages',row=>literal(row) && row.page<=2],
      ['all eligible source pages',row=>literal(row) && (!includeLong || row.length<=200)],
      ['after the first page only',row=>literal(row) && !onFirst(row)]] : []),
    ...(includeLong ? [['all exact including long owners',literal],['long exact owners only',row=>literal(row) && row.length>200]] : []),
    ...(matchFixedSpaces ? [['all including exact projected fwSpace',()=>true],
      ['new matches after source page two',row=>literal(row) || row.page>=2],
      ['new matches after source page three',row=>literal(row) || row.page>=3],
      ['new matches after source page four',row=>literal(row) || row.page>=4]] : []),
    ...(traceLanes ? [
      ['new matches after source page one',row=>literal(row) || row.page>=1],
      ['new matches from second source page left',row=>literal(row) || row.page>=2 || row.page===1 && row.lane===0],
      ['new matches from second source page right',row=>literal(row) || row.page>=2 || row.page===1 && row.lane===1],
      ['new matches from first right and later',row=>literal(row) || row.page>=1 || row.page===0 && row.lane===1],
    ] : []),
  ];
  for (const [scope,include] of scopes) {
    if (scope==='title only' && titleAlreadyAuthored) continue;
    const edits = titleAlreadyAuthored ? [] : [[saved.start,saved.start+saved.placement.length,original.placement]];
    const additions = [];
    let nextStyle = 1+Math.max(...[...saved.styles.keys()].map(Number));
    for (const row of proof) {
      if (!include(row)) continue;
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
    try {
      const result = count(changed,scope);
      const boundary = traceLanes && process.env.WINDOWS_HWP_TRACE_SHIFTED_RIGHT === '1' &&
        scope==='new matches after source page two' ? shiftedRightBoundary(changed) : null;
      results.push({ ...result,ownerEdits:additions.length,...(boundary ? { shiftedRightBoundary:boundary } : {}) });
    }
    finally { changed.free(); }
  }
  const longBySourcePage = {};
  if (includeLong) for (const row of proof.filter(row=>row.length>200))
    longBySourcePage[row.page+1] = (longBySourcePage[row.page+1] ?? 0) + 1;
  return { originalPages:source.pageCount(),fixed:proof.length,fixedFirstPage:proof.filter(row=>row.page===0).length,
    titleAlreadyAuthored,...(matchFixedSpaces ? { projectedOwners:proof.filter(row=>row.projected).length } : {}),
    ...(includeLong ? {longBySourcePage} : {}),results };
};
