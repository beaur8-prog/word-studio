'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');

const signs = document => {
  const result = [];
  for (let section = 0; section < document.getSectionCount(); section++) {
    const paragraphs = [];
    for (let p = 0; p < document.getParagraphCount(section); p++) {
      const text = document.getTextRange(section,p,0,0xffffffff);
      const { controlCount, lines } = JSON.parse(document.getParagraphLineMetrics(section,p,'[]'));
      const { fontSize } = JSON.parse(document.getCharPropertiesAt(section,p,0));
      const { lineSpacing, lineSpacingType, pageBreakBefore, spacingBefore, spacingAfter } = JSON.parse(document.getParaPropertiesAt(section,p));
      const { page } = JSON.parse(document.getPageOfPosition(section,p));
      paragraphs.push({ p, section, text, controlCount, fontSize, lineSpacing, lineSpacingType, pageBreakBefore,
        before:spacingBefore,after:spacingAfter,page:page+1,line:lines?.[0]?.height });
    }
    result.push(paragraphs);
  }
  return result;
};

const clean = ({ text, ...value }) => value;
const geometry = document => Array.from({length:Math.min(document.pageCount(),9)},(_,n)=>{
  const { runs }=JSON.parse(document.getPageTextLayout(n));
  const values=runs.filter(run=>!run.headerFooter&&run.paraIdx!==0xffffffff&&run.text.trim()&&!run.cellPath?.length);
  return [n+1,values.reduce((sum,v)=>sum+v.text.replace(/\s/gu,'').length,0),values[0]?.paraIdx,values[0]?.secIdx];
});

async function main() {
  const root=path.dirname(require.resolve('@rhwp/core'));
  const { default:init,HwpDocument }=await import(pathToFileURL(path.join(root,'rhwp.js')).href);
  await init({module_or_path:await fs.readFile(path.join(root,'rhwp_bg.wasm'))});
  const [before,current]=await Promise.all(process.argv.slice(2).map(async name=>new Uint8Array(await fs.readFile(name))));
  const a=new HwpDocument(before), b=new HwpDocument(current);
  try {
    const old=signs(a), next=signs(b), records=[];
    const sameSequence=old.length===next.length&&old.every((section,i)=>section.length===next[i].length&&
      section.every((value,j)=>value.text===next[i][j].text&&value.controlCount===next[i][j].controlCount));
    for (const [index,section] of old.entries()) {
      const n=next[index]??[];
      for (const [p,value] of section.entries()) {
        const newValue=n[p];
        if (value.text || value.controlCount || !newValue || newValue.text || newValue.controlCount || value.fontSize===newValue.fontSize) continue;
        records.push({ before:value, after:newValue });
      }
    }
    const eligible=records.filter(({before,after})=>before.page<=6&&after.page<=7);
    const trials=[];
    if (sameSequence) {
      for (const group of [eligible,...eligible.map(item=>[item])]) {
        const d=new HwpDocument(current);
        try {
          d.beginBatch();
          try {
            for (const { before:value } of group) {
              const result=JSON.parse(d.applyCharFormat(value.section,value.p,0,0,JSON.stringify({fontSize:value.fontSize})));
              if (!result.ok) throw Error('Diagnostic blank style change rejected');
            }
          } finally { d.endBatch(); }
          if (d.pageCount()!==b.pageCount() || group===eligible) trials.push({
            at:group.map(({before:value})=>[value.section,value.p]),pages:d.pageCount(),geometry:geometry(d),
          });
        } finally {d.free();}
      }
    }
    console.log('WINDOWS_EMPTY_STYLE_BREAKS '+JSON.stringify({before:a.pageCount(),after:b.pageCount(),sameSequence,
      blanks:records.filter(({before})=>before.page<=7).map(({before,after})=>[clean(before),clean(after)]),
      earlier:geometry(a),current:geometry(b),trials}));
  } finally {a.free();b.free();}
}
main().catch(error => { console.error(error);process.exitCode=1; });
