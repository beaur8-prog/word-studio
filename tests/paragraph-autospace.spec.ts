import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
const fixture = createRequire(import.meta.url)('./hwp-script-cells-fixture.cjs');
let directory: string, files: Record<string, string>;
test.beforeAll(async () => { directory = await fs.mkdtemp(path.join(os.tmpdir(), 'word-autospace-')); files = await fixture(directory); });
test.afterAll(async () => { if (directory) await fs.rm(directory, { recursive: true, force: true }); });

async function savedPolicies(page: any, bytes: number[], extension: string) {
  await page.goto('/');
  return page.evaluate(async ({bytes,extension}) => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let input = new File([new Uint8Array(bytes)], 'policy.' + extension), doc: any;
    if (extension !== 'docx') {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts'), { importEditableHwp } = await import('/src/hwp/editing.ts');
      const session = new HwpReaderSession();
      try { await session.open(input); doc = importEditableHwp(await session.editable(), input.name); }
      finally { session.close(); }
    }
    const states = [], W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    for (let save = 0; save < 3; save++) {
      if (!doc) { const loaded = await importFile(input); doc = {...loaded.document,content:parseDocumentHtml(loaded.html!)}; }
      const blob = await exportDocx(doc), zip = await JSZip.loadAsync(await blob.arrayBuffer());
      const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'),'application/xml');
      states.push([...xml.getElementsByTagNameNS(W,'p')].map(p=>{
        const pp = [...p.children].find(c=>c.localName==='pPr');
        const flag = (name: string) => { const el = pp && [...pp.children].find(c=>c.localName===name); return el ? !['0','false','off'].includes(el.getAttributeNS(W,'val') || '') : true; };
        return {text:[...p.getElementsByTagNameNS(W,'t')].map(t=>t.textContent).join(''),de:flag('autoSpaceDE'),dn:flag('autoSpaceDN'),order:pp ? [...pp.children].map(c=>c.localName) : []};
      }));
      input = new File([blob],'again.docx'); doc = undefined;
    }
    return states;
  }, {bytes,extension});
}

for (const format of ['hwp','hwpx']) test(`${format}: auto spacing stays disabled in body, ordinary and nested cells through two DOCX saves`,async({page})=>{
  const states = await savedPolicies(page,[...await fs.readFile(files[format])],format);
  for (const state of states) for (const text of ['After astral body','Table Latin 456','Nested Latin 789','After nested table']) {
    const p = state.find((p: any)=>p.text.includes(text)); expect(p,text).toBeTruthy();expect(p!.de,text).toBe(false);expect(p!.dn,text).toBe(false);
    for (const later of ['spacing','ind','jc','rPr','sectPr']) if(p!.order.includes(later)) expect(p!.order.indexOf('autoSpaceDN')).toBeLessThan(p!.order.indexOf(later));
    expect(p!.order.indexOf('autoSpaceDE')).toBeLessThan(p!.order.indexOf('autoSpaceDN'));
  }
});

for (const defaults of [false,true]) test(`external DOCX auto-spacing style inheritance, partial overrides and ${defaults?'explicit':'implicit'} defaults survive two saves`,async({page})=>{
  const names = ['BASE','DERIVED','DIRECT-DE','DIRECT-DN','DIRECT-BOTH','DEFAULT'];
  const blob=await Packer.toBuffer(new Document({sections:[{children:names.map(text=>new Paragraph(text))}]}));
  const zip=await JSZip.loadAsync(blob),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  zip.file('word/styles.xml',`<w:styles xmlns:w="${W}">${defaults?'<w:docDefaults><w:pPrDefault><w:pPr><w:autoSpaceDE/><w:autoSpaceDN w:val="false"/></w:pPr></w:pPrDefault></w:docDefaults>':''}<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:pPr><w:autoSpaceDE w:val="false"/><w:autoSpaceDN/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Base"/><w:pPr><w:autoSpaceDN w:val="false"/></w:pPr></w:style></w:styles>`);
  const properties=['<w:pStyle w:val="Base"/>','<w:pStyle w:val="Derived"/>','<w:pStyle w:val="Derived"/><w:autoSpaceDE/>','<w:pStyle w:val="Derived"/><w:autoSpaceDN/>','<w:pStyle w:val="Derived"/><w:autoSpaceDE/><w:autoSpaceDN/>',''];
  zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${names.map((text,i)=>`<w:p><w:pPr>${properties[i]}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`).join('')}<w:sectPr/></w:body></w:document>`);
  const input=await zip.generateAsync({type:'uint8array'}),states=await savedPolicies(page,[...input],'docx');
  for(const state of states) expect(state.filter((p:any)=>p.text).map(({text,de,dn}:any)=>({text,de,dn}))).toEqual([
    {text:'BASE',de:false,dn:true},{text:'DERIVED',de:false,dn:false},{text:'DIRECT-DE',de:true,dn:false},
    {text:'DIRECT-DN',de:false,dn:true},{text:'DIRECT-BOTH',de:true,dn:true},{text:'DEFAULT',de:true,dn:!defaults},
  ]);
});
