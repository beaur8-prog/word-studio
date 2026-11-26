import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';

for (const alignment of ['top', 'middle', 'bottom']) for (const continuing of [true, false])
test('repeated header with ' + alignment + ' vertical merge ' + (continuing ? 'continues at rows' : 'keeps its protected block') + ' after preceding text', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const states = await page.evaluate(async ({ alignment, continuing }) => {
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const p = (text: string, height = 20) => '<p style="font-family:Arial;font-size:12pt;line-height:' + height + 'px;margin:0">' + text + '</p>';
    const table = '<table style="width:400px;margin:0"><colgroup><col style="width:140px"><col style="width:260px"></colgroup>' +
      '<tr data-row-split="avoid" style="height:30px"><th colwidth="140" style="padding:2px;border:1px solid black">' + p('HEADER LEFT') + '</th><th colwidth="260" style="padding:2px;border:1px solid black">' + p('HEADER RIGHT') + '</th></tr>' +
      Array.from({length:5}, (_, i) => '<tr data-row-split="' + (continuing ? 'allow' : 'avoid') + '" style="height:80px">' +
        (i ? '' : '<td colwidth="140" rowspan="5" ' + (continuing ? 'data-cell-split="allow"' : '') + ' style="padding:2px;border:1px solid black;vertical-align:' + alignment + '">' + p('LABEL A') + p('LABEL B') + '</td>') +
        '<td colwidth="260" style="padding:2px;border:1px solid black">' + p('BODY ROW ' + i) + '</td></tr>').join('') + '</table>';
    const loaded = await importFile(new File([p('LEAD',450) + table + p('TAIL',200)], 'header-merge.html'));
    let doc = { ...loaded.document, content:parseDocumentHtml(loaded.html!),
      page:{...defaultPage,size:'Custom',width:150,height:180,margin:10,pageNumbers:false}};
    const result = [];
    for (let save=0;save<3;save++) {
      doc.title = 'MERGED HEADER ' + alignment + ' ' + continuing + ' ' + save;
      const layout = await createPrintLayout(generateHTML(doc.content,editorExtensions),doc.page);
      try {
        const starts = [...layout.element.querySelectorAll<HTMLElement>('.print-content p')].filter(p=>!p.closest('th')).map(p=>{
          const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);walker.nextNode();
          const range=document.createRange();range.setStart(walker.currentNode,0);range.setEnd(walker.currentNode,1);
          const box=range.getBoundingClientRect(),paper=p.closest<HTMLElement>('.print-page')!,origin=paper.getBoundingClientRect();
          return {text:p.textContent!,page:Number(paper.dataset.page),y:box.top-origin.top};
        });
        result.push({doc,pages:layout.pages,starts,headers:layout.element.querySelectorAll('th').length});
      } finally {layout.dispose();}
      if (save<2) {const next=await importFile(new File([await exportDocx(doc)],'saved.docx'));doc={...next.document,content:parseDocumentHtml(next.html!)};}
    }
    return result;
  }, {alignment,continuing});
  for (const state of states) {
    // Each saved input is a fresh user editing session. Retaining the view
    // across rapid imports can turn three independent single clicks into
    // ProseMirror's cell-selecting triple click. That gesture has its own test.
    if (state !== states[0]) await page.reload();
    expect(state.starts).toHaveLength(9);
    expect(state.starts.map(p=>p.text).sort()).toEqual(['LEAD','TAIL','LABEL A','LABEL B',...Array.from({length:5},(_,i)=>'BODY ROW '+i)].sort());
    expect(state.pages).toBe(continuing ? 2 : 3);
    expect(state.starts.find(p=>p.text==='BODY ROW 0')?.page).toBe(continuing ? 1 : 2);
    await page.locator('input[type=file]').first().setInputFiles({name:'header-merge.wordstudio',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(state.doc))});
    await expect(page.getByRole('textbox',{name:'문서 제목',exact:true})).toHaveValue(state.doc.title);
    for (const zoom of [100,50,125,200,100]) {
      const slider=page.getByRole('slider',{name:'확대 축소'});await slider.fill(String(zoom));await slider.dispatchEvent('input');await slider.dispatchEvent('change');
      await expect.poll(()=>page.evaluate(({zoom,expected})=>{
        const root=document.querySelector<HTMLElement>('.tiptap[contenteditable=true]')!,sheet=root.closest<HTMLElement>('.document-stack')!;
        const origin=root.getBoundingClientRect().top,scale=zoom/100,stride=Number(sheet.dataset.pageHeight)+24;
        const actual=[...root.querySelectorAll<HTMLElement>('p')].filter(p=>!p.closest('th,.ProseMirror-widget,.pagination-headers,.section-papers')).map(p=>{
          const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);walker.nextNode();
          if(walker.currentNode===p)return null;
          const range=document.createRange();range.setStart(walker.currentNode,0);range.setEnd(walker.currentNode,1);
          const yy=(range.getBoundingClientRect().top-origin)/scale,page=Math.max(1,1+Math.floor((yy+.5)/stride));
          return {text:p.textContent!,page,y:yy-(page-1)*stride};
        }).filter(Boolean);
        return {pages:Number(sheet.dataset.pages),problems:actual.flatMap(a=>{
          const e=expected.find(e=>e.text===a!.text);return !e||e.page!==a!.page||Math.abs(e.y-a!.y)>1.25?[{actual:a,expected:e}]:[];
        })};
      },{zoom,expected:state.starts})).toEqual({pages:state.pages,problems:[]});
    }
    const editor=page.getByRole('textbox',{name:'문서 편집 영역',exact:true}),before=await editor.evaluate(el=>(el as any).editor.getJSON());
    await editor.locator('td p').filter({hasText:'BODY ROW 0'}).click();await page.keyboard.press('End');await page.keyboard.insertText(' EDIT');
    await expect(editor).toContainText('BODY ROW 0 EDIT');await page.keyboard.press('ControlOrMeta+z');
    await expect.poll(()=>editor.evaluate(el=>(el as any).editor.getJSON())).toEqual(before);
    await page.keyboard.press('ControlOrMeta+Shift+z');await expect(editor).toContainText('BODY ROW 0 EDIT');
    await page.keyboard.press('ControlOrMeta+z');await expect.poll(()=>editor.evaluate(el=>(el as any).editor.getJSON())).toEqual(before);
  }
});

const WORD_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const mergedLabels = Array.from({ length: 20 }, (_, i) => 'MERGE-' + String(i + 1).padStart(2, '0'));
const rowLabels = Array.from({ length: 5 }, (_, i) => 'ROW-' + String(i + 1).padStart(2, '0'));
const headerLabels = ['HEADER MERGE', 'HEADER DATA'];

async function standardRepeatedMergeDocx() {
  const paragraph = (text: string, line = 300) =>
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="' + line + '" w:lineRule="exact"/></w:pPr>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t>' + text + '</w:t></w:r></w:p>';
  const cell = (width: number, content: string, merge = '') =>
    '<w:tc><w:tcPr><w:tcW w:w="' + width + '" w:type="dxa"/>' + merge + '</w:tcPr>' + content + '</w:tc>';
  const continuationRows = rowLabels.map((label, index) =>
    '<w:tr><w:trPr><w:cantSplit w:val="0"/><w:trHeight w:val="2100" w:hRule="atLeast"/></w:trPr>' +
    (index === 0
      ? cell(2100, mergedLabels.map(label => paragraph(label)).join(''), '<w:vMerge w:val="restart"/>')
      : cell(2100, '<w:p/>', '<w:vMerge/>')) + cell(3900, paragraph(label)) + '</w:tr>'
  ).join('');
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side =>
    '<w:' + side + ' w:val="single" w:sz="6" w:color="000000"/>'
  ).join('');
  const margins = ['top', 'left', 'bottom', 'right'].map(side =>
    '<w:' + side + ' w:w="30" w:type="dxa"/>'
  ).join('');
  const table = '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>' + borders +
    '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar>' + margins + '</w:tblCellMar></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="2100"/><w:gridCol w:w="3900"/></w:tblGrid>' +
    '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="450" w:hRule="atLeast"/><w:tblHeader/></w:trPr>' +
    cell(2100, paragraph(headerLabels[0])) + cell(3900, paragraph(headerLabels[1])) + '</w:tr>' +
    continuationRows + '</w:tbl>';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="' + WORD_NAMESPACE + '"><w:body>' +
    paragraph('LEAD', 3300) + table + paragraph('TAIL') +
    '<w:sectPr><w:pgSz w:w="8504" w:h="10205"/><w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="567" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>' +
    '</w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

function placedMergeMarkers(root: HTMLElement, printing: boolean) {
  const sheet = printing ? root : root.closest<HTMLElement>('.document-stack')!;
  const bodyOrigin = root.getBoundingClientRect();
  const paperHeight = Number(sheet.dataset.pageHeight);
  const stride = paperHeight + 24;
  const selected = [...root.querySelectorAll<HTMLElement>('td p')].filter(paragraph =>
    /^(?:MERGE|ROW)-\d{2}$/.test(paragraph.textContent || '') && !paragraph.closest('.ProseMirror-widget')
  );
  const markers = selected.map(paragraph => {
    const range = document.createRange();
    range.selectNodeContents(paragraph);
    const ink = range.getBoundingClientRect();
    const cell = paragraph.closest('td')!.getBoundingClientRect();
    const paper = paragraph.closest<HTMLElement>('.print-page');
    const page = paper ? Number(paper.dataset.page) : 1 + Math.floor((ink.top - bodyOrigin.top + 0.5) / stride);
    const origin = paper?.getBoundingClientRect() ?? bodyOrigin;
    const y = ink.top - origin.top - (paper ? 0 : (page - 1) * stride);
    const clip = paper
      ? paragraph.closest('.print-content')!.getBoundingClientRect()
      : { top: origin.top + (page - 1) * stride + Number(sheet.dataset.margin),
          bottom: origin.top + (page - 1) * stride + paperHeight - Number(sheet.dataset.bottomMargin) };
    const header = paper?.querySelector('tr')?.getBoundingClientRect();
    return {
      text: paragraph.textContent!,
      page,
      y,
      inside: ink.width > 0 && ink.height > 0 && ink.left >= cell.left - 1 && ink.right <= cell.right + 1 &&
        ink.top >= Math.max(cell.top, clip.top, header?.bottom ?? 0) - 1 && ink.bottom <= Math.min(cell.bottom, clip.bottom) + 1,
    };
  });
  return { markers, pages: printing ? root.querySelectorAll('.print-page').length : Number(sheet.dataset.pages) };
}

test('standard DOCX repeating header and body merge cross a page through two saves', async ({ page }, info) => {
  await page.goto('/');
  let bytes = await standardRepeatedMergeDocx();
  const observations = [];
  for (let pass = 0; pass < 3; pass++) {
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    const structure = await page.evaluate(({ xml, word }) => {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const direct = (parent: Element | undefined, name: string) => Array.from(parent?.children ?? [])
        .filter(child => child.namespaceURI === word && child.localName === name);
      const body = doc.getElementsByTagNameNS(word, 'body')[0];
      const tables = direct(body, 'tbl');
      const grid = direct(direct(tables[0], 'tblGrid')[0], 'gridCol').map(column => column.getAttributeNS(word, 'w'));
      const rows = direct(tables[0], 'tr').map(row => {
        const props = direct(row, 'trPr')[0];
        const header = direct(props, 'tblHeader').some(tag => !['0', 'false', 'off'].includes(tag.getAttributeNS(word, 'val') ?? ''));
        const split = direct(props, 'cantSplit')[0]?.getAttributeNS(word, 'val');
        const cells = direct(row, 'tc').map(cell => {
          const merge = direct(direct(cell, 'tcPr')[0], 'vMerge')[0];
          const text = [...cell.getElementsByTagNameNS(word, 't')].map(text => text.textContent);
          return { merge: merge ? merge.getAttributeNS(word, 'val') || 'continue' : null, text };
        });
        return { header, allowsSplit: ['0', 'false', 'off'].includes(split ?? ''), cells };
      });
      return { tables: tables.length, grid, rows };
    }, { xml, word: WORD_NAMESPACE });
    expect(structure).toEqual({
      tables: 1,
      grid: ['2100', '3900'],
      rows: [
        { header: true, allowsSplit: false, cells: headerLabels.map(text => ({ merge: null, text: [text] })) },
        ...rowLabels.map((text, i) => ({
          header: false,
          allowsSplit: true,
          cells: [
            { merge: i === 0 ? 'restart' : 'continue', text: i === 0 ? mergedLabels : [] },
            { merge: null, text: [text] },
          ],
        })),
      ],
    });

    const name = 'standard-repeated-merge-' + pass;
    await page.locator('input[type=file]').first().setInputFiles({
      name: name + '.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: bytes,
    });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(name);
    const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true });
    await expect(editor.locator('th')).toHaveText(headerLabels);
    await expect(editor.locator('td[rowspan="5"] p')).toHaveText(mergedLabels);
    await expect(page.locator('.pagination-headers table th')).toHaveText(headerLabels);
    const edit = await editor.evaluate(placedMergeMarkers, false);

    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    const printRoot = page.locator('.print-layout');
    await expect(printRoot.locator('.print-page')).toHaveCount(2);
    for (const sheet of await printRoot.locator('.print-page').all()) {
      const table = sheet.locator('.print-content table');
      await expect(table).toHaveCount(1);
      await expect(table.locator('tr').first().locator('th')).toHaveText(headerLabels);
      await expect(table.locator('th')).toHaveCount(headerLabels.length);
    }
    const print = await printRoot.evaluate(placedMergeMarkers, true);
    await page.getByRole('button', { name: '편집으로 돌아가기', exact: true }).click();
    if (pass < 2) {
      await page.getByRole('button', { name: '파일', exact: true }).click();
      const download = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Word 문서 (.docx) Microsoft Word에서 열고 편집', exact: true }).click();
      const saved = info.outputPath('standard-repeated-merge-save-' + (pass + 1) + '.docx');
      await (await download).saveAs(saved);
      bytes = await readFile(saved);
    }
    observations.push({ edit, print });
  }

  await writeFile(info.outputPath('repeated-merge-stages.json'), JSON.stringify(observations, null, 2));
  for (const { edit, print } of observations) {
    expect(edit.pages).toBe(2);
    expect(print.pages).toBe(2);
    for (const mode of [edit, print]) {
      const merged = mode.markers.filter(marker => marker.text.startsWith('MERGE-'));
      const rows = mode.markers.filter(marker => marker.text.startsWith('ROW-'));
      expect(merged.map(marker => marker.text)).toEqual(mergedLabels);
      expect([...new Set(merged.map(marker => marker.page))]).toEqual([1, 2]);
      expect(rows.map(marker => marker.text)).toEqual(rowLabels);
      expect(rows.map(marker => marker.page)).toEqual([1, 1, 2, 2, 2]);
      expect(mode.markers.filter(marker => !marker.inside)).toEqual([]);
    }
    const edited = new Map(edit.markers.map(marker => [marker.text, marker]));
    for (const marker of print.markers) {
      const comparable = edited.get(marker.text)!;
      expect(comparable.page, marker.text).toBe(marker.page);
      expect(Math.abs(comparable.y - marker.y), marker.text).toBeLessThan(1.25);
    }
    expect(print.markers.map(({ text, page }) => [text, page])).toEqual(
      observations[0].print.markers.map(({ text, page }) => [text, page])
    );
  }
});
