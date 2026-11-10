import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';

test('missing Malgun preserves natural Hangul advances and authored fallback lists', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const result = await page.evaluate(async () => {
    const available = async (source: string) => {
      try { await new FontFace('Gothic availability probe', source).load(); return true; } catch { return false; }
    };
    const installed = await available('local("Malgun Gothic"),local("맑은 고딕")');
    const narrow = document.documentElement.classList.contains('narrow-korean-gothic');
    const originalLegacy = await available('local("한양중고딕")');
    const host = document.createElement('div'); host.className = 'tiptap'; document.body.append(host);
    const span = (family: string | null, text: string) => {
      const node = document.createElement('span');
      node.style.cssText = 'display:inline-block;white-space:pre;font-size:20px;letter-spacing:normal;font-weight:400';
      if (family) node.style.fontFamily = family;
      node.textContent = text; host.append(node); return node;
    };
    try {
      const expectedHangul = span('Noto Sans KR', '가나다바사아'), expectedLatin = span('Arial', 'Abc123');
      const hangul = [null, 'Malgun Gothic', '맑은 고딕'].map(f => span(f, '가나다바사아'));
      const latin = [null, 'Malgun Gothic', '맑은 고딕'].map(f => span(f, 'Abc123'));
      const legacy = span('한양중고딕', '가나다바사아');
      const authored = span('"Malgun Gothic", monospace', '가나다');
      await document.fonts.load('20px "Noto Sans KR"', '가나다바사아Abc123');
      await document.fonts.ready;
      return { installed, narrow, originalLegacy, legacyWidth:legacy.getBoundingClientRect().width, legacyFamily:getComputedStyle(legacy).fontFamily, legacySource:legacy.style.fontFamily, hangul: hangul.map(n => n.getBoundingClientRect().width),
        latin: latin.map(n => n.getBoundingClientRect().width), expectedHangul: expectedHangul.getBoundingClientRect().width,
        expectedLatin: expectedLatin.getBoundingClientRect().width,
        authored: getComputedStyle(authored).fontFamily, source: authored.style.fontFamily };
    } finally { host.remove(); }
  });
  expect(result.authored).toBe(result.source);
  if (result.narrow && !result.originalLegacy) expect(result.legacyWidth).toBeCloseTo(result.expectedHangul,2);
  if (!result.narrow) expect(result.legacyFamily).toBe(result.legacySource);
  if (!result.installed && result.narrow) {
    for (const width of result.hangul) expect(width).toBeCloseTo(result.expectedHangul, 2);
    for (const width of result.latin) expect(width).toBeCloseTo(result.expectedLatin, 2);
  }
});

test('an available requested face remains ahead of the optional Korean Gothic fallback', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('textbox', { name: '문서 편집 영역', exact: true }).waitFor();
  const result = await page.evaluate(async () => {
    const family = 'Malgun Gothic', text = 'Abc123';
    const face = new FontFace(family, 'local("Courier New"),local("Liberation Mono"),local("DejaVu Sans Mono")');
    await face.load(); document.fonts.add(face);
    const host = document.createElement('div'); host.className = 'tiptap'; document.body.append(host);
    try {
      const node = document.createElement('span'); node.style.cssText = 'font-family:"Malgun Gothic";font-size:20px;white-space:pre';
      node.textContent = text; host.append(node);
      const canvas = document.createElement('canvas'), ctx = canvas.getContext('2d')!;
      ctx.font = '20px "Malgun Gothic"';
      const expected = ctx.measureText(text).width;
      return { actual: node.getBoundingClientRect().width, expected, family: getComputedStyle(node).fontFamily };
    } finally { host.remove(); document.fonts.delete(face); }
  });
  expect(result.family.split(',')[0]).toBe('"Malgun Gothic"');
  expect(result.actual).toBeCloseTo(result.expected, 2);
});

test('Korean Gothic edit and two DOCX saves retain requested fonts and original sizes', async ({ page }) => {
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({ name: 'gothic.html', mimeType: 'text/html',
    buffer: Buffer.from('<p style="font-family:Malgun Gothic;font-size:12pt;line-height:18pt">한글 서체와 문서 123</p>') });
  const editor = page.getByRole('textbox', { name: '문서 편집 영역', exact: true }), p = editor.locator('p').first();
  await p.click(); await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowRight' : 'End');
  await page.keyboard.type(' EDIT'); await expect(p).toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+z'); await expect(p).not.toContainText('EDIT');
  await page.keyboard.press('ControlOrMeta+Shift+z'); await expect(p).toContainText('EDIT');
  const rows = await editor.evaluate(async (root: any) => {
    const { exportDocx, importFile } = await import('/src/document-io.ts');
    const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
    const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { defaultPage } = await import('/src/types.ts');
    const { default: JSZip } = await import('/node_modules/.vite/deps/jszip.js');
    let model: any = { format:'word-studio',version:1,id:'gothic',title:'Gothic',updatedAt:new Date().toISOString(),
      page:defaultPage,comments:[],content:root.editor.getJSON() };
    const states = [];
    for (let save=0; save<3; save++) {
      const html = generateHTML(model.content, editorExtensions), layout = await createPrintLayout(html, model.page);
      let printed: any;
      try {
        const p = layout.element.querySelector('.print-content p')!;
        const text = p.querySelector('span') || p;
        printed = { text:p.textContent,pages:layout.pages,size:getComputedStyle(text).fontSize };
      } finally { layout.dispose(); }
      const blob=await exportDocx(model), bytes=[...new Uint8Array(await blob.arrayBuffer())];
      const zip=await JSZip.loadAsync(await blob.arrayBuffer()), xml=await zip.file('word/document.xml')!.async('string');
      states.push({...printed,xml,bytes});
      const opened=await importFile(new File([blob],'gothic.docx'));
      model={...opened.document,content:parseDocumentHtml(opened.html!)};
    }
    return states;
  });
  for (const [index,row] of rows.entries()) {
    await writeFile(test.info().outputPath('textbox-save-korean-gothic-'+index+'.docx'),Buffer.from(row.bytes));
    expect(row.text).toBe('한글 서체와 문서 123 EDIT'); expect(row.pages).toBe(1); expect(row.size).toBe('16px');
    expect(row.xml).toContain('Malgun Gothic');
    expect(row.xml).not.toMatch(/AppleGothic|Noto Sans|Studio Korean Gothic/);
  }
});

test('a system without the measured narrow face retains its native fallback selection', async ({page})=>{
  await page.goto('/');
  await page.getByRole('textbox',{name:'문서 편집 영역',exact:true}).waitFor();
  const result=await page.evaluate(async()=>{
    const {initializeNarrowKoreanGothic}=await import('/src/word-font-fallbacks.ts');
    const Original=globalThis.FontFace,root=document.documentElement,previous=root.classList.contains('narrow-korean-gothic');
    (globalThis as any).FontFace=function(family:string,source:string,options:any){
      return new Original(family,source.includes('AppleSDGothicNeo-Regular')?'local("WS Definitely Missing Native Face")':source,options);
    };
    try{
      const available=await initializeNarrowKoreanGothic(()=>{});
      const host=document.createElement('div');host.className='tiptap';document.body.append(host);
      const span=document.createElement('span');span.style.fontFamily='한양중고딕';span.textContent='Native fallback';host.append(span);
      try{return {available,active:root.classList.contains('narrow-korean-gothic'),family:getComputedStyle(span).fontFamily,source:span.style.fontFamily};}
      finally{host.remove();}
    }finally{globalThis.FontFace=Original;root.classList.toggle('narrow-korean-gothic',previous);}
  });
  expect(result.available).toBe(false);expect(result.active).toBe(false);expect(result.family).toBe(result.source);
});
