import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';

const inventory = JSON.parse(readFileSync(process.env.WINDOWS_PUBLIC_QA_INVENTORY!, 'utf8'));
const sample = inventory.find(item => item.id === '02-tables' && item.verified);
if (!sample) throw Error('The exact public Word original is required');

test('Windows Word requested Latin face with installed variants and source language', async ({ browser, baseURL }) => {
  test.setTimeout(180_000);
  const data = [...readFileSync(sample.path)];
  const variants = [
    { label: 'baseline' },
    { label: 'source language', lang: 'en-US' },
    { label: 'Calibri', regular: ['Calibri', 'Calibri-Regular'], italic: ['Calibri Italic', 'Calibri-Italic'] },
    { label: 'Segoe UI', regular: ['Segoe UI', 'SegoeUI'], italic: ['Segoe UI Italic', 'SegoeUI-Italic'] },
    { label: 'Trebuchet MS', regular: ['Trebuchet MS', 'TrebuchetMS'], italic: ['Trebuchet MS Italic', 'TrebuchetMS-Italic'] },
    { label: 'Arial', regular: ['Arial', 'ArialMT'], italic: ['Arial Italic', 'Arial-ItalicMT'] },
    { label: 'Tahoma', regular: ['Tahoma'], italic: ['Tahoma Italic'] },
  ];
  for (const variant of variants) {
    const context = await browser.newContext({ baseURL, viewport: { width: 1440, height: 1000 } });
    try {
      const page = await context.newPage();
      await page.goto('/');
      const result = await page.evaluate(async ({ bytes, variant }) => {
        const { initializeRegularFontFallback } = await import('/src/word-font-fallbacks.ts');
        const { importFile } = await import('/src/document-io.ts');
        const { editorExtensions, parseDocumentHtml } = await import('/src/extensions.ts');
        const { wordLayoutOptions } = await import('/src/word-trailing-spaces.ts');
        const { createPrintLayout } = await import('/src/print-layout.ts');
        const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
        const { paragraphLineOffsets } = await import('/src/paragraph-fragments.ts');
        if (variant.lang) document.documentElement.lang = variant.lang;
        let status: string | undefined;
        if (variant.regular && variant.italic) {
          const local = (names: string[]) => names.map(name => 'local(' + JSON.stringify(name) + ')').join(',');
          const face = (style: 'normal' | 'italic', names: string[]) =>
            ({style, source:local(names), ascent:'92.919921875%', descent:'23.046875%', unicodeRange:'U+20-024F,U+2000-206F'});
          status = await initializeRegularFontFallback({family:'Gill Sans MT',
            nativeNames:['Gill Sans MT','GillSansMT','Gill Sans MT Bold','GillSansMT-Bold','Gill Sans MT Italic','GillSansMT-Italic'],
            faces:[face('normal',variant.regular),face('italic',variant.italic)]});
          if (status !== 'fallback') return {status};
        }
        const file = await importFile(new File([Uint8Array.from(bytes)], 'fixed.docx'));
        const model = { ...file.document, content:parseDocumentHtml(file.html!) };
        const html = generateHTML(model.content, editorExtensions);
        const print = await createPrintLayout(html, model.page, undefined, undefined, { ...wordLayoutOptions(model), nativeDesktop: true });
        try {
          const pages = [...print.element.querySelectorAll<HTMLElement>('.print-page')];
          const chosen = pages[0].querySelectorAll<HTMLElement>('.print-content > p');
          const firstParagraphs = [...chosen].slice(0,12).map(p => ({height:p.getBoundingClientRect().height,
            lines:paragraphLineOffsets(p).length + 1, chars:p.textContent?.replace(/\s+/g,'').length}));
          const counts = pages.map(p => p.querySelectorAll('.print-content tr').length);
          const fonts = [...document.fonts].filter(font => font.family === 'Gill Sans MT').map(font => [font.weight,font.style,font.status]);
          return {status,pages: print.pages,firstParagraphs,counts,fonts};
        } finally { print.dispose(); }
      }, { bytes:data, variant });
      console.log('WINDOWS_WORD_FALLBACK_WITNESS ' + JSON.stringify({ label:variant.label, ...result }));
      if (variant.label === 'baseline') expect(result.pages).toBe(8);
    } finally { await context.close(); }
  }
});
