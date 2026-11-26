import { expect, test } from '@playwright/test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const req = createRequire(import.meta.url);

async function initialTable(format: 'hwp' | 'hwpx', stale: boolean) {
  const root = path.dirname(req.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(path.join(root, 'rhwp.js')).href);
  await init({ module_or_path: await fs.readFile(path.join(root, 'rhwp_bg.wasm')) });
  const seed = HwpDocument.createEmpty();
  try {
    seed.createBlankDocument();
    seed.pasteHtml(0, 0, 0, '<table style="width:540px"><tr><td><p>LEFT</p></td><td><p>NAMEPLATE</p></td><td><p>RIGHT</p></td></tr></table><p>AFTER</p>');
    const files = unzipSync(seed.exportHwpx());
    if (stale) {
      const section = 'Contents/section0.xml', header = 'Contents/header.xml';
      let xml = strFromU8(files[section]);
      const tableEnd = xml.indexOf('</hp:tbl>');
      const start = xml.indexOf('<hp:linesegarray>', tableEnd);
      const end = xml.indexOf('</hp:linesegarray>', start) + '</hp:linesegarray>'.length;
      expect(tableEnd).toBeGreaterThan(0);
      expect(start).toBeGreaterThan(tableEnd);
      const line = xml.slice(start, end).match(/<hp:lineseg\b[^>]*>/)?.[0];
      const height = Number(line?.match(/\bvertsize="(\d+)"/)?.[1]);
      const width = Number(line?.match(/\bhorzsize="(\d+)"/)?.[1]);
      expect(height).toBeGreaterThan(0);
      expect(width).toBeGreaterThan(0);
      const make = (pos: number, y: number, ink: number, baseline: number) => `<hp:lineseg textpos="${pos}" vertpos="${y}" vertsize="${height}" textheight="${ink}" baseline="${baseline}" spacing="360" horzpos="0" horzsize="${width}" flags="393216"/>`;
      xml = xml.slice(0, start) + `<hp:linesegarray>${make(0, 0, 600, 510)}${make(16, 960, height, Math.round(height * .85))}</hp:linesegarray>` + xml.slice(end);
      xml = xml.replace(/<hp:outMargin\b[^>]*>/, '<hp:outMargin left="0" right="0" top="283" bottom="0"/>');
      files[section] = strToU8(xml);
      files[header] = strToU8(strFromU8(files[header]).replace(/(<hh:charPr\b[^>]*\bid="0"[^>]*\bheight=")\d+/, '$1600'));
    }
    const hwpx = zipSync(files), authored = new HwpDocument(hwpx);
    try {
      const bytes = format === 'hwp' ? authored.exportHwp() : hwpx;
      const native = new HwpDocument(bytes);
      try {
        const { lines, stored, tables } = JSON.parse(native.getParagraphLineMetrics(0, 0, '[]'));
        expect(stored).toBe(true);
        expect(lines.length).toBe(stale ? 2 : 1);
        if (stale) {
          expect(lines[1].verticalPosition).toBe(12.8);
          const page = JSON.parse(native.getPageDef(0));
          const table = JSON.parse(native.getPageControlLayout(0)).controls.find(c => c.type === 'table' && c.paraIdx === 0);
          // The native compositor does not reserve this stale control line.
          expect(Math.abs(table.y - (page.marginTop + page.marginHeader) / 75 - tables[0].outerTop)).toBeLessThan(.1);
        }
        return [...bytes];
      } finally { native.free(); }
    } finally { authored.free(); }
  } finally { seed.free(); }
}

for (const format of ['hwp', 'hwpx'] as const) for (const stale of [false, true])
  test(`${format}, ${stale ? 'uncorroborated stored lead' : 'ordinary table'} keeps an inline table at the start`, async ({ page }, info) => {
    const bytes = await initialTable(format, stale);
    await page.goto('/');
    const result = await page.evaluate(async ({ bytes, format }) => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { importEditableHwp } = await import('/src/hwp/editing.ts');
      const { exportDocx, importFile } = await import('/src/document-io.ts');
      const { parseDocumentHtml, editorExtensions } = await import('/src/extensions.ts');
      const { createPrintLayout } = await import('/src/print-layout.ts');
      const { generateHTML } = await import('/node_modules/.vite/deps/@tiptap_react.js');
      const original = new HwpReaderSession(), file = new File([Uint8Array.from(bytes)], 'Initial.' + format);
      try {
        await original.open(file);
        let doc = importEditableHwp(await original.editable(), file.name);
        const states = [], saves = [];
        for (let pass = 0; pass < 3; pass++) {
          const layout = await createPrintLayout(generateHTML(doc.content, editorExtensions), doc.page);
          try {
            const body = layout.element.querySelector('.print-content')!, table = body.querySelector('table')!;
            states.push({ pages: layout.pages, first: body.firstElementChild?.tagName, top: table.getBoundingClientRect().top - body.getBoundingClientRect().top,
              text: body.textContent?.replace(/\s/g, '') });
          } finally { layout.dispose(); }
          if (pass === 2) break;
          const blob = await exportDocx(doc);
          saves.push([...new Uint8Array(await blob.arrayBuffer())]);
          const imported = await importFile(new File([blob], 'Again.docx'));
          doc = { ...imported.document, content: parseDocumentHtml(imported.html) };
        }
        return { states, saves };
      } finally { original.close(); }
    }, { bytes, format });
    for (const [i, saved] of result.saves.entries())
      await fs.writeFile(info.outputPath(`textbox-save-lead-${format}-${stale}-${i + 1}.docx`), Buffer.from(saved));
    for (const state of result.states) {
      expect(state.pages).toBe(1);
      expect(state.first).toBe('TABLE');
      expect(Math.abs(state.top)).toBeLessThan(.2);
      expect(state.text).toBe('LEFTNAMEPLATERIGHTAFTER');
    }
  });
