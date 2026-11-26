import { expect, test } from '@playwright/test';

test('HWP continuation restores source lines only after the verified first fragment', async ({ page }) => {
  await page.goto('/');
  const proof = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const labels = Array.from({ length: 7 }, (_, i) => [`Left ${i}`, `Right ${i}`]);
    const text = 'Alpha beta';
    const settings = {
      size: 'Custom', width: 100, height: 70, orientation: 'portrait', margin: 10,
      header: '', footer: '', pageNumbers: false,
    } as const;
    function source() {
      const wrapper = document.createElement('div');
      wrapper.innerHTML = '<table style="width:220px;table-layout:fixed;margin:0;border-collapse:collapse"><colgroup><col style="width:110px"><col style="width:110px"></colgroup></table>';
      const table = wrapper.querySelector('table')!;
      labels.forEach((cells, i) => {
        const row = table.insertRow();
        row.dataset.testRow = String(i);
        row.dataset.rowSplit = 'avoid';
        row.style.height = (i === 0 ? 30 : 55) + 'px';
        for (const label of cells) {
          const cell = document.createElement(i === 0 ? 'th' : 'td');
          cell.style.cssText = 'padding:0;border:0;vertical-align:top';
          const paragraph = document.createElement('p');
          paragraph.style.cssText = 'font:12px/20px Arial;margin:0';
          paragraph.textContent = label;
          cell.append(paragraph); row.append(cell);
        }
      });
      const paragraph = table.rows[4].cells[0].querySelector('p')!;
      paragraph.textContent = text;
      paragraph.dataset.hwpSourceSpaces = 'true';
      paragraph.dataset.testContinuation = 'true';
      return wrapper;
    }
    const before = await createPrintLayout(source().innerHTML, settings);
    const first = before.element.querySelector<HTMLTableElement>('.print-content table')!;
    const body = first.closest('.print-content')!.getBoundingClientRect();
    const cut = Number(first.rows[first.rows.length - 1].dataset.testRow);
    const line = before.element.querySelector<HTMLElement>('[data-test-continuation]')!;
    const sizes = { outer: line.parentElement!.getBoundingClientRect().width, text: line.getBoundingClientRect().width };
    const box = first.getBoundingClientRect();
    const sourcePolicy = { row: cut, rows: labels.length, bodyWidth: body.width, bodyHeight: body.height,
      tableTop: box.top - body.top, cells: labels[cut].map(hwpWrapFingerprint) };
    before.dispose();
    const base = source();
    const candidate = base.querySelector<HTMLElement>('[data-test-continuation]')!;
    candidate.dataset.hwpSourceWrap = JSON.stringify({ fingerprint: hwpWrapFingerprint(text), ends: [5, 9], cellWidth: sizes });
    const trial = async (mode: 'natural' | 'matching' | 'cell edited' | 'measure') => {
      const markup = base.cloneNode(true) as HTMLDivElement;
      const table = markup.querySelector('table')!;
      if (mode !== 'natural') table.dataset.hwpTableContinuation = JSON.stringify(sourcePolicy);
      if (mode === 'cell edited') table.rows[cut].cells[0].querySelector('p')!.textContent = 'Edited';
      const rendered = await createPrintLayout(markup.innerHTML, settings, undefined, undefined,
        { paintHwpCellLines: mode !== 'measure' });
      try {
        const projected = rendered.element.querySelector('[data-test-continuation]')!;
        return { text: projected.textContent, breaks: projected.querySelectorAll('br').length, pages: rendered.pages };
      } finally { rendered.dispose(); }
    };
    return { cut, natural: await trial('natural'), matching: await trial('matching'),
      edited: await trial('cell edited'), measure: await trial('measure') };
  });
  expect(proof.cut).toBeGreaterThan(0);
  expect(proof.cut).toBeLessThan(4);
  expect(proof.matching).toEqual({ ...proof.natural, breaks: 1 });
  expect(proof.natural).toEqual({ ...proof.measure, breaks: 0 });
  expect(proof.edited).toEqual(proof.natural);
});

test('native export uses only the witnessed cell heads in the clear bottom margin', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { createPrintLayout } = await import('/src/print-layout.ts');
    const { hwpOverhangCellFingerprint, nativeOverhangOriginAttribute } = await import('/src/hwp-table-continuation.ts');
    const { normalizeHwpWrapText, hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const settings = { size: 'Custom', width: 100, height: 90, orientation: 'portrait',
      margin: 10, header: '', footer: '', pageNumbers: false } as const;
    const probe = await createPrintLayout('<p>Paper dimensions</p>', settings);
    const paper = probe.element.querySelector('.print-content')!.getBoundingClientRect();
    probe.dispose();
    const originals = ['The first written phrase continues over more lines', 'Short words followed by a longer last phrase'];
    const heads = ['The first', 'Short words'];
    const wrapper = document.createElement('div');
    wrapper.innerHTML = '<table style="width:220px;table-layout:fixed;margin:0;border-collapse:collapse"><colgroup><col style="width:110px"><col style="width:110px"></colgroup></table>';
    const table = wrapper.querySelector('table')!;
    for (const [index, texts] of [['Header A', 'Header B'], ['Tall row A', 'Tall row B'], originals, ['Tail A', 'Tail B']].entries()) {
      const row = table.insertRow();
      row.dataset.testRow = String(index);
      row.dataset.rowSplit = index === 2 ? 'allow' : 'avoid';
      if (index === 1) row.style.height = (paper.height - 30) + 'px';
      for (const text of texts) {
        const cell = document.createElement(index ? 'td' : 'th');
        cell.style.cssText = 'padding:0;border:0;vertical-align:top';
        const p = document.createElement('p');
        p.style.cssText = 'font:12px/20px Arial;margin:0';
        p.textContent = text;
        row.append(cell); cell.append(p);
      }
    }
    const proof = { row:2, rows:4, bodyWidth:paper.width, bodyHeight:paper.height, tableTop:0,
      cells: originals.map(value => hwpOverhangCellFingerprint([value])),
      overhang: { page:1, end:paper.height + 16, defaultHostOrigin:true, heads:heads.map(text => {
        const normalized = normalizeHwpWrapText(text);
        return { end:normalized.length, fingerprint:hwpWrapFingerprint(normalized) };
      }) },
    };
    async function fragments(mode: 'verified' | 'wrong' | 'ordinary') {
      if (mode === 'ordinary') delete table.dataset.hwpTableContinuation;
      else table.dataset.hwpTableContinuation = JSON.stringify(mode === 'verified' ? proof :
        { ...proof, overhang: { ...proof.overhang, heads: [proof.overhang.heads[0],
          { ...proof.overhang.heads[1], fingerprint:'0000000000000000' }] } });
      const layout = await createPrintLayout(wrapper.innerHTML, settings, undefined, undefined,
        { paintHwpCellLines:false, hwpNativeExport:true });
      try {
        return [...layout.element.querySelectorAll<HTMLTableRowElement>('[data-test-row="2"]')].map(row => ({
          page:Number(row.closest<HTMLElement>('.print-page')?.dataset.page),
          text:[...row.cells].map(cell => cell.textContent ?? ''),
          verifiedOrigin: row.closest('table')!.hasAttribute(nativeOverhangOriginAttribute),
        }));
      } finally { layout.dispose(); }
    }
    return { originals, heads, ordinary:await fragments('ordinary'), verified:await fragments('verified'), wrong:await fragments('wrong') };
  });
  expect(result.wrong).toEqual(result.ordinary);
  expect(result.ordinary[0].page).toBeGreaterThan(1);
  expect(result.verified[0].page).toBe(1);
  expect(result.verified.map(part => part.verifiedOrigin)).toEqual([true, ...result.verified.slice(1).map(() => false)]);
  expect(result.verified[0].text.map(text => text.trimEnd())).toEqual(result.heads);
  expect(result.originals.every((value, col) => result.verified.map(part => part.text[col]).join('') === value)).toBe(true);
});
