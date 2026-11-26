import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';

const probeFamily = 'Studio Apple Gothic space probe';
const faceAsset = '/src/assets/fonts/texgyreadventor-regular.otf';
const linePrefix = 'One two three four five six seven';
const scenarios = ['right-slot', 'explicit-wrap', 'mixed-run', 'two-slots', 'legacy-studio', 'later-two-slots', 'spans-paper', 'footnote-context'] as const;
type Scenario = typeof scenarios[number];

function fixtureSource(scenario: Scenario) {
  const compound = scenario === 'later-two-slots' || scenario === 'spans-paper';
  const text = `${linePrefix} arrives after the picture wraps this paragraph.`;
  const tail = ' Source text remains editable around all pictures and preserves ordinary spaces.';
  return {
    scenario,
    compound,
    firstPicture: compound ? 'LEFT ART' : 'WRAP ART',
    middlePicture: scenario === 'later-two-slots' ? 'MIDDLE ART' : '',
    text: compound ? text + tail.repeat(scenario === 'spans-paper' ? 74 : 4) : text,
  };
}

async function emulateMatchingCalibriFallback(page: Page) {
  await page.addInitScript(({ probeFamily, faceAsset }) => {
    const BrowserFontFace = window.FontFace;
    window.FontFace = class extends BrowserFontFace {
      constructor(family: string, source: string | BufferSource, options?: FontFaceDescriptors) {
        if (family === 'Studio Calibri presence probe') source = 'local("WS deliberately unavailable Calibri")';
        if (family === probeFamily) source = `url("${faceAsset}")`;
        super(family, source, options);
      }
    };
    // Model a matching probe on any host. Document text uses its original CSS font.
    const measure = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function () {
      if (!(this instanceof HTMLSpanElement)) return measure.call(this);
      const original = this.style.fontFamily;
      if (this.style.position !== 'fixed' || this.style.visibility !== 'hidden' || this.style.fontSize !== '64px' ||
          !['Wm', 'i', '0', ' '].includes(this.textContent ?? '') || original.replace(/["']/g, '') !== 'Calibri')
        return measure.call(this);
      this.style.fontFamily = JSON.stringify(probeFamily);
      try { return measure.call(this); }
      finally { this.style.fontFamily = original; }
    };
  }, { probeFamily, faceAsset });
}

async function makeFixture(page: Page, source: ReturnType<typeof fixtureSource>) {
  const bytes = await page.evaluate(async ({ source, linePrefix, probeFamily }) => {
    const { scenario, compound, firstPicture, middlePicture, text } = source;
    const { importFile, exportDocx } = await import('/src/document-io.ts');
    const { parseDocumentHtml } = await import('/src/extensions.ts');
    const { defaultPage, pageMargins, paperDimensions } = await import('/src/types.ts');
    // The prefix fits this slot after measured space correction. The same
    // prefix plus its final U+0020 does not fit, nor does the unadjusted prefix.
    const sample = document.createElement('span');
    sample.style.cssText = 'position:fixed;visibility:hidden;white-space:pre;font:16px Calibri;font-variant-ligatures:none;font-feature-settings:"liga" 0';
    document.body.append(sample);
    const width = (text: string) => { sample.textContent = text; return sample.getBoundingClientRect().width; };
    const original = width(linePrefix), originalSpace = width(' ');
    sample.style.fontFamily = JSON.stringify(probeFamily);
    const delta = 16 * 463 / 2048 - width(' ');
    sample.remove();
    const corrected = original + delta * (linePrefix.split(' ').length - 1);
    const slot = Math.ceil(corrected + .35);
    if (slot >= Math.min(original, corrected + originalSpace + delta) - .15)
      throw Error(`Fixture cannot distinguish the fit rules: slot ${slot}, corrected ${corrected}, original ${original}, space ${originalSpace}, delta ${delta}`);

    const settings = { ...defaultPage, pageNumbers: false }, margins = pageMargins(settings), px = 96 / 25.4;
    const left = margins.left * px, right = (paperDimensions(settings)[0] - margins.right) * px;
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const picture = (name: string, offsetX: number, offsetY: number, height: number) => {
      const start = left + offsetX;
      const geometry = {
        horizontalPosition: { relative: 'page', offset: Math.round(start * 9525) },
        verticalPosition: { relative: 'page', offset: Math.round((margins.top * px + offsetY) * 9525) },
        wrap: { type: 1, side: 'bothSides' }, layoutInCell: true,
      };
      const data = JSON.stringify(geometry).replaceAll('"', '&quot;');
      return `<img data-inline-image="true" width="${Math.round(right - start - slot)}" height="${height}" alt="${name}" src="data:image/png;base64,${png}" data-image-floating="${data}">`;
    };
    const firstHeight = scenario === 'later-two-slots' ? 29 : scenario === 'spans-paper' ? 128 : 120;
    const first = picture(firstPicture, scenario === 'two-slots' ? 56 : 0, 0, firstHeight);
    const middle = middlePicture ? picture(middlePicture, 88, 54, 112) : '';
    const run = scenario === 'mixed-run' ? text.replace('seven', '<span style="font-family:Georgia">seven</span>') : text;
    const rule = compound ? 'exact' : 'auto', line = compound ? '15pt' : '1';
    const body = `<p data-word-paragraph-style="Normal" data-word-line-rule="${rule}" data-word-line="${line}" style="font-family:Calibri;font-size:12pt;${compound ? 'line-height:20px;' : ''}text-align:left;margin:0">${run}</p>`;
    const html = compound
      ? `<p data-word-paragraph-style="Normal" data-word-line-rule="exact" data-word-line=".75pt" style="font-size:1pt;line-height:1px;margin:0">${first}${middle}</p>${body}`
      : `${body}<p>${first}</p>`;
    // The reference belongs to another paragraph, so the target itself is a
    // plain body paragraph that would otherwise qualify for correction.
    const noteStory = scenario === 'footnote-context'
      ? '<p>Separate reference<span data-note-reference="foot" data-note-kind="footnote" data-note-label="1">1</span></p>' +
        '<div data-document-note="foot" data-note-kind="footnote" data-note-label="1"><p>Footnote body remains visible.</p></div>'
      : '';
    const opened = await importFile(new File([html + noteStory], 'flow.html'));
    const doc = { ...opened.document, page: settings, content: parseDocumentHtml(opened.html!),
      ...(scenario === 'legacy-studio' ? {} : { wordDefaultHangingSpaces: true as const }),
      ...(scenario === 'explicit-wrap' ? { wordWrapTrailSpaces: true } : {}) };
    return [...new Uint8Array(await (await exportDocx(doc)).arrayBuffer())];
  }, { source, linePrefix, probeFamily });
  const buffer = Buffer.from(bytes);
  if (source.scenario === 'legacy-studio') {
    const zip = await JSZip.loadAsync(buffer);
    const core = await zip.file('docProps/core.xml')!.async('string');
    expect(core).toContain('<dc:creator>Word Studio</dc:creator>');
    const custom = await zip.file('docProps/custom.xml')?.async('string');
    expect(custom ?? '').not.toContain('WordStudio.DefaultTrailingSpaces');
  }
  if (source.scenario === 'explicit-wrap') {
    const zip = await JSZip.loadAsync(buffer);
    const config = await zip.file('word/settings.xml')!.async('string');
    expect(config).toContain('<w:wrapTrailSpaces');
    const explicit = config.replace(/<w:wrapTrailSpaces(?:\s[^>]*)?\s*\/>/, '<w:wrapTrailSpaces w:val="true"/>');
    expect(explicit).toContain('<w:wrapTrailSpaces w:val="true"/>');
    zip.file('word/settings.xml', explicit);
    return zip.generateAsync({ type: 'nodebuffer' });
  }
  return buffer;
}

async function readLayout(root: Locator, print: boolean, firstPicture: string, middlePicture: string) {
  return root.evaluate((root, { print, firstPicture, middlePicture }) => {
    const paragraphs = [...root.querySelectorAll('p')].filter(p => p.textContent?.trim() &&
      (!print || p.closest('.print-content')) && !p.closest('[data-document-note]') && !p.querySelector('[data-note-reference]'));
    const primary = root.querySelector(`img[alt="${firstPicture}"]`)!.getBoundingClientRect();
    const middle = middlePicture ? root.querySelector(`img[alt="${middlePicture}"]`)!.getBoundingClientRect() : undefined;
    const stack = root.closest<HTMLElement>('.document-stack');
    const stackY = stack?.getBoundingClientRect().y ?? 0;
    const stride = Number(stack?.dataset.pageHeight) + 24;
    const lines: { y: number; glyphs: { text: string; box: DOMRect }[] }[] = [];
    const pages = new Set<number>();
    for (const p of paragraphs) {
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const run = walker.currentNode as Text;
        for (let index = 0; index < run.length; index++) {
          if (run.data[index] === ' ') continue;
          const range = document.createRange();
          range.setStart(run, index); range.setEnd(run, index + 1);
          const box = range.getBoundingClientRect();
          const glyph = { text: run.data[index], box };
          const last = lines.at(-1);
          if (last && Math.abs(last.y - box.y) <= .5) last.glyphs.push(glyph);
          else lines.push({ y: box.y, glyphs: [glyph] });
          pages.add(print ? Number(p.closest<HTMLElement>('.print-page')?.dataset.page)
            : Math.max(1, 1 + Math.floor((box.y - stackY) / stride)));
        }
      }
    }
    // Measure visible characters only. Spacer widgets do not have glyph bounds.
    const intersectsBand = (glyph: DOMRect, image: DOMRect) => glyph.bottom > image.top + .5 && glyph.top < image.bottom - .5;
    const overlaps = (glyph: DOMRect, image: DOMRect) => intersectsBand(glyph, image) && glyph.right > image.left + .5 && glyph.left < image.right - .5;
    const firstLine = lines[0]?.glyphs ?? [];
    const bodyBox = paragraphs[0]?.getBoundingClientRect();
    const laterSplit = middle && lines.slice(1).some(({ glyphs }) =>
      glyphs.some(({ box }) => intersectsBand(box, middle) && box.right <= middle.left + .5) &&
      glyphs.some(({ box }) => intersectsBand(box, middle) && box.left >= middle.right - .5));
    return {
      text: paragraphs.map(p => p.textContent ?? '').join(''),
      pages: [...pages].sort((a, b) => a - b),
      firstLine: firstLine.map(glyph => glyph.text).join(''),
      firstLeft: firstLine.some(({ box }) => box.right <= primary.left + .5),
      firstRight: firstLine.some(({ box }) => box.left >= primary.right - .5),
      firstIsOneSlot: !!(bodyBox && firstLine.length && Math.abs(primary.left - bodyBox.left) <= 1 &&
        firstLine.every(({ box }) => box.left >= primary.right - .5 && intersectsBand(box, primary))),
      freeLeft: primary.left - (bodyBox?.left ?? 0), freeRight: (bodyBox?.right ?? 0) - primary.right,
      laterSplit: !!laterSplit,
      overlapsArt: lines.some(({ glyphs }) => glyphs.some(({ box }) => overlaps(box, primary) || (middle && overlaps(box, middle)))),
      widgets: paragraphs.reduce((count, p) => count + p.querySelectorAll('[data-flow-widget]').length, 0),
      styles: paragraphs.map(p => ({
        enabled: p.hasAttribute('data-word-floating-spaces'),
        adjustment: parseFloat(getComputedStyle(p).wordSpacing) || 0,
        wrap: getComputedStyle(p).whiteSpace,
      })),
    };
  }, { print, firstPicture, middlePicture });
}

for (const scenario of scenarios) {
  test(`Word float-space eligibility: ${scenario}`, async ({ page }) => {
    await emulateMatchingCalibriFallback(page);
    await page.goto('/');
    await expect.poll(() => page.evaluate(family => [...document.fonts].some(face =>
      face.family.replace(/["']/g, '') === family && face.status === 'loaded'), probeFamily)).toBe(true);
    const source = fixtureSource(scenario);
    await page.locator('input[type=file]').first().setInputFiles({ name: `${scenario}.docx`,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await makeFixture(page, source) });
    await expect(page.getByRole('textbox', { name: '문서 제목', exact: true })).toHaveValue(scenario);

    const verify = async (root: Locator, print: boolean) => {
      const { firstPicture, middlePicture, compound } = source;
      await expect(root.locator(`img[alt="${firstPicture}"]`)).toBeAttached();
      if (middlePicture) await expect(root.locator(`img[alt="${middlePicture}"]`)).toBeAttached();
      if (scenario === 'footnote-context') {
        await expect(root.locator('[data-note-reference]')).toBeAttached();
        await expect(root.locator('[data-document-note]')).toContainText('Footnote body remains visible.');
      }
      const read = () => readLayout(root, print, firstPicture, middlePicture);
      await expect.poll(async () => {
        const state = await read();
        return compound || scenario === 'footnote-context' ? state.firstIsOneSlot : state.widgets > 0;
      }).toBe(true);
      const state = await read();
      expect(state.text).toBe(source.text);
      expect(state.overlapsArt).toBe(false);

      if (scenario === 'right-slot') {
        expect(state.firstLine).toBe(linePrefix.replaceAll(' ', ''));
        expect(state.styles).toHaveLength(1);
        expect(state.styles[0].enabled).toBe(true);
        expect(state.styles[0].adjustment).toBeLessThan(0);
        expect(state.styles[0].wrap).toBe('pre-wrap');
      } else {
        expect(state.styles.every(style => !style.enabled && style.adjustment === 0 && style.wrap === 'break-spaces')).toBe(true);
        if (scenario === 'explicit-wrap') expect(state.firstLine).not.toBe(linePrefix.replaceAll(' ', ''));
      }
      if (scenario === 'two-slots') {
        expect(state.freeLeft).toBeGreaterThan(40);
        expect(state.freeRight).toBeGreaterThan(120);
        expect(state.firstLeft).toBe(true);
        expect(state.firstRight).toBe(true);
      }
      if (scenario === 'later-two-slots') expect(state.laterSplit).toBe(true);
      if (scenario === 'spans-paper') {
        expect(state.pages[0]).toBe(1);
        expect(state.pages.length).toBeGreaterThan(1);
      }
    };
    await verify(page.getByRole('textbox', { name: '문서 편집 영역', exact: true }), false);
    await page.getByRole('button', { name: '보기', exact: true }).click();
    await page.getByRole('button', { name: '인쇄 미리보기', exact: true }).click();
    await verify(page.locator('.print-layout'), true);
  });
}
