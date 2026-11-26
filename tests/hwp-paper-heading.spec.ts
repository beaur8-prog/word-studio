import { expect, test } from '@playwright/test';

test('source paper lettering requires stored and painted percentage pitches', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { sourcePaperHeadingLine, sourceTightInlinePicture } = await import('/src/hwp/paragraph-layout.ts');
    const pageInfo = { marginTop: 30, marginHeader: 15, footerArea: { y: 660 },
      columns: [{ x: 50, width: 230 }, { x: 290, width: 230 }] };
    const stored = [
      { width: 230, verticalPosition: 70, height: 20, textHeight: 20, spacing: 10 },
      { width: 230, verticalPosition: 100, height: 14, textHeight: 14, spacing: 7 },
    ];
    const painted = stored.map(line => ({ secIdx: 0, paraIdx: 0, y: 45 + line.verticalPosition, h: line.textHeight, text: 'x' }));
    const heading = (proof = { stored: true, lines: stored }, runs = painted) =>
      sourcePaperHeadingLine({
        getParagraphLineMetrics: () => JSON.stringify(proof),
        getPageInfo: () => JSON.stringify(pageInfo),
        getPageTextLayout: () => JSON.stringify({ runs }),
      }, 0, 0, 0, ' 큰 글자 작은 글자', 150, 20, 30);
    const pictureProps = { treatAsChar: true, restrictInPage: true, textWrap: 'TopAndBottom', height: 7500 };
    const picture = { x: 55, y: 556, w: 200, h: 100, page: 0 };
    const pictureFits = (position = picture, props = pictureProps, height = 100) =>
      sourceTightInlinePicture({
        getPictureProperties: () => JSON.stringify(props),
        getPageInfo: () => JSON.stringify(pageInfo),
        getParagraphLineMetrics: () => JSON.stringify({ stored: true, controlCount: 1,
          lines: [{ textStart: 0, height, textHeight: 100, spacing: 6 }] }),
      }, 0, 10, 0, position);
    return {
      valid: heading()?.paperHeading?.width,
      recreated: Boolean(heading({ stored: false, lines: stored })),
      unpainted: Boolean(heading(undefined, painted.slice(0, 1))),
      differentPitch: Boolean(heading({ stored: true, lines: [stored[0], { ...stored[1], spacing: 10 }] })),
      picture: pictureFits(), outOfPaper: pictureFits({ ...picture, y: 561 }),
      outOfColumn: pictureFits({ ...picture, x: 115 }),
      caption: pictureFits(picture, { ...pictureProps, hasCaption: true }),
      freeDescent: pictureFits({ ...picture, y: 550 }),
      mismatchedHeight: pictureFits(picture, pictureProps, 101),
    };
  });
  expect(result).toEqual({ valid: 230, recreated: false, unpainted: false, differentPitch: false,
    picture: true, outOfPaper: false, outOfColumn: false, caption: false, freeDescent: false, mismatchedHeight: false });
});

test('paper lettering treats floating decorations as view content and declines changed source geometry', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async () => {
    const { encodeHwpBodyLine } = await import('/src/hwp-body-line.ts');
    const { hwpWrapFingerprint } = await import('/src/hwp-text-advances.ts');
    const { layoutHwpBodyLines } = await import('/src/hwp-body-line-metrics.ts');
    const root = document.createElement('div');
    root.className = 'tiptap';
    root.style.cssText = 'position:fixed;left:-10000px;top:0;width:476px';
    root.innerHTML = '<div data-column-count="2" style="column-count:2;column-gap:16px">'
      + '<p style="margin:0;font:20px/30px Arial"><img class="ProseMirror-separator" width="0" height="0">'
      + '<img data-inline-image="true" width="230" height="40" style="position:absolute;">'
      + '<span style="font-size:20px">Larger words </span><br data-flow-widget="true">'
      + '<span style="font-size:14px">Smaller words</span></p></div>';
    const paragraph = root.querySelector('p')!;
    const picture = paragraph.querySelector<HTMLImageElement>('[data-inline-image]')!;
    picture.dataset.imageFloating = JSON.stringify({
      horizontalPosition: { relative: 'page', offset: 0 }, verticalPosition: { relative: 'page', offset: 0 },
      wrap: { type: 3 }, layoutInCell: false, hwpPaperDrawing: { width: 230, height: 40 },
    });
    paragraph.dataset.hwpBodyLine = encodeHwpBodyLine({ percent: 150, font: 20, base: 30,
      paperHeading: { width: 230, fingerprint: hwpWrapFingerprint(paragraph.textContent!) } });
    document.body.append(root);
    const active = () => { layoutHwpBodyLines(root); return paragraph.dataset.hwpBodyLayout === 'paper-heading'; };
    try {
      const confirmed = active();
      const pitches = [...paragraph.querySelectorAll('[data-hwp-body-run]')].map(run => getComputedStyle(run).lineHeight);
      const ordinaryBreak = paragraph.querySelector('br')!;
      delete ordinaryBreak.dataset.flowWidget;
      const realBreak = active();
      ordinaryBreak.dataset.flowWidget = 'true';
      picture.width += 1;
      const resized = active();
      picture.width -= 1;
      root.style.width = '480px';
      const widened = active();
      root.style.width = '476px';
      paragraph.lastElementChild!.textContent += ' edited';
      const edited = active();
      return { confirmed, pitches, realBreak, resized, widened, edited };
    } finally { root.remove(); }
  });
  expect(result).toEqual({ confirmed: true, pitches: ['30px', '21px'],
    realBreak: false, resized: false, widened: false, edited: false });
});
