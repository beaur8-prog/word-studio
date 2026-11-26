import { characterScaleProperty, parseCharacterScale } from './character-scale';
import { horizontalInkFits, lineOpportunities, sourceLineBoundaries } from './hwp-source-cell-lines';
import { hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } from './hwp-text-advances';
import { paragraphLineOffsets } from './paragraph-fragments';
import { textBoxFrame } from './textbox-frame';
import type { HwpFloatingTextBox } from './hwp-floating-textbox';

const precision = .05 + 1 / 15 + 2 / 64;
const storySelector = ':scope > [data-textbox-content]';
export const hwpFloatingBoxFontFallback = 'Noto Sans KR';
export const hwpFloatingBoxProofSelector = 'p[data-hwp-source-wrap*="floatingTextBox"]';

type Slot = {
  story: HTMLParagraphElement;
  owner: HTMLElement;
  frame: ReturnType<typeof textBoxFrame>;
  proof: HwpFloatingTextBox;
  nodes: Node[];
  cuts: NonNullable<ReturnType<typeof sourceLineBoundaries>>;
  nested?: HTMLElement;
};

const near = (a: number, b: number) => Math.abs(a - b) <= precision;
const firstFace = (css: CSSStyleDeclaration) => css.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');

function readSlot(block: Element, outer: HTMLElement, index: number): Slot | undefined {
  const nested = block instanceof HTMLElement && block.hasAttribute('data-document-textbox') ? block : undefined;
  const owner = nested?.querySelector<HTMLElement>(storySelector) ?? outer;
  const p = nested ? owner.firstElementChild : block;
  if (!(p instanceof HTMLParagraphElement) || nested && owner.children.length !== 1 ||
      p.querySelector(':not(span,strong,b,em,i),[data-docx-bookmark],[data-document-field],[id]')) return;
  const source = parseHwpSourceWrap(p.dataset.hwpSourceWrap), proof = source?.floatingTextBox;
  const frame = textBoxFrame((nested ?? outer.parentElement)?.dataset.documentTextbox);
  const text = p.textContent;
  if (!proof || !source || !frame || frame.shape !== 'rect' && (!nested || frame.shape !== 'roundRect') || proof.order !== index || proof.count !== outer.children.length ||
      !near(frame.width, proof.width) || Math.abs(frame.height - proof.height) > precision ||
      !!proof.child !== Boolean(nested) || nested && (frame.position !== 'inline' || !frame.fill || frame.border) ||
      hwpWrapFingerprint(text) !== source.fingerprint || normalizeHwpWrapText(text).length !== source.ends.at(-1)) return;
  const css = getComputedStyle(p);
  if (css.textAlign !== proof.alignment || !near(parseFloat(css.textIndent), -proof.hanging) ||
      !near(parseFloat(css.lineHeight), proof.pitch) || css.transform !== 'none') return;
  const walk = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
  while (walk.nextNode()) {
    if (!walk.currentNode.textContent?.trim()) continue;
    const style = getComputedStyle(walk.currentNode.parentElement!);
    const mark = walk.currentNode.parentElement?.closest<HTMLElement>('[style*="--word-character-scale"]');
    if (!proof.faces.includes(firstFace(style)) || Math.abs(parseFloat(style.fontSize) - proof.size) > .07 ||
        style.verticalAlign !== 'baseline' || style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb' ||
        style.textShadow !== 'none' || style.textDecorationLine !== 'none' ||
        (parseCharacterScale(mark?.style.getPropertyValue(characterScaleProperty)) ?? 100) !== proof.scale) return;
  }
  const cuts = sourceLineBoundaries(p, source.ends.slice(0, -1));
  return cuts ? { story: p, owner, frame, proof, nodes: [...p.childNodes], cuts, nested } : undefined;
}

function baseline(line: HTMLElement) {
  const marker = document.createElement('span');
  marker.style.cssText = 'display:inline-block;vertical-align:baseline;width:0;height:0;font-size:0;padding:0;border:0;margin:0';
  line.append(marker);
  try { return marker.getBoundingClientRect().bottom; }
  finally { marker.remove(); }
}

function glyphRows(band: HTMLElement, range: Range) {
  const glyphs: DOMRect[] = [];
  const leaves = document.createTreeWalker(band, NodeFilter.SHOW_TEXT);
  while (leaves.nextNode()) {
    if (!leaves.currentNode.textContent?.trim()) continue;
    range.selectNodeContents(leaves.currentNode);
    glyphs.push(...[...range.getClientRects()].filter(r => r.width > 0 && r.height > 0));
  }
  return glyphs;
}

/** Apply source line endpoints and baselines to the detached print. The
 * ordinary editable drawing, font marks and standard DOCX stay unchanged. */
export function layoutHwpFloatingTextBoxLines(box: HTMLElement) {
  const frame = textBoxFrame(box.dataset.documentTextbox), outer = box.querySelector<HTMLElement>(storySelector);
  if (!frame || frame.position !== 'floating' || frame.vertical !== 'page' || frame.shape !== 'rect' ||
      frame.align !== 'center' || frame.border || frame.fill || !outer || !outer.querySelector(hwpFloatingBoxProofSelector)) return;
  const slots = [...outer.children].map((child, i) => readSlot(child, outer, i));
  if (!slots.every((slot): slot is Slot => !!slot) || slots.filter(slot => slot.nested).length !== 1) return;
  const child = slots.find(slot => slot.nested)!;
  const nested = child.nested!, oldCss = nested.style.cssText, sourceChild = child.proof.child!;
  const range = document.createRange();
  const restore = () => {
    nested.style.cssText = oldCss;
    for (const slot of slots) slot.story.replaceChildren(...slot.nodes);
  };
  const apply = (fallback: boolean): boolean => {
    Object.assign(nested.style, { alignSelf: 'center', marginTop: sourceChild.before + 'px',
      marginBottom: (sourceChild.band - child.proof.height) + 'px', position: 'relative' });
    const edits = slots.map(slot => {
      const { story, proof, cuts } = slot;
      const contents = lineOpportunities(story, cuts, range);
      for (const br of contents.querySelectorAll('wbr')) br.remove();
      const bands = [...contents.children] as HTMLElement[];
      for (const [i, band] of bands.entries()) {
        Object.assign(band.style, { display: 'block', width: `calc(100% / ${proof.scale / 100})`,
          transform: `scaleX(${proof.scale / 100})`, transformOrigin: 'left center', minHeight: proof.pitch + 'px',
          textIndent: i ? '0px' : -proof.hanging + 'px',
          textAlignLast: i === bands.length - 1 || proof.alignment !== 'justify' ? 'auto' : 'justify' });
        band.style.setProperty('line-height', proof.pitch + 'px', 'important');
        for (const run of [band, ...band.querySelectorAll<HTMLElement>('span,strong,b,em,i')]) {
          run.style.whiteSpace = 'pre';
          if (fallback && run.style.fontFamily) run.style.fontFamily += ', "' + hwpFloatingBoxFontFallback + '"';
        }
      }
      story.replaceChildren(contents);
      return { slot, bands };
    });
    const outerBounds = outer.getBoundingClientRect(), initialChild = child.owner.getBoundingClientRect();
    if (!near(outerBounds.width, frame.width) || !near(outerBounds.height, frame.height) ||
        !near(initialChild.left - outerBounds.left, sourceChild.x)) return false;
    const delta = outerBounds.top + sourceChild.y - initialChild.top;
    if (Math.abs(delta) >= child.proof.height / 3) return false;
    nested.style.top = `${delta}px`;
    let previousBottom = -Infinity;
    for (const { slot, bands } of edits) {
      const limit = slot.owner.getBoundingClientRect();
      const { story, proof, cuts } = slot;
      if (!near(limit.height, proof.height) || !near(limit.width, proof.width)) return false;
      for (const [i, band] of bands.entries()) {
        const target = limit.top + proof.baselines[i];
        const shift = target - baseline(band);
        if (Math.abs(shift) >= proof.size) return false;
        band.style.position = 'relative';
        band.style.top = `${shift}px`;
        if (!near(baseline(band), target)) return false;
        const ink = glyphRows(band, range);
        const radius = slot.frame?.cornerRadius ?? 0;
        if (!ink.length || ink.some(glyph => glyph.top < limit.top || glyph.top < previousBottom || glyph.bottom > limit.bottom ||
            radius > 0 && (glyph.top < limit.top + radius || glyph.bottom > limit.bottom - radius) &&
            (glyph.left < limit.left + radius || glyph.right > limit.right - radius))) return false;
        previousBottom = Math.max(...ink.map(glyph => glyph.bottom));
      }
      const ends = paragraphLineOffsets(story);
      if (ends.length !== cuts.length - 2 || ends.some((end, i) => end !== cuts[i + 1].offset) ||
          story.textContent !== slot.nodes.map(node => node.textContent).join('') ||
          !horizontalInkFits(story, range, limit, proof.scale / 100)) return false;
    }
    return true;
  };
  if (apply(false)) { outer.dataset.hwpPrintFloatingTextbox = 'source'; return; }
  restore();
  if (apply(true)) { outer.dataset.hwpPrintFloatingTextbox = 'fallback'; return; }
  restore();
}
