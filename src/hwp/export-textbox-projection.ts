import { textBoxFrame, type TextBoxFrame } from '../textbox-frame';
import { characterScaleProperty, parseCharacterScale } from '../character-scale';
import { hwpWrapFingerprint, normalizeHwpWrapText, parseHwpSourceWrap } from '../hwp-text-advances';
import type { Alignment, HwpExportTextBox, NativeParagraph, NativeRun } from './export-textboxes';

const px = (value: string) => Number.parseFloat(value);
const rgb = /^#[0-9a-f]{6}$/i;
const shapeStory = ':scope > [data-textbox-content]';
export const nativeBoxAnchor = 'data-hwp-native-box-anchor';
const alignments = new Set(['left', 'center', 'right', 'justify']);

function solidTextColor(value: string): string | undefined {
  const match = /^rgb\((\d+), (\d+), (\d+)\)$/.exec(value);
  if (!match) return;
  const [r, g, b] = match.slice(1).map(Number);
  if ([r, g, b].some(channel => channel > 255)) return;
  return '#' + [r, g, b].map(channel => channel.toString(16).padStart(2, '0')).join('');
}

function nativeRun(text: string, style: CSSStyleDeclaration, typography: boolean): NativeRun | undefined {
  const face = style.fontFamily.split(',')[0].trim().replace(/^["']|["']$/g, '');
  const size = px(style.fontSize), color = solidTextColor(style.color);
  if (!face || face.length > 100 || !color || !(size >= 4 && size <= 96) ||
      /[\x00-\x1f\u2028\u2029\ufffc]/u.test(text) ||
      !['400', '700'].includes(style.fontWeight) || !['normal', 'italic'].includes(style.fontStyle) ||
      style.textDecorationLine !== 'none' || style.textShadow !== 'none' || style.verticalAlign !== 'baseline') return;
  const run = { text, face, color, size, bold: style.fontWeight === '700', italic: style.fontStyle === 'italic' };
  if (!typography) return run;
  const authoredScale = style.getPropertyValue(characterScaleProperty).trim();
  const scale = authoredScale ? parseCharacterScale(authoredScale) : 100;
  const tracking = style.letterSpacing === 'normal' ? 0 : Math.round(px(style.letterSpacing) * 100 / size);
  if (!scale || !Number.isInteger(scale) || scale < 70 || scale > 130 || !(tracking >= -50 && tracking <= 50)) return;
  return { ...run, scale, tracking };
}

function sourceParagraph(paragraph: Element, typography = false): NativeParagraph | undefined {
  if (paragraph.tagName !== 'P' || paragraph.querySelector(':not(span,strong,b,em,i),[id],[data-document-field],[data-docx-tab],[data-page-number]')) return;
  const style = getComputedStyle(paragraph);
  const indent = px(style.textIndent), left = px(style.marginLeft) - Math.max(0, -indent), right = px(style.marginRight);
  const before = px(style.marginTop), after = px(style.marginBottom), pitch = px(style.lineHeight);
  if (!alignments.has(style.textAlign) || !Number.isFinite(indent) || Math.abs(indent) > 1000 ||
      ![left, right, before, after].every(value => value >= -.02 && value <= 1000) || !(pitch >= 4 && pitch <= 300) ||
      style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb') return;
  const runs: NativeRun[] = [], walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const leaf = walker.currentNode as Text;
    if (!leaf.length) continue;
    const run = nativeRun(leaf.data, getComputedStyle(leaf.parentElement!), typography);
    if (!run || runs.length >= 128) return;
    runs.push(run);
  }
  if (!runs.length) {
    const run = nativeRun('', style, typography);
    if (!run) return;
    runs.push(run);
  }
  return { runs, align: style.textAlign as Alignment, indent, left: Math.max(0, left), right, before, after, pitch };
}

function plainFrame(frame: TextBoxFrame | undefined): frame is TextBoxFrame {
  return !!frame && frame.align === 'center' && !frame.group && !frame.drawingGroup && !frame.decoration &&
    !frame.gradient && !frame.autoFit && !frame.noWrap && !frame.border && !frame.behind;
}

function projectNestedBox(container: HTMLElement, frame: TextBoxFrame, body: Element): HwpExportTextBox | undefined {
  const anchorText = container.getAttribute(nativeBoxAnchor);
  const [x, y, centered] = anchorText?.split(' ').map(Number) ?? [];
  if (!(x >= 0 && x <= 2000 && y >= 0 && y <= 2000) || ![0, 1].includes(centered) ||
      !plainFrame(frame) || frame.position !== 'floating' || frame.shape !== 'rect' || frame.wrap === 'square' || frame.fill ||
      frame.vertical !== 'page' || body.children.length < 3 || body.children.length > 16 || (body.textContent?.length ?? 0) > 12000) return;
  const marker = () => 'WORDSTUDIO-NATIVE-TEXTBOX-' + crypto.randomUUID();
  const paragraphs: NativeParagraph[] = [];
  let nested: HwpExportTextBox['nested'];
  for (const [slot, block] of [...body.children].entries()) {
    const childFrame = block instanceof HTMLElement && textBoxFrame(block.dataset.documentTextbox);
    const childBody = childFrame && block.querySelector(shapeStory);
    if (childFrame && (!childBody || childBody.children.length !== 1 || block.children.length !== 1 ||
        !plainFrame(childFrame) || childFrame.position !== 'inline' || childFrame.shape !== 'roundRect' ||
        childFrame.cornerRadius === undefined || !childFrame.fill || nested)) return;
    const story = childBody ? childBody.firstElementChild! : block;
    const source = story instanceof HTMLElement ? parseHwpSourceWrap(story.dataset.hwpSourceWrap) : undefined;
    const witness = source && source.floatingTextBox;
    const paragraph = sourceParagraph(story, true);
    const expected = childFrame || frame;
    const literal = story.textContent ?? '', innerWidth = expected.width - expected.padding[1] - expected.padding[3];
    if (!source || !witness?.stored || !paragraph || witness.count !== body.children.length || witness.order !== slot ||
        witness.stored.at(-1)![0] >= literal.length || hwpWrapFingerprint(literal) !== source.fingerprint ||
        Math.abs(expected.width - witness.width) > .12 || Math.abs(expected.height - witness.height) > .12 ||
        paragraph.align !== witness.alignment || Math.abs(paragraph.pitch - witness.pitch) > .12 ||
        Math.abs(paragraph.indent + witness.hanging) > .12 || paragraph.left > .12 || paragraph.right > .12 ||
        witness.stored.some((line, i) => Math.abs(line[5] / 75 - innerWidth) > .25 ||
          Math.abs(line[2] / 75 - witness.size) > .08 || Math.abs((line[2] + line[4]) / 75 - witness.pitch) > .12 ||
          i > 0 && normalizeHwpWrapText(literal.slice(0,line[0])).length !== source.ends[i - 1]) ||
        paragraph.runs.some(run => !witness.faces.includes(run.face) || run.scale !== witness.scale || Math.abs(run.size - witness.size) > .07)) return;
    if (!childFrame || !witness.child) {
      if (childFrame || witness.child) return;
      paragraphs.push({ ...paragraph, stored:witness.stored });
      continue;
    }
    if (!witness.child.host || Math.abs(witness.child.host[5] / 75 -
        (frame.width - frame.padding[1] - frame.padding[3])) > .25) return;
    const { width, height, padding, fill, cornerRadius } = childFrame;
    nested = { slot, child: { marker:marker(), width, height, padding, fill, cornerRadius,
      before:witness.child.before, paragraphs:[{ ...paragraph, stored:witness.stored }] } };
    paragraphs.push({ ...paragraph, align:'center', indent:0, left:0, right:0, before:witness.child.before, after:0,
      stored:[witness.child.host], runs:[{ ...paragraph.runs[0], text:'' }] });
  }
  return nested ? { marker:marker(), width:frame.width, height:frame.height, padding:frame.padding, paragraphs,
    anchor:{x,y,centered:centered === 1}, nested } : undefined;
}

function anchorParagraph(marker: string) {
  const host = document.createElement('p');
  host.textContent = marker;
  return host;
}

/** A printable standalone rectangle becomes one real Hancom text box.
 * Complex stories keep the existing textual fallback and the save reports it. */
export function prepareHwpExportTextBoxes(root: HTMLElement) {
  const boxes: HwpExportTextBox[] = [];
  let fallback = 0;
  for (const container of root.querySelectorAll<HTMLElement>('div[data-document-textbox]')) {
    if (!root.contains(container)) continue;
    const frame = textBoxFrame(container.dataset.documentTextbox);
    const content = container.querySelector(shapeStory);
    if (container.parentElement === root && frame && content && container.children.length === 1 && boxes.length < 100) {
      const nested = projectNestedBox(container, frame, content);
      if (nested?.nested) {
        container.replaceWith(anchorParagraph(nested.marker), anchorParagraph(nested.nested.child.marker));
        boxes.push(nested);
        continue;
      }
    }
    if (container.parentElement !== root || !frame || !content || container.children.length !== 1 ||
        frame.position !== 'inline' || frame.shape !== 'rect' || frame.align !== 'center' ||
        frame.group || frame.drawingGroup || frame.decoration || frame.fill || frame.gradient || frame.autoFit || frame.noWrap ||
        !frame.border || !rgb.test(frame.border) || !(frame.borderWidth! > 0 && frame.borderWidth! <= 3) ||
        frame.width < 100 || frame.width > 2000 || frame.height < 80 || frame.height > 1800 ||
        !content.children.length || content.children.length > 16 || boxes.length >= 100 ||
        content.textContent!.length > 12000) {
      fallback++;
      continue;
    }
    const paragraphs = [...content.children].map(paragraph => sourceParagraph(paragraph));
    if (!paragraphs.every((paragraph): paragraph is NativeParagraph => !!paragraph)) { fallback++; continue; }
    const marker = 'WORDSTUDIO-NATIVE-TEXTBOX-' + crypto.randomUUID();
    container.replaceWith(anchorParagraph(marker));
    boxes.push({ marker, width: frame.width, height: frame.height, padding: frame.padding,
      border: frame.border, stroke: frame.borderWidth!, paragraphs });
  }
  return { boxes, fallback };
}
