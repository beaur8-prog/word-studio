import type JSZip from 'jszip';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const all = (root: Element | Document, tag: string) => Array.from(root.getElementsByTagNameNS(W, tag));
export const val = (el: Element | undefined, key = 'val') => el?.getAttributeNS(W, key) ?? '';
export const child = (root: Element | undefined, name: string) => root && Array.from(root.children).find(el => el.localName === name);
export const hex = (value: string) => /^[\da-f]{6}$/i.test(value) ? `#${value}` : '';
export function ancestor(element: Element, name: string): Element | null {
  let parent = element.parentElement;
  while (parent && parent.localName !== name) parent = parent.parentElement;
  return parent;
}
const on = (el: Element) => !['0', 'false', 'off'].includes(val(el));

/** Resolve property-by-property: direct formatting overrides character/paragraph
 * styles, basedOn ancestors, then document defaults. Paragraph-mark rPr belongs
 * to the empty paragraph/line box; it must never become a table-cell default. */
export async function docxFormat(zip: JSZip) {
  const file = zip.file('word/styles.xml');
  const styles = new DOMParser().parseFromString(file ? await file.async('string') : '<styles/>', 'application/xml');
  const byId = new Map(all(styles, 'style').map(s => [val(s, 'styleId'), s]));
  const normal = all(styles, 'style').find(s => val(s, 'type') === 'paragraph' && val(s, 'default') === '1');
  const normalTable = all(styles, 'style').find(s => val(s, 'type') === 'table' && val(s, 'default') === '1');
  const defaults = all(styles, 'docDefaults')[0];
  const numberingFile = zip.file('word/numbering.xml');
  const numberingXml = numberingFile ? new DOMParser().parseFromString(await numberingFile.async('string'), 'application/xml') : undefined;
  const counters = new Map<string, number[]>();
  const numberCache = new WeakMap<Element, { id: string; level: number; value: number; format: string; pattern: string; label: string } | null>();
  const themeFile = zip.file('word/theme/theme1.xml');
  const theme = themeFile ? new DOMParser().parseFromString(await themeFile.async('string'), 'application/xml') : undefined;
  const settingsFile = zip.file('word/settings.xml');
  const settings = settingsFile ? new DOMParser().parseFromString(await settingsFile.async('string'), 'application/xml') : undefined;
  const balanceWidths = settings && all(settings, 'balanceSingleByteDoubleByteWidth')[0];
  const themeLanguage = settings && val(all(settings, 'themeFontLang')[0], 'eastAsia');
  function themeFont(value: string, language: string) {
    const group = theme?.getElementsByTagNameNS('*', value.startsWith('major') ? 'majorFont' : 'minorFont')[0];
    if (!group || !/^(major|minor)(Ascii|HAnsi|EastAsia|Bidi)$/.test(value)) return '';
    const type = value.endsWith('EastAsia') ? 'ea' : value.endsWith('Bidi') ? 'cs' : 'latin';
    const direct = child(group, type)?.getAttribute('typeface');
    if (direct) return direct;
    const script = /^ko/i.test(language) ? 'Hang' : /^ja/i.test(language) ? 'Jpan'
      : /^zh-(TW|HK|MO|Hant)/i.test(language) ? 'Hant' : /^zh/i.test(language) ? 'Hans' : '';
    return Array.from(group.children).find(el => el.localName === 'font' && el.getAttribute('script') === script)?.getAttribute('typeface') || '';
  }
  function chain(id: string, seen = new Set<string>()): Element[] {
    if (seen.has(id) || seen.size > 30) return [];
    seen.add(id);
    const style = byId.get(id);
    return style ? [...chain(val(child(style, 'basedOn')), seen), style] : [];
  }
  function paragraphStyles(p: Element) {
    return chain(val(child(child(p, 'pPr'), 'pStyle')) || val(normal, 'styleId'));
  }
  function numberingLevel(p: Element, selectedLevel?: number) {
    let id = '', level = 0;
    for (const props of [...paragraphStyles(p).map(s => child(s, 'pPr')), child(p, 'pPr')]) {
      const num = child(props, 'numPr');
      if (child(num, 'numId')) id = val(child(num, 'numId'));
      if (child(num, 'ilvl')) level = Number(val(child(num, 'ilvl'))) || 0;
    }
    if (selectedLevel !== undefined) level = selectedLevel;
    const num = numberingXml && all(numberingXml, 'num').find(n => val(n, 'numId') === id);
    const abstract = numberingXml && all(numberingXml, 'abstractNum').find(n => val(n, 'abstractNumId') === val(child(num || undefined, 'abstractNumId')));
    const override = num && Array.from(num.children).find(n => n.localName === 'lvlOverride' && Number(val(n, 'ilvl')) === level);
    const definition = child(override, 'lvl') || abstract && Array.from(abstract.children).find(n => n.localName === 'lvl' && Number(val(n, 'ilvl')) === level);
    return { id, level, definition, start: Number(val(child(override, 'startOverride')) || val(child(definition, 'start'))) || 1 };
  }
  function numberText(value: number, format: string) {
    if (/^(lower|upper)Letter$/.test(format)) {
      let text = ''; for (let n = value; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text;
      return format.startsWith('lower') ? text.toLowerCase() : text;
    }
    if (/^(lower|upper)Roman$/.test(format)) {
      let text = '', remaining = value;
      for (const [n, label] of [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']] as const)
        while (remaining >= n && text.length < 1000) { text += label; remaining -= n; }
      return format.startsWith('lower') ? text.toLowerCase() : text;
    }
    return format === 'decimalZero' ? String(value).padStart(2, '0') : String(value);
  }
  function runLayers(p: Element, run?: Element) {
    return [child(child(defaults, 'rPrDefault'), 'rPr'), ...paragraphStyles(p).map(s => child(s, 'rPr')),
      ...(run ? chain(val(child(child(run, 'rPr'), 'rStyle'))).map(s => child(s, 'rPr')) : []),
      run ? child(run, 'rPr') : undefined];
  }
  function runCss(layers: (Element | undefined)[], script: 'ascii' | 'eastAsia' = 'ascii') {
    const style = document.createElement('span').style;
    style.fontSize = '11pt'; style.color = '#000000';
    const resolved: Record<string, string> = {};
    let language = themeLanguage || '';
    let underline = '', strike = '', decorationStyle = '';
    for (const props of layers) {
      if (!props) continue;
      language = val(child(props, 'lang'), 'eastAsia') || language;
      const fonts = child(props, 'rFonts');
      for (const key of ['ascii', 'hAnsi', 'eastAsia', 'cs']) {
        const family = themeFont(val(fonts, `${key}Theme`), language) || val(fonts, key);
        if (family) resolved[key] = family;
      }
      const size = Number(val(child(props, 'sz')));
      if (size > 0) style.fontSize = `${size / 2}pt`;
      const color = child(props, 'color');
      if (color) {
        const scheme = theme?.getElementsByTagNameNS('*', 'clrScheme')[0];
        const aliases: Record<string, string> = { dark1: 'dk1', dark2: 'dk2', light1: 'lt1', light2: 'lt2', text1: 'dk1', text2: 'dk2', background1: 'lt1', background2: 'lt2', hyperlink: 'hlink', followedHyperlink: 'folHlink' };
        const key = val(color, 'themeColor');
        const slot = child(scheme, aliases[key] || key)?.firstElementChild;
        let resolved = hex(slot?.getAttribute('val') || slot?.getAttribute('lastClr') || '') || hex(val(color)) || '#000000';
        for (const [attribute, tint] of [['themeTint', true], ['themeShade', false]] as const) {
          const amount = val(color, attribute);
          if (/^[\da-f]{2}$/i.test(amount)) resolved = '#' + [1, 3, 5].map(i => {
            const component = parseInt(resolved.slice(i, i + 2), 16), ratio = parseInt(amount, 16) / 255;
            return Math.round(tint ? component + (255 - component) * (1 - ratio) : component * ratio).toString(16).padStart(2, '0');
          }).join('');
        }
        style.color = resolved;
      }
      const spacing = child(props, 'spacing');
      // Word's legacy fullwidth balancing applies tracking twice to CJK
      // characters. Flatten the compatibility setting into run formatting.
      if (spacing) style.letterSpacing = `${Number(val(spacing)) / 20 *
        (script === 'eastAsia' && balanceWidths && on(balanceWidths) ? 2 : 1)}pt`;
      for (const [tag, property, enabled, disabled] of [
        ['b', 'font-weight', 'bold', 'normal'], ['i', 'font-style', 'italic', 'normal'],
      ]) {
        const setting = child(props, tag);
        if (setting) style.setProperty(property, on(setting) ? enabled : disabled);
      }
      const u = child(props, 'u');
      if (u) { underline = ['none', '0', 'false'].includes(val(u)) ? '' : 'underline'; decorationStyle = val(u) === 'double' ? 'double' : val(u) === 'dotted' ? 'dotted' : /dash/i.test(val(u)) ? 'dashed' : 'solid'; }
      for (const tag of ['strike', 'dstrike']) {
        const setting = child(props, tag);
        if (setting) { strike = on(setting) ? 'line-through' : ''; if (tag === 'dstrike' && on(setting)) decorationStyle = 'double'; }
      }
      if (u || child(props, 'strike') || child(props, 'dstrike')) {
        style.textDecoration = [underline, strike].filter(Boolean).join(' ') || 'none';
        style.textDecorationStyle = decorationStyle || 'solid';
      }
      const vertical = child(props, 'vertAlign');
      if (vertical) style.verticalAlign = val(vertical) === 'superscript' ? 'super' : val(vertical) === 'subscript' ? 'sub' : 'baseline';
      const caps = child(props, 'caps'), smallCaps = child(props, 'smallCaps');
      if (caps) style.textTransform = on(caps) ? 'uppercase' : 'none';
      if (smallCaps) style.fontVariantCaps = on(smallCaps) ? 'small-caps' : 'normal';
      const highlight = val(child(props, 'highlight'));
      if (highlight) style.backgroundColor = highlight === 'none' ? 'transparent' : highlight;
      const shading = hex(val(child(props, 'shd'), 'fill'));
      if (shading) style.backgroundColor = shading;
    }
    const family = resolved[script] || resolved.ascii || resolved.hAnsi || resolved.eastAsia;
    if (family) style.fontFamily = family;
    return style.cssText;
  }
  return {
    numbering(p: Element) {
      if (numberCache.has(p)) return numberCache.get(p);
      const { id, level, definition, start } = numberingLevel(p);
      if (!id || id === '0' || !definition) { numberCache.set(p, null); return null; }
      const values = counters.get(id) || [];
      values[level] = values[level] === undefined ? start : values[level] + 1;
      values.length = level + 1; counters.set(id, values);
      const format = val(child(definition, 'numFmt')) || 'decimal';
      const pattern = val(child(definition, 'lvlText')) || `%${level + 1}.`;
      const info = { id, level, value: values[level], format, pattern,
        label: pattern.replace(/%([1-9])/g, (_, n) => {
          const ancestor = numberingLevel(p, Number(n) - 1);
          return numberText(values[Number(n) - 1] ?? ancestor.start, val(child(ancestor.definition, 'numFmt')) || format);
        }) };
      numberCache.set(p, info); return info;
    },
    cell(cell: HTMLElement, source: Element) {
      const table = ancestor(source, 'tbl')!, props = child(table, 'tblPr');
      const layers = chain(val(child(props, 'tblStyle')) || val(normalTable, 'styleId')).map(s => child(s, 'tblPr'));
      decorateCell(cell, source, layers);
    },
    autoSpacing(p: Element) {
      // Resolve each standard property independently, including explicit false.
      const result = { de: true, dn: true };
      for (const props of [child(child(defaults, 'pPrDefault'), 'pPr'), ...paragraphStyles(p).map(s => child(s, 'pPr')), child(numberingLevel(p).definition, 'pPr'), child(p, 'pPr')]) {
        const de = child(props, 'autoSpaceDE'), dn = child(props, 'autoSpaceDN');
        if (de) result.de = on(de);
        if (dn) result.dn = on(dn);
      }
      return result;
    },
    lineRule(p: Element) {
      let rule = 'auto';
      for (const props of [child(child(defaults, 'pPrDefault'), 'pPr'), ...paragraphStyles(p).map(s => child(s, 'pPr')), child(numberingLevel(p).definition, 'pPr'), child(p, 'pPr')]) {
        const spacing = child(props, 'spacing');
        if (val(spacing, 'line')) rule = val(spacing, 'lineRule') || 'auto';
      }
      return rule;
    },
    run(run: Element, script: 'ascii' | 'eastAsia' = 'ascii') { return runCss(runLayers(ancestor(run, 'p')!, run), script); },
    tabs(p: Element) {
      const stops = new Map<number, { position: number; type: string; leader: string }>();
      for (const props of [child(child(defaults, 'pPrDefault'), 'pPr'), ...paragraphStyles(p).map(s => child(s, 'pPr')), child(numberingLevel(p).definition, 'pPr'), child(p, 'pPr')]) {
        for (const tab of Array.from(child(props, 'tabs')?.children ?? [])) {
          const position = Number(val(tab, 'pos')) / 15, type = val(tab);
          if (type === 'clear') stops.delete(position);
          else if (['left', 'right', 'center', 'decimal'].includes(type) && Number.isFinite(position))
            stops.set(position, { position, type, leader: val(tab, 'leader') });
        }
      }
      return JSON.stringify([...stops.values()]);
    },
    paragraph(p: Element) {
      const style = document.createElement('p').style;
      const empty = !all(p, 't').some(t => t.textContent);
      style.cssText = runCss([...runLayers(p), ...(empty ? [child(child(p, 'pPr'), 'rPr')] : [])]);
      // A non-empty paragraph's line strut follows its text. An unused style
      // font (often a missing legacy font) must not enlarge every line box.
      if (!empty) {
        let largest = 0;
        for (const run of all(p, 'r').filter(r => all(r, 't').some(t => t.textContent))) {
          const actual = document.createElement('span').style;
          actual.cssText = runCss(runLayers(p, run), /[\u1100-\u11ff\u2e80-\ua4cf\uac00-\ud7af]/u.test(run.textContent || '') ? 'eastAsia' : 'ascii');
          const size = parseFloat(actual.fontSize);
          if (size > largest) { largest = size; style.fontSize = actual.fontSize; style.fontFamily = actual.fontFamily; }
        }
      }
      style.marginTop = '0px'; style.marginBottom = '0px'; style.lineHeight = 'normal';
      // OOXML defaults to enabled; an inherited false and a direct override
      // must both survive the same property-by-property style resolution.
      style.widows = '2'; style.orphans = '2';
      for (const props of [child(child(defaults, 'pPrDefault'), 'pPr'), ...paragraphStyles(p).map(s => child(s, 'pPr')), child(numberingLevel(p).definition, 'pPr'), child(p, 'pPr')]) {
        if (!props) continue;
        const wrap = child(props, 'wordWrap');
        if (wrap) style.wordBreak = on(wrap) ? 'normal' : 'keep-all';
        const widow = child(props, 'widowControl');
        if (widow) { style.widows = on(widow) ? '2' : '1'; style.orphans = style.widows; }
        const alignment = val(child(props, 'jc'));
        if (['left', 'right', 'center', 'both', 'distribute'].includes(alignment)) style.textAlign = ['both', 'distribute'].includes(alignment) ? 'justify' : alignment;
        const spacing = child(props, 'spacing');
        for (const [attr, property] of [['before', 'margin-top'], ['after', 'margin-bottom']]) {
          if (val(spacing, attr) !== '') style.setProperty(property, `${Number(val(spacing, attr)) / 15}px`);
        }
        const line = Number(val(spacing, 'line')), rule = val(spacing, 'lineRule');
        if (line > 0) style.lineHeight = !rule || rule === 'auto' ? String(line / 240) : `${line / 20}pt`;
        const indent = child(props, 'ind');
        for (const [attr, alternate, property] of [['left', 'start', 'margin-left'], ['right', 'end', 'margin-right']]) {
          const value = val(indent, alternate) || val(indent, attr);
          if (value !== '') style.setProperty(property, `${Number(value) / 15}px`);
        }
        if (val(indent, 'hanging')) style.textIndent = `${-Number(val(indent, 'hanging')) / 15}px`;
        else if (val(indent, 'firstLine')) style.textIndent = `${Number(val(indent, 'firstLine')) / 15}px`;
      }
      return style.cssText;
    },
  };
}

export function decorateCell(cell: HTMLElement, source: Element, tableLayers: (Element | undefined)[] = []) {
  const props = child(source, 'tcPr');
  const noWrap = child(props, 'noWrap');
  if (noWrap && !['0', 'false', 'off'].includes(val(noWrap))) cell.style.whiteSpace = 'nowrap';
  const table = ancestor(source, 'tbl')!, tableProps = child(table, 'tblPr');
  const inherited = (group: string, side: string) => [...tableLayers, tableProps].map(p => child(child(p, group), side)).filter(Boolean).at(-1);
  const row = source.parentElement!;
  const rows = Array.from(table.children).filter(el => el.localName === 'tr');
  const rowIndex = rows.indexOf(row);
  const cells = Array.from(row.children).filter(el => el.localName === 'tc');
  const grid = Array.from(child(table, 'tblGrid')?.children ?? []).map(col => Number(val(col, 'w')) / 15);
  const span = (c: Element) => Number(val(child(child(c, 'tcPr'), 'gridSpan'))) || 1;
  const offset = cells.slice(0, cells.indexOf(source)).reduce((n, c) => n + span(c), Number(val(child(child(row, 'trPr'), 'gridBefore'))) || 0);
  const widths = grid.slice(offset, offset + span(source));
  if (widths.length === span(source) && widths.every(n => n > 0)) cell.setAttribute('colwidth', widths.map(n => Math.max(1, Math.round(n))).join(','));
  const shading = child(props, 'shd');
  let fill = hex(val(shading, 'fill'));
  const percentage = val(shading).match(/^pct(\d+(?:\.\d+)?)$/);
  if (percentage) {
    const foreground = hex(val(shading, 'color')) || '#000000', background = fill || '#ffffff';
    const ratio = Math.min(100, Number(percentage[1])) / 100;
    fill = '#' + [1, 3, 5].map(i => Math.round(parseInt(foreground.slice(i, i + 2), 16) * ratio +
      parseInt(background.slice(i, i + 2), 16) * (1 - ratio)).toString(16).padStart(2, '0')).join('');
  }
  cell.style.backgroundColor = fill || 'transparent';
  const vertical = val(child(props, 'vAlign'));
  cell.style.verticalAlign = vertical === 'center' ? 'middle' : vertical === 'bottom' ? 'bottom' : 'top';
  const borders = child(props, 'tcBorders');
  const margins = child(props, 'tcMar');
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const edge = side === 'top' ? rowIndex === 0 : side === 'bottom' ? rowIndex + (Number(cell.getAttribute('rowspan')) || 1) >= rows.length : side === 'left' ? offset === 0 : offset + span(source) >= grid.length;
    const border = child(borders, side) ?? child(borders, side === 'left' ? 'start' : side === 'right' ? 'end' : side)
      ?? inherited('tblBorders', edge ? side : ['top', 'bottom'].includes(side) ? 'insideH' : 'insideV');
    const styles: Record<string, string> = { single: 'solid', dashed: 'dashed', dashSmallGap: 'dashed', dotted: 'dotted', double: 'double', nil: 'none', none: 'none' };
    const kind = border ? styles[val(border)] ?? 'solid' : 'none';
    // OOXML double sz describes each stroke; CSS includes both strokes and
    // the intervening gap in its total border width.
    const width = (Number(val(border, 'sz')) / 8 || .75) * (kind === 'double' ? 3 : 1);
    cell.style.setProperty(`border-${side}`, `${kind === 'none' ? 0 : width}pt ${kind} ${hex(val(border, 'color')) || '#000000'}`);
    const alternate = side === 'left' ? 'start' : side === 'right' ? 'end' : side;
    const margin = child(margins, side) ?? child(margins, alternate) ?? inherited('tblCellMar', side) ?? inherited('tblCellMar', alternate);
    const padding = margin ? Number(val(margin, 'w')) / 15 : ['left', 'right'].includes(side) ? 7.2 : 0;
    cell.style.setProperty(`padding-${side}`, `${padding}px`);
  }
}
