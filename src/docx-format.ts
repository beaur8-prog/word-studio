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
  const defaults = all(styles, 'docDefaults')[0];
  const themeFile = zip.file('word/theme/theme1.xml');
  const theme = themeFile ? new DOMParser().parseFromString(await themeFile.async('string'), 'application/xml') : undefined;
  const settingsFile = zip.file('word/settings.xml');
  const settings = settingsFile ? new DOMParser().parseFromString(await settingsFile.async('string'), 'application/xml') : undefined;
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
      if (color) style.color = hex(val(color)) || '#000000';
      const spacing = child(props, 'spacing');
      if (spacing) style.letterSpacing = `${Number(val(spacing)) / 20}pt`;
      for (const [tag, property, enabled, disabled] of [
        ['b', 'font-weight', 'bold', 'normal'], ['i', 'font-style', 'italic', 'normal'],
      ]) {
        const setting = child(props, tag);
        if (setting) style.setProperty(property, on(setting) ? enabled : disabled);
      }
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
    run(run: Element, script: 'ascii' | 'eastAsia' = 'ascii') { return runCss(runLayers(ancestor(run, 'p')!, run), script); },
    tabs(p: Element) {
      const stops = new Map<number, { position: number; type: string; leader: string }>();
      for (const props of [child(child(defaults, 'pPrDefault'), 'pPr'), ...paragraphStyles(p).map(s => child(s, 'pPr')), child(p, 'pPr')]) {
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
      style.marginTop = '0px'; style.marginBottom = '0px'; style.lineHeight = 'normal';
      for (const props of [child(child(defaults, 'pPrDefault'), 'pPr'), ...paragraphStyles(p).map(s => child(s, 'pPr')), child(p, 'pPr')]) {
        if (!props) continue;
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

export function decorateCell(cell: HTMLElement, source: Element) {
  const props = child(source, 'tcPr');
  const table = ancestor(source, 'tbl')!, tableProps = child(table, 'tblPr');
  const row = source.parentElement!;
  const rows = Array.from(table.children).filter(el => el.localName === 'tr');
  const rowIndex = rows.indexOf(row);
  const cells = Array.from(row.children).filter(el => el.localName === 'tc');
  const grid = Array.from(child(table, 'tblGrid')?.children ?? []).map(col => Number(val(col, 'w')) / 15);
  const span = (c: Element) => Number(val(child(child(c, 'tcPr'), 'gridSpan'))) || 1;
  const offset = cells.slice(0, cells.indexOf(source)).reduce((n, c) => n + span(c), Number(val(child(child(row, 'trPr'), 'gridBefore'))) || 0);
  const widths = grid.slice(offset, offset + span(source));
  if (widths.length === span(source) && widths.every(n => n > 0)) cell.setAttribute('colwidth', widths.map(n => Math.max(1, Math.round(n))).join(','));
  const fill = hex(val(child(props, 'shd'), 'fill'));
  cell.style.backgroundColor = fill || 'transparent';
  const vertical = val(child(props, 'vAlign'));
  cell.style.verticalAlign = vertical === 'center' ? 'middle' : vertical === 'bottom' ? 'bottom' : 'top';
  const borders = child(props, 'tcBorders'), inheritedBorders = child(tableProps, 'tblBorders');
  const margins = child(props, 'tcMar'), inheritedMargins = child(tableProps, 'tblCellMar');
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const edge = side === 'top' ? rowIndex === 0 : side === 'bottom' ? rowIndex + (Number(cell.getAttribute('rowspan')) || 1) >= rows.length : side === 'left' ? offset === 0 : offset + span(source) >= grid.length;
    const border = child(borders, side) ?? child(borders, side === 'left' ? 'start' : side === 'right' ? 'end' : side)
      ?? child(inheritedBorders, edge ? side : ['top', 'bottom'].includes(side) ? 'insideH' : 'insideV');
    const styles: Record<string, string> = { single: 'solid', dashed: 'dashed', dashSmallGap: 'dashed', dotted: 'dotted', double: 'double', nil: 'none', none: 'none' };
    const kind = border ? styles[val(border)] ?? 'solid' : 'none';
    const width = Number(val(border, 'sz')) / 8 || .75;
    cell.style.setProperty(`border-${side}`, `${kind === 'none' ? 0 : width}pt ${kind} ${hex(val(border, 'color')) || '#000000'}`);
    const alternate = side === 'left' ? 'start' : side === 'right' ? 'end' : side;
    const margin = child(margins, side) ?? child(margins, alternate) ?? child(inheritedMargins, side) ?? child(inheritedMargins, alternate);
    const padding = margin ? Number(val(margin, 'w')) / 15 : ['left', 'right'].includes(side) ? 7.2 : 0;
    cell.style.setProperty(`padding-${side}`, `${padding}px`);
  }
}
