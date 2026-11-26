const fs = require('node:fs/promises');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');

let fixture;
module.exports = () => fixture ??= (async () => {
  const engine = path.dirname(require.resolve('@rhwp/core'));
  const { default: init, HwpDocument } = await import(pathToFileURL(engine + '/rhwp.js'));
  await init({ module_or_path: await fs.readFile(engine + '/rhwp_bg.wasm') });
  const author = HwpDocument.createEmpty();
  try {
    author.createBlankDocument();
    author.pasteHtml(0, 0, 0, '<p>Lead</p><p>First bordered line</p><p>Last bordered line</p><p>Plain break</p><p>Disconnected control</p><p>Incomplete begins</p><p>Blocking empty</p><p>Incomplete ends</p>');
    const packageParts = unzipSync(author.exportHwpx());
    const header = strFromU8(packageParts['Contents/header.xml']);
    const fill = header.match(/<hh:borderFill id="1"[^>]*>[\s\S]*?<\/hh:borderFill>/)?.[0];
    const paragraph = header.match(/<hh:paraPr id="0"[^>]*>[\s\S]*?<\/hh:paraPr>/)?.[0];
    if (!fill || !paragraph) throw Error('Missing source defaults');
    const solid = fill.replace('id="1"', 'id="3"')
      .replace(/(<hh:(?:left|right|top|bottom)Border )type="NONE" width="0\.1 mm" color="#000000"/g,
        '$1type="SOLID" width="0.12 mm" color="#007055"');
    const style = (id, connect) => paragraph.replace('id="0"', 'id="' + id + '"')
      .replace(/(<hc:(?:left|right) value=")0"/g, (_, margin) => margin + '1200"')
      .replace('borderFillIDRef="2"', 'borderFillIDRef="3"').replace('connect="0"', 'connect="' + connect + '"');
    const styles = style(20, 1) + style(21, 0);
    packageParts['Contents/header.xml'] = strToU8(header
      .replace('<hh:borderFills itemCnt="2">', '<hh:borderFills itemCnt="3">')
      .replace('</hh:borderFills>', solid + '</hh:borderFills>')
      .replace('<hh:paraProperties itemCnt="20">', '<hh:paraProperties itemCnt="22">')
      .replace('</hh:paraProperties>', styles + '</hh:paraProperties>'));
    const section = strFromU8(packageParts['Contents/section0.xml']);
    packageParts['Contents/section0.xml'] = strToU8(section.replace('<hp:t>Blocking empty</hp:t>', '<hp:t/>').replace(/<hp:p id="(1|2|4|5|6|7)" paraPrIDRef="0"/g,
      (_, i) => '<hp:p id="' + i + '" paraPrIDRef="' + (i === '4' ? 21 : 20) + '"'));
    const hwpx = zipSync(packageParts, { level: 1 }), doc = new HwpDocument(hwpx);
    let hwp;
    try { hwp = doc.exportHwp(); } finally { doc.free(); }
    for (const data of [hwp, hwpx]) {
      const input = new HwpDocument(data);
      try {
        for (const index of [1, 2, 4, 5, 6, 7]) {
          const p = JSON.parse(input.getParaPropertiesAt(0, index));
          if (index === 6 && input.getParagraphLength(0, index) !== 0) throw Error('Failed to author the blank carrier');
          if (p.borderConnect !== (index !== 4) || p.borderLeft.type !== 1 || p.borderLeft.width !== 1 ||
              p.borderLeft.color !== '#007055' || p.marginLeft <= 0 || p.marginRight <= 0)
            throw Error('Fixture did not author a native paragraph border: ' + JSON.stringify(p));
        }
      } finally { input.free(); }
    }
    return { hwp, hwpx };
  } finally { author.free(); }
})();
