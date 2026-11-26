import { expect, test } from '@playwright/test';
import * as CFB from 'cfb';
import { strFromU8, unzipSync } from 'fflate';
import { createHash } from 'node:crypto';
import { inflateRawSync } from 'node:zlib';

const sha = (data: Uint8Array) => createHash('sha256').update(data).digest('hex');

for (const format of ['hwp', 'hwpx'] as const) {
  test(`${format} keeps oversized BMP and PNG payloads in the same editable table`, async ({ page }) => {
    await page.goto('/');
    const result = await page.evaluate(async format => {
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { defaultPage } = await import('/src/types.ts');
      const width = 800, height = 700, stride = width * 3;
      const bmp = new Uint8Array(54 + stride * height), header = new DataView(bmp.buffer);
      header.setUint16(0, 0x4d42, true);
      header.setUint32(2, bmp.length, true);
      header.setUint32(10, 54, true);
      header.setUint32(14, 40, true);
      header.setInt32(18, width, true);
      header.setInt32(22, height, true);
      header.setUint16(26, 1, true);
      header.setUint16(28, 24, true);
      header.setUint32(34, stride * height, true);
      bmp.set([27, 113, 241], 54);

      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      const context = canvas.getContext('2d')!;
      const pixels = context.createImageData(width, height);
      let seed = 0x7cb81a5f;
      for (let index = 0; index < pixels.data.length; index += 4) {
        seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5;
        for (let channel = 0; channel < 3; channel++) pixels.data[index + channel] = (seed >>> (channel * 8)) & 0xff;
        pixels.data[index + 3] = 255;
      }
      context.putImageData(pixels, 0, 0);
      const encodedPng = await new Promise<Blob>((resolve, reject) => canvas.toBlob(
        blob => blob ? resolve(blob) : reject(new Error('Could not create the synthetic PNG')), 'image/png'));
      const png = new Uint8Array(await encodedPng.arrayBuffer());
      const bufferBase64 = (bytes: Uint8Array) => {
        const chunks: string[] = [];
        for (let offset = 0; offset < bytes.length; offset += 8192)
          chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + 8192)));
        return btoa(chunks.join(''));
      };
      const images = [bufferBase64(bmp), bufferBase64(png)];
      const digest = async (bytes: Uint8Array) => [...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))]
        .map(byte => byte.toString(16).padStart(2, '0')).join('');
      const expected = await Promise.all([digest(bmp), digest(png)]);
      const markup = `<table><tbody><tr><td><p>Picture one<img src="data:image/bmp;base64,${images[0]}" width="180" height="150"/></p>
        <p>Picture two<img src="data:image/png;base64,${images[1]}" width="160" height="120"/></p></td><td><p>Editable next cell</p></td></tr></tbody></table>`;
      const writer = new HwpReaderSession();
      const reader = new HwpReaderSession();
      try {
        const exported = await writer.export(markup, defaultPage, format);
        const opened = await reader.open(new File([new Uint8Array(exported.bytes)], 'picture-table.' + format));
        const imported = await reader.editable();
        const doc = new DOMParser().parseFromString(imported.html, 'text/html');
        return { expected, payloadLengths: images.map(source => source.length), pages: opened.count,
          cells: [...doc.querySelectorAll('table td')].map(cell => cell.textContent!.replace(/\s+/g, ' ').trim()),
          pictureSizes: [...doc.querySelectorAll('table img')].map(img => [Number(img.getAttribute('width')), Number(img.getAttribute('height'))]),
          bytes: bufferBase64(exported.bytes) };
      } finally { reader.close(); writer.close(); }
    }, format);

    expect(result.payloadLengths.every(length => length > 2_000_000)).toBe(true);
    expect(result.pages).toBe(1);
    expect(result.cells).toEqual(['Picture one Picture two', 'Editable next cell']);
    expect(result.pictureSizes).toEqual([[180, 150], [160, 120]]);
    const bytes = Buffer.from(result.bytes, 'base64');
    if (format === 'hwpx') {
      const files = unzipSync(bytes), xml = strFromU8(files['Contents/section0.xml']);
      const ids = [...xml.matchAll(/binaryItemIDRef="(image\d+)"/g)].map(match => match[1]);
      expect(ids).toHaveLength(2);
      expect([sha(files[`BinData/${ids[0]}.bmp`]), sha(files[`BinData/${ids[1]}.png`])]).toEqual(result.expected);
      expect(xml.match(/<hp:imgDim dimwidth="60000" dimheight="52500"\/>/g)).toHaveLength(2);
    } else {
      const compound = CFB.read(bytes, { type: 'buffer' });
      const parts = compound.FullPaths.flatMap((name, index) => {
        if (!/^Root Entry\/BinData\/BIN\d+\.(bmp|png)$/i.test(name)) return [];
        const data = Buffer.from(compound.FileIndex[index].content);
        return [[name.slice(-3).toLowerCase(), sha(inflateRawSync(data))]];
      });
      expect(Object.fromEntries(parts)).toEqual({ bmp: result.expected[0], png: result.expected[1] });
    }
  });

  test(`${format} paints a standalone oversized inline BMP after reopening`, async ({ page }) => {
    await page.goto('/');
    const result = await page.evaluate(async format => {
      const { exportEditableHwp } = await import('/src/hwp/editing.ts');
      const { parseDocumentHtml } = await import('/src/extensions.ts');
      const { HwpReaderSession } = await import('/src/hwp/reader.ts');
      const { defaultPage, pageMargins } = await import('/src/types.ts');
      const width = 800, height = 700, stride = width * 3;
      const bmp = new Uint8Array(54 + stride * height);
      const header = new DataView(bmp.buffer);
      header.setUint16(0, 0x4d42, true);
      header.setUint32(2, bmp.length, true);
      header.setUint32(10, 54, true);
      header.setUint32(14, 40, true);
      header.setInt32(18, width, true);
      header.setInt32(22, height, true);
      header.setUint16(26, 1, true);
      header.setUint16(28, 24, true);
      header.setUint32(34, stride * height, true);
      const chunks: string[] = [];
      for (let offset = 0; offset < bmp.length; offset += 8192)
        chunks.push(String.fromCharCode(...bmp.subarray(offset, offset + 8192)));
      const encoded = btoa(chunks.join(''));
      const content = parseDocumentHtml(`<p>Before picture</p><p style="line-height: 140%; margin-left: 36px; text-indent: -20px"><img data-inline-image="true" src="data:image/bmp;base64,${encoded}" width="180" height="150"></p><p>After picture</p>`);
      const doc = { format:'word-studio' as const, version:1 as const, id: crypto.randomUUID(), title:'Picture line',
        content, page:defaultPage, comments:[], updatedAt:new Date().toISOString() };
      const { blob } = await exportEditableHwp(doc, format);
      const reader = new HwpReaderSession();
      try {
        const opened = await reader.open(new File([blob], 'standalone-picture.' + format));
        const painted = new DOMParser().parseFromString((await reader.page(0)).svg, 'image/svg+xml');
        const imported = new DOMParser().parseFromString((await reader.editable()).html, 'text/html');
        return {
          pages: opened.count, encodedSize: encoded.length,
          images: [...painted.querySelectorAll('image')].map(image => ['x','width','height'].map(key => Number(image.getAttribute(key)))),
          expectedX: pageMargins(defaultPage).left * 96 / 25.4 + 16,
          inline: imported.querySelectorAll('p img[data-inline-image]:not([data-image-floating])').length,
          text: [...imported.querySelectorAll('p')].map(paragraph => paragraph.textContent?.trim()).filter(Boolean),
        };
      } finally { reader.close(); }
    }, format);
    expect(result.encodedSize).toBeGreaterThan(2_000_000);
    expect(result.pages).toBe(1);
    expect(result.images).toHaveLength(1);
    const [[x,w,h]] = result.images;
    expect(x).toBeCloseTo(result.expectedX, 0);
    expect([w,h]).toEqual([180, 150]);
    expect(result.inline).toBe(1);
    expect(result.text).toEqual(['Before picture', 'After picture']);
  });
}
