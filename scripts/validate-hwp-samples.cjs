'use strict';
// Public corpus only. Inputs and rendered evidence stay outside Git.
const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const { chromium, expect } = require('@playwright/test');
const root = path.resolve(__dirname, '..');
const directory = path.join(root, 'analysis/hwp-validation');

async function main() {
  const manifest = JSON.parse(await fs.readFile(path.join(root, 'tests/hwp-samples.json'), 'utf8'));
  await fs.mkdir(path.join(directory, 'samples'), { recursive: true });
  await fs.mkdir(path.join(directory, 'studio'), { recursive: true });
  for (const sample of manifest) {
    const target = path.join(directory, 'samples', sample.file);
    let bytes = await fs.readFile(target).catch(() => null);
    if (!bytes) {
      const response = await fetch(sample.url, { signal: AbortSignal.timeout(60000) });
      if (!response.ok) throw Error('Sample download failed: ' + sample.file);
      bytes = Buffer.from(await response.arrayBuffer());
      await fs.writeFile(target, bytes);
    }
    if (crypto.createHash('sha256').update(bytes).digest('hex') !== sample.sha256) throw Error('Public sample digest mismatch: ' + sample.file);
  }
  const browser = await chromium.launch({ channel: process.env.PLAYWRIGHT_CHANNEL || 'chrome', headless: true });
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  const errors = [], network = [];
  page.on('pageerror', error => errors.push(error.message));
  page.on('request', request => { if (/^https?:/.test(request.url()) && !request.url().startsWith('http://127.0.0.1:5191/')) network.push(request.url()); });
  await page.goto('http://127.0.0.1:5191');
  await page.locator('.tiptap').fill('HWP VALIDATION — PRESERVE THIS WORD DOCUMENT');
  const results = [];
  try {
    for (const sample of manifest.filter(sample => /\.(hwp|hwpx)$/.test(sample.file))) {
      await page.locator('input[type=file]').first().setInputFiles(path.join(directory, 'samples', sample.file));
      const reader = page.getByRole('dialog', { name: 'HWP 읽기', exact: true });
      await expect(reader).toBeVisible();
      await expect(reader.locator('.hwp-loading')).not.toBeVisible({ timeout: 60000 });
      if (sample.file.startsWith('12-')) {
        await expect(reader.getByRole('alert')).toBeVisible();
        results.push({ file: sample.file, rejected: true, reason: await reader.getByRole('alert').textContent() });
      } else {
        await expect(reader.getByTestId('hwp-page')).toBeVisible();
        const count = Number((await reader.getByTestId('hwp-page-count').textContent()).replace(/\D/g, ''));
        await reader.getByLabel('HWP 확대').selectOption('50');
        const pages = count <= 7 ? Array.from({ length: count }, (_, i) => i + 1) : [...new Set([1, 2, Math.min(10, count), count])];
        const rendered = [];
        for (const number of pages) {
          await reader.getByLabel('HWP 쪽 번호').fill(String(number));
          await expect(reader.getByTestId('hwp-page')).toHaveAttribute('aria-label', number + '쪽');
          await expect(reader.locator('.hwp-loading')).not.toBeVisible();
          await page.evaluate(() => document.fonts.ready);
          const name = sample.file.replace(/\.(hwp|hwpx)$/, '') + '-page-' + number + '.png';
          await reader.getByTestId('hwp-page').screenshot({ path: path.join(directory, 'studio', name) });
          const geometry = await reader.getByTestId('hwp-page').evaluate(el => {
            const svg = el.querySelector('svg');
            return { width: Number(svg.getAttribute('width')), height: Number(svg.getAttribute('height')),
              textElements: svg.querySelectorAll('text').length, images: svg.querySelectorAll('image').length,
              paintElements: svg.querySelectorAll('path,line,rect,ellipse,polygon').length };
          });
          rendered.push({ number, screenshot: 'studio/' + name, ...geometry });
        }
        await reader.getByRole('button', { name: '텍스트 복사', exact: true }).click();
        await expect(reader.locator('.hwp-status')).toContainText('텍스트를 복사했습니다', { timeout: 60000 });
        const copied = await page.evaluate(() => navigator.clipboard.readText());
        if (sample.file.startsWith('02-') && copied.length < 100000) throw Error('Nested table/textbox content missing from text copy');
        if (sample.file.startsWith('05-') && copied.length < 2000) throw Error('Complex table content missing from text copy');
        results.push({ file: sample.file, pageCount: count, copiedCharacters: copied.length, rendered });
        console.log(sample.file + ': ' + count + ' pages, ' + rendered.length + ' screenshots');
      }
      await reader.getByRole('button', { name: 'HWP 읽기 닫기' }).click();
      await expect(page.locator('.tiptap')).toHaveText('HWP VALIDATION — PRESERVE THIS WORD DOCUMENT');
    }
    const report = { engine: '@rhwp/core 0.8.7', browser: await browser.version(), results, rendererErrors: errors, externalRequests: network };
    await fs.writeFile(path.join(directory, 'studio-results.json'), JSON.stringify(report, null, 2) + '\n');
    if (errors.length || network.length) throw Error('Renderer error or unexpected external request');
  } finally { await browser.close(); }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
