'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
const dependencyPaths = [root, path.dirname(root), path.join(path.dirname(root), 'slide-studio')];

// Original paper-and-pencil artwork; no Microsoft artwork or font is embedded.
const ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1024 1024">
  <defs>
    <linearGradient id="blue" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#4995ee"/><stop offset="1" stop-color="#1851b3"/></linearGradient>
    <linearGradient id="paper" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#ffffff"/><stop offset="1" stop-color="#e1edff"/></linearGradient>
  </defs>
  <rect x="62" y="62" width="900" height="900" rx="212" fill="url(#blue)"/>
  <path d="M256 226h348l151 151v384a47 47 0 0 1-47 47H303a47 47 0 0 1-47-47z" fill="#07337b" opacity=".2"/>
  <path d="M237 199h346l151 151v384a47 47 0 0 1-47 47H284a47 47 0 0 1-47-47z" fill="url(#paper)"/>
  <path d="M583 199v114a37 37 0 0 0 37 37h114z" fill="#b4d5ff"/>
  <path d="M326 371h173M326 449h310M326 527h310M326 605h187" stroke="#3b79ce" stroke-width="29" stroke-linecap="round"/>
  <g transform="translate(686 633) rotate(39)">
    <rect x="-40" y="-177" width="80" height="266" rx="20" fill="#15346e"/>
    <rect x="-40" y="-141" width="80" height="37" fill="#8cc6ff"/>
    <path d="M-40 89H40L0 164z" fill="#ffffff"/>
    <path d="M-15 136h30L0 164z" fill="#15346e"/>
  </g>
</svg>`;

async function createWindowsIcon(destination) {
  const sharp = require(require.resolve('sharp', { paths: dependencyPaths }));
  const sizes = [16, 24, 32, 48, 64, 128, 256];
  const images = await Promise.all(sizes.map(size => sharp(Buffer.from(ICON_SVG)).resize(size, size).png().toBuffer()));
  const directory = Buffer.alloc(6 + 16 * sizes.length);
  directory.writeUInt16LE(1, 2);
  directory.writeUInt16LE(sizes.length, 4);
  let offset = directory.length;
  images.forEach((png, index) => {
    const entry = 6 + index * 16;
    directory[entry] = directory[entry + 1] = sizes[index] % 256;
    directory.writeUInt16LE(1, entry + 4);
    directory.writeUInt16LE(32, entry + 6);
    directory.writeUInt32LE(png.length, entry + 8);
    directory.writeUInt32LE(offset, entry + 12);
    offset += png.length;
  });
  await fs.mkdir(destination, { recursive: true });
  const icon = path.join(destination, 'WordStudio.ico');
  await fs.writeFile(icon, Buffer.concat([directory, ...images]));
  await fs.writeFile(path.join(destination, 'WordStudio.svg'), ICON_SVG);
  return icon;
}

async function createMacIcon(destination) {
  const sharp = require(require.resolve('sharp', { paths: dependencyPaths }));
  await fs.mkdir(destination, { recursive: true });
  const iconset = await fs.mkdtemp(path.join(destination, '.WordStudio-'));
  const directory = iconset + '.iconset';
  await fs.rename(iconset, directory);
  try {
    for (const size of [16, 32, 128, 256, 512]) {
      for (const scale of [1, 2]) {
        await sharp(Buffer.from(ICON_SVG)).resize(size * scale, size * scale).png()
          .toFile(path.join(directory, `icon_${size}x${size}${scale === 2 ? '@2x' : ''}.png`));
      }
    }
    await fs.copyFile(path.join(directory, 'icon_512x512@2x.png'), path.join(destination, 'WordStudio.png'));
    const icon = path.join(destination, 'WordStudio.icns');
    execFileSync('/usr/bin/iconutil', ['-c', 'icns', directory, '-o', icon]);
    return icon;
  } finally { await fs.rm(directory, { recursive: true, force: true }); }
}

async function createLinuxIcon(destination) {
  const sharp = require(require.resolve('sharp', { paths: dependencyPaths }));
  await fs.mkdir(destination, { recursive: true });
  const icon = path.join(destination, '512x512.png');
  await sharp(Buffer.from(ICON_SVG)).resize(512, 512).png().toFile(icon);
  return icon;
}

module.exports = { createWindowsIcon, createMacIcon, createLinuxIcon };
if (require.main === module) createWindowsIcon(path.resolve(process.argv[2] || path.join(root, 'electron')))
  .then(filename => console.log(filename))
  .catch(error => { console.error(error.message); process.exitCode = 1; });
