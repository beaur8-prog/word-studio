'use strict';
const fs = require('node:fs');
const path = require('node:path');

function buildLithosReleaseNote(source, version) {
  const start = source.indexOf('## 9. 버전별 변경 내역\n');
  const end = source.indexOf('\n## 10. ', start);
  if (start < 0 || end < 0) throw Error('Guide changelog section not found');
  const changes = source.slice(start, end);
  const versionHeading = `### ${version} ·`;
  const current = changes.indexOf(versionHeading);
  if (current < 0) throw Error('Current version changelog not found: ' + version);
  const next = changes.indexOf('\n### ', current + versionHeading.length);
  const currentChanges = changes.slice(current, next < 0 ? undefined : next).trim();
  let note = source.slice(0, start) + '## 9. 이번 버전의 변경 사항\n\n' + currentChanges + '\n' + source.slice(end);
  // Keep current usage instructions while removing references to superseded app versions.
  note = note.replace('**1.4.0부터 HWP/HWPX를', '**HWP/HWPX를')
    .replace('1.3.x는 읽기 전용이며, 1.4.0부터 아래 절차로', '아래 절차로')
    .replace('### HWP/HWPX 편집과 Word 전환 (1.4.0)', '### HWP/HWPX 편집과 Word 전환')
    .replace('1.4.0 앱은 다단의 페이지 흐름', '이 앱은 다단의 페이지 흐름')
    .replace(' 1.4.1에서 확인한 약 **168MB 편집본**의 열기·저장·자동 복구 기록도 보존했습니다.', '');
  note = note.replace(/\]\(([^()\/#?:\s]+\.md)(#[^)]*)?\)/g, (_, name, fragment = '') =>
    `](https://github.com/beaur8-prog/word-studio/blob/v${version}/${name}${fragment})`);
  const assetFolder = `Word-Studio-${version}`;
  const base = `https://github.com/beaur8-prog/word-studio/releases/download/v${version}/`;
  for (const suffix of ['windows-x64-setup.exe', 'mac-arm64.zip', 'linux-x64.deb']) {
    const name = `${assetFolder}-${suffix}`;
    note = note.replaceAll(base + name, assetFolder + '/' + name);
  }
  const extra = `추가 파일: [SHA-256 체크섬](${assetFolder}/SHA256SUMS.txt) · [사용 안내 원본 받기](${assetFolder}/${assetFolder}-User-Guide.md?download=1)\n\n` +
    `[GitHub 릴리즈](https://github.com/beaur8-prog/word-studio/releases/tag/v${version})에서도 같은 파일을 받을 수 있습니다.\n\n`;
  if (!note.includes('Linux 배포본은 Ubuntu')) throw Error('Download section marker missing');
  note = note.replace('Linux 배포본은 Ubuntu', extra + 'Linux 배포본은 Ubuntu');
  note += '\n## 공유 페이지 이용 안내\n\n이 페이지와 첨부 파일은 Mac의 Lithos에서 제공합니다. Mac과 Lithos가 켜져 있고 인터넷에 연결되어 있어야 열립니다. 앱 재시작이나 외부 연결 재연결 후에는 공유 주소가 바뀔 수 있습니다. 이 링크에 접속되지 않을 때는 위의 GitHub 릴리즈를 이용하세요.\n';
  const versions = new Set(note.match(/\b1\.\d+\.(?:\d+|x)\b/g) || []);
  if (versions.size !== 1 || !versions.has(version)) throw Error('Old app version remains in the share page: ' + [...versions].join(', '));
  const attachmentLinks = note.match(/\]\(Word-Studio-[^)]+\)/g) || [];
  if (attachmentLinks.length !== 5 || attachmentLinks.some(link => !link.startsWith(`](${assetFolder}/`))) throw Error('Unexpected share attachments');
  return note;
}

module.exports = { buildLithosReleaseNote };
if (require.main === module) {
  const root = path.resolve(__dirname, '..');
  const version = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8')).version;
  const result = buildLithosReleaseNote(fs.readFileSync(path.join(root, 'USER_GUIDE.md'), 'utf8'), version);
  if (process.argv[2]) fs.writeFileSync(process.argv[2], result);
  else process.stdout.write(result);
}
