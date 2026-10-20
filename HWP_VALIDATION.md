# HWP/HWPX 읽기 검증

2026-10-20 · Word Studio 1.3.0 배포 검증 기록. Mac ZIP·Windows Setup·Linux DEB에 HWP 읽기와 페이지 분할 수정이 포함됩니다.

## 구현과 사용

파일 → 파일 열기에서 HWP 5.x 또는 HWPX를 선택합니다. 읽기 전용 화면에서 쪽 이동·확대·찾기·텍스트 복사·PDF 저장·인쇄를 사용하고, 닫으면 기존 Word 문서로 돌아옵니다. HWP 원본과 Word 편집 이력은 바꾸지 않습니다.

MIT 라이선스의 RHWP 0.8.7을 별도 Worker에서 실행합니다. 공개 1.3.0은 페이지 분할 패치를 적용한 `0.8.7-wordstudio.1`, 현재 1.3.1 소스는 HWPX manifest 경로 수정까지 포함한 `0.8.7-wordstudio.2`입니다. 정확한 소스 커밋·패치·빌드 방법·배포 바이너리 해시는 [vendor/rhwp-core](vendor/rhwp-core/README.md)에 보존합니다. WebAssembly·한국어 Noto 대체 글꼴을 앱에 포함하며 네트워크에 문서를 보내지 않습니다. SVG의 실행 코드·외부 리소스를 제거하고 쪽별 clip ID를 분리합니다. 45초 제한과 취소 시 Worker 종료를 적용합니다. 파일 40 MiB·1,000쪽, HWPX 내부 10,000개·압축 해제 100 MiB, 출력 데이터 100 MiB 제한입니다. HWP 3.x·암호/DRM·HWP 저장·편집은 지원하지 않습니다.

중첩 표·글상자를 빠뜨리던 원본 텍스트 API 대신 실제 화면의 layout run을 복사합니다. 긴 취업규칙의 텍스트 복사는 5,901자에서 현재 130,180자로 보완했습니다. 페이지마다 독립적인 SVG 자원 ID를 사용해 다중 쪽 PDF의 표 잘림 영역 충돌도 방지합니다.

## 1.3.1 소스: HWPX manifest 상대 경로 수정 (배포 전)

`META-INF/container.xml` → `Contents/content.hpf` → `href="section0.xml"` → `Contents/section0.xml` 구조를 1.3.0 엔진에 넣으면 보고된 필수 파일 누락 오류가 재현됩니다. ZIP 루트에 사본을 추가하면 열리며, 루트 사본의 내용이 다르면 그 잘못된 본문을 선택하는 것도 확인했습니다.

manifest의 section·바탕쪽·내장 바이너리 경로를 OPF 폴더 기준으로 한 번 해석한 뒤 파서와 지연 로더에 전달합니다. 상대 경로가 실제로 존재하면 우선하며, 없으면 기존 한컴 파일의 `Contents/section0.xml`·`BinData/image1.bmp`와 같은 ZIP 루트 표기를 지원합니다. `.`·`..`를 ZIP 내부에서 정리하고 아카이브 밖 경로·절대 경로·원격 URL은 거부합니다. 외부 링크 항목은 다시 쓰지 않습니다. `content.hpf`·`header.xml`의 위치 변경이나 조판 규칙 변경은 이번 수정 범위에 포함하지 않습니다.

두 문단만 있는 합성 A4 HWPX에는 ZIP 루트 section 사본이 없습니다. 회귀 검사는 열기, 루트 표기 대조본과 렌더링 동일성, 동명 루트 파일보다 정상 본문 우선, 하위 폴더와 spine 순서, 필수 파일 누락과 잘못된 경로를 확인합니다. fixture 생성 출처는 [fixtures 설명](tests/fixtures/hwp/README.md)에 있습니다.

실제 공개 HWPX 원본 3개(`07-image`, `09-diagonal`, `11-table-text`)의 `version.xml`에는 `Hancom Office Hangul` 13.0.0.3622 / 12.0.0.535가 기록되어 있습니다. 원본들은 ZIP 루트 표기를 사용합니다. 이 원본 대조군과 **manifest 경로만 상대 표기로 바꾼 파생본**을 구분하여 검사합니다. 후자는 실제 한컴 앱에서 새로 저장한 파일이라는 의미가 아닙니다. [한컴의 구조 설명](https://tech.hancom.com/hwpxformat/)도 본문과 패키지 파일이 `Contents/`에 있음을 설명합니다.

1.3.1 로컬 검증 결과:

- 새 경로 회귀 5개는 1.3.0 엔진에서 모두 실패하고 수정 엔진에서 모두 통과합니다. 바탕쪽 본문 표시와 지연 로딩한 내장 PNG의 원본 바이트 보존도 검사합니다. 전체 Node 단위 검사 11개와 브라우저 기능 검사 90개가 통과했습니다.
- 지원 공개 원본 11개의 **227쪽 전체 SVG 해시·쪽수·추출 텍스트가 1.3.0과 동일**합니다. HWPX 상대 표기 파생본 3개는 기존 오류가 사라지고 각각 5·1·1쪽이며, 대응 원본과 전체 SVG·텍스트가 같습니다. 암호/구형 샘플의 거부도 유지됩니다. 이는 기존 결과의 보존 검사이며 원본 뷰어와 완전한 조판 일치를 뜻하지 않습니다.
- 최적화 WASM을 새로 빌드했고, WASM 대상 Rust Clippy(`-D warnings`)와 수정 파일 포맷 검사를 통과했습니다. WASM SHA-256은 `20426b9f42bd63059966da66bbdcaedfb626a432af53b846903c570d2a94859d`입니다. upstream 전체 Rust 회귀·Native Rust/Skia 검증은 실행하지 않았습니다.
- Mac 1.3.1 패키지에 같은 WASM이 들어 있음을 ASAR에서 확인하고 로컬 서명을 검증했습니다. 격리 프로필에서 8개 파일의 열기·전체 텍스트 복사·PDF 저장·기존 Word 문서 보존이 통과했습니다. 새 두 문단 HWPX는 화면·PDF 모두 1쪽, 긴 취업규칙은 모두 163쪽이며 렌더러 오류는 0개입니다. 사용자 앱의 문서나 공개 1.3.0 다운로드는 바꾸지 않았습니다.

원시 로그·전후 원본/파생본 결과·비교 결과는 로컬 `.cache/hwpx-manifest-fix/`, 실제 앱 화면·PDF·JSON은 `.cache/hwp-native-MTbwaC/`에 있습니다. `node --test tests/hwpx-manifest.test.cjs`로 새 회귀를 실행하며, `RHWP_CORE_DIRECTORY=/path/to/1.3.0-core`를 지정하면 같은 검사를 이전 엔진에 적용할 수 있습니다.

## 163/154쪽 차이 수정

2025 표준취업규칙은 **154쪽에서 163쪽으로 수정**하여 폴라리스의 총 쪽수와 일치합니다. 단순한 확대/축소나 문서별 쪽수 보정은 사용하지 않습니다.

- 가로/높이가 짧은 용지에서 표 행을 남은 높이보다 36~320px 더 허용하던 엔진 분기를 막았습니다. 다음 페이지의 행이 현재 페이지에 합쳐져 종이 밖에서 잘리던 문제입니다. 이 수정만 적용하면 162쪽이 됩니다.
- 셀 첫 줄에 포함된 원본 문단 앞 여백을 셀의 비정상 시작 좌표로 오인하던 조건을 수정했습니다. 저장된 정상 여백을 뺀 좌표로 기존 프레임 경계를 판단하면 163쪽이 됩니다. 수정된 문단이나 임의 좌표에 보정을 적용하지 않습니다.
- 동일 원본의 모든 쪽을 검사했을 때 텍스트가 물리적 종이 높이를 벗어나던 페이지는 **17개 → 0개**입니다. 생성 쪽번호·공백·반복 표 머리를 제외하고 원문 위치별 문자 집합 90,934개와 SHA-256 `f674b4c36598a8bc0e21274b5727fa85f0191db93d0da3bc6884b8a85b4f8279`가 전후 동일합니다. 이는 수정 전 표시 API가 내보낸 원문 문자 보존 검사이며, 원본 문서 전체의 완벽한 해석을 증명하는 수치는 아닙니다.
- 24개 교대 높이 행으로 만든 합성 HWP/HWPX는 기존 엔진에서 3쪽·15/16행 잘림, 수정본에서 4쪽·모든 48개 셀 한 번씩 표시·복사됩니다. 실제 앱의 PDF도 각각 4쪽입니다.
- 폴라리스·앱의 전체 163쪽을 각각 캡처해 같은 종이 비율로 비교 쌍을 생성했습니다. 앞부분·중간·끝부분과 차이가 큰 페이지를 시각 점검했습니다. 총 쪽수는 같지만 글꼴·문단 간격·일부 페이지 마지막 줄 위치·표 하단선·쪽번호 배치 차이는 남습니다. **모든 페이지 경계가 완전히 일치한다고 주장하지 않습니다.**
- 다른 지원 공개 샘플 10개의 쪽수는 모두 유지했습니다. Mac 실제 패키지에서 긴 원본의 읽기 화면과 저장한 PDF가 모두 163쪽이며, 기존 Word 문서를 보존했습니다.

## 공개 샘플과 비교

파일 출처·정확한 업스트림 커밋·크기·SHA-256은 `tests/hwp-samples.json`에 있습니다. 원본 12개(지원 형식 11개, 거부 확인 1개)와 PDF 참고본 3개를 내려받았습니다. 원본은 Git과 앱 배포물에 포함하지 않습니다.

| 샘플 | 비교 기준 | 기준 쪽수 | Studio 쪽수 | 확인 범위 |
| --- | --- | ---: | ---: | --- |
| 2025 표준근로계약서 HWP | 폴라리스 웹 | 7 | 7 | 전 7쪽 화면 |
| 2025 표준취업규칙 HWP | 폴라리스 웹 | 163 | 163 | 전 쪽 캡처·높이 경계·원문 문자 보존, 주요 경계 시각 점검 |
| 보건소 분장사무 HWP | 폴라리스 웹 / 동봉 한컴 PDF | 4 / 4 | 4 | 전 4쪽 화면 |
| 그림 HWP | 폴라리스 웹 | 5 | 5 | 전 5쪽 화면 |
| 수식 HWP | 폴라리스 웹 | 1 | 1 | 화면 |
| 대각선 표 HWPX | 폴라리스 웹 | 1 | 1 | 화면 |
| 국립국어원 업무계획 HWP | 동봉 한컴 PDF | 35 | 35 | 1·2·10·35쪽 |
| 문단 맞춤 HWP | 동봉 한컴 PDF | 1 | 1 | 화면 |
| 복합 표 HWP | 파싱·표시·복사 검사 | — | 4 | 전 4쪽 화면, 3,365자 복사 |
| 그림 HWPX | 파싱·표시·복사 검사 | — | 5 | 전 5쪽 화면 |
| 표와 텍스트 HWPX | 파싱·표시·복사 검사 | — | 1 | 화면·복사 |
| 암호가 있는 HWP 3.x | 거부 확인 | — | 거부 | 기존 Word 문서 유지 |

폴라리스는 실제 서비스에 같은 공개 파일을 열어 비교했습니다. 한컴 PDF는 업스트림에 같은 이름으로 동봉된 참고본이며 PDF 메타데이터의 `Hwp 2022`·`Hancom PDF`를 확인했습니다. 이 Mac에서 한컴 앱을 직접 실행해 새로 출력한 결과는 아닙니다.

쪽수가 같은 것이 조판 일치를 뜻하지는 않습니다. 설치 글꼴·폭·기준선, 표 행 분할·여백, 복잡한 수식과 도형에는 차이가 남습니다. 정확한 페이지 참조·제출용 출력은 원본 뷰어에서도 확인해야 합니다. 상세 화면 비교와 측정 자료는 재현 시 `analysis/hwp-validation/report.html`에 생성합니다. 이번 163쪽 비교 쌍과 원문 보존 검사는 로컬 `analysis/hwp-pagination/`에 있습니다.

대각선 표 HWPX에서는 빨간 대각선과 색이 일부 누락됩니다. 정제 전 엔진 SVG에도 빨간 stroke가 없고 화면·인쇄 프로필의 결과가 같아, SVG 보안 필터의 손실이 아닌 해석기 측 제한으로 확인했습니다. 가는 표 테두리도 폴라리스보다 희미하며 글꼴 굵기·폭·기준선 차이가 보입니다. 현재 지원은 읽기·내용 확인용이며 원본과 동일한 조판을 보장하지 않습니다.

## 검사와 재현

- Windows·Ubuntu 22.04에서 전체 브라우저 검사 89개와 네이티브 단위 검사 6개가 각각 통과했습니다. [전체 브라우저 검사 기록](https://github.com/beaur8-prog/word-studio/actions/runs/39354510294). HWP 회귀 8개에는 SVG 자원 충돌·저장 실패 창 보호·취소 검사가 포함됩니다. 캐시 없는 개발 서버에서 첫 HWP를 열 때의 모듈 최적화 재시작도 재현 후 수정했습니다.
- Mac arm64 패키지에서 실제 파일 열기 IPC·파일 필터, 오프라인 Worker/WASM·한국어 글꼴, OS 클립보드 복사와 원상 복원, PDF 저장, Word 문서 보존을 확인했습니다. 합성 HWP/HWPX 각각 1쪽, 교대 높이 가로 표 각각 4쪽, 공개 계약서 7쪽, 보건소 표 4쪽, 긴 취업규칙 163쪽은 화면과 실제 PDF 쪽수가 일치합니다. 렌더러 오류는 0개입니다.
- 공개 샘플 11개·화면 37쪽을 검사했고 외부 네트워크 요청·렌더러 오류는 0개입니다. 나머지 암호/구형 1개는 거부합니다.
- Windows setup.exe와 Ubuntu 22.04·24.04의 실제 설치 앱에서도 HWP/HWPX 열기·전체 텍스트 복사·PDF 각각 1쪽·가로 표 4쪽 저장·기존 Word 문서 보존이 통과했습니다. 설치·재설치·제거·사용자 데이터 보존도 확인했습니다. Ubuntu 24.04는 22.04에서 만든 동일한 Debian 패키지를 검사합니다. [CI 실행 기록](https://github.com/beaur8-prog/word-studio/actions/runs/39359083083)의 Windows·linux·linux-24 아티팩트에 결과 JSON·화면·PDF가 있습니다.

```sh
npm ci
npm run build
npx playwright test tests/hwp-reader.spec.ts
npm run test:native
# npm run dev로 로컬 서버를 켠 뒤 공개 샘플 다운로드·검사
npm run test:hwp-samples
# 긴 원본의 전체 쪽수·물리 높이·원문 위치별 문자 검사
node scripts/audit-hwp-pagination.cjs
# 원본 NPM 0.8.7 압축 해제본을 지정하면 변경 전후 문자 보존도 비교
node scripts/audit-hwp-pagination.cjs /path/to/original-rhwp-core
# 별도 프로필로 실제 Electron/패키지 검사
npm run test:hwp-native
# Mac 실제 배포 앱 검사
WORD_STUDIO_EXECUTABLE='/path/to/Word Studio.app/Contents/MacOS/Word Studio' HWP_PUBLIC_SAMPLES=1 npm run test:hwp-native
# 긴 원본도 실제 앱에서 163쪽 PDF로 저장
WORD_STUDIO_EXECUTABLE='/path/to/Word Studio.app/Contents/MacOS/Word Studio' HWP_LONG_SAMPLE=1 npm run test:hwp-native
```

비교 자료는 로컬에만 둡니다. `polaris-results.json`과 `polaris/`의 실제 화면 캡처, `hancom-pdf/`의 48 DPI 참고본 렌더링을 `analysis/hwp-validation/`에 준비한 뒤, Pillow·numpy·reportlab·fonttools가 있는 Python으로 `scripts/compare-hwp-layout.py`를 실행하면 29쌍의 PNG·HTML·JSON과 `output/pdf/HWP-Layout-Comparison.pdf`를 만듭니다. 보고서용 Noto Sans KR 글꼴은 고정 커밋에서 받고 SHA-256을 검사합니다. 종이 영역을 잘라 방향·비율을 유지해 비교하며, 근사 먹물 영역 측정값을 호환성 정확도 점수로 해석하지 않습니다.

## 참고와 라이선스

- [한컴 HWP/OWPML 공개 형식](https://download.hancom.com/support/downloadCenter/hwpOwpml)
- [고용노동부 2025 표준근로계약서·취업규칙 원본](https://www.moel.go.kr/info/etc/dataroom/view.do?bbs_seq=20250300356)
- [폴라리스 웹 HWP 뷰어](https://hwp.polarisoffice.com/)
- [RHWP 소스·지원 범위](https://github.com/edwardkim/rhwp) · [MIT 라이선스](electron/licenses/RHWP-MIT.txt)
- [Noto Sans KR 라이선스](electron/licenses/Noto-Sans-KR-OFL.txt) · [Noto Serif KR 라이선스](electron/licenses/Noto-Serif-KR-OFL.txt)
