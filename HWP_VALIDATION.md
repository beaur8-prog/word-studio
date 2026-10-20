# HWP/HWPX 읽기·편집·변환 검증

2026-10-20 · Word Studio 1.4.0 배포 검증 기록. Mac ZIP·Windows Setup·Linux DEB에 한글 편집·형식 전환과 기존 읽기·페이지 분할·HWPX manifest 경로 수정이 포함됩니다. 1.3.x 원본 보기 검증 이력도 아래에 보존합니다.

## 1.4.1 용량 확장·연속 스크롤 (개발 빌드 · 미배포)

1.4.0의 24 MiB 제한은 원본 HWP/HWPX 파일이 아닌 중간 HTML에 적용됐습니다. 압축 해제와 내장 그림의 Base64 인코딩 때문에 정상적인 21MB 파일도 이 한도를 넘을 수 있었습니다. 초기 수정의 100 MiB보다 한도를 더 늘리고 열기·그림·저장·복구 경로를 함께 맞췄습니다. 설정은 [electron/document-limits.json](electron/document-limits.json)에서 공유합니다.

| 구분 | 1.4.1 개발 빌드 |
| --- | --- |
| HWP/HWPX·DOCX·TXT 원본 및 한글 저장 결과 | 128 MiB |
| 압축을 푼 HWPX/DOCX | 512 MiB · 내부 파일 10,000개 |
| 편집 변환 HTML 및 HTML 문서 | UTF-8 기준 256 MiB |
| 내장 그림 한 개 | 원본 128 MiB에 해당하는 Base64 길이 |
| Word Studio 편집본·자동 복구 한 건 | 384 MiB |
| 한 쪽 SVG·전체 출력 SVG | 각각 192 MiB·256 MiB |
| 뷰어 쪽 수 | 1,000쪽 |

큰 HTML과 한글 저장 결과는 Worker 사이에서 전송 가능한 버퍼로 넘겨 문자열·바이너리의 추가 복사를 줄입니다. 열기에는 120초, 편집/저장 변환에는 180초, 개별 쪽에는 45초 제한을 둡니다. 요청은 순서대로 처리하며 실제 Worker 실행 시점부터 시간을 계산합니다. 닫기·취소는 Worker와 대기 요청을 정리합니다. 변환 크기 초과 안내에는 원본 파일 크기와 다른 기준임을 명시합니다. 일반 파일 저장에도 다시 열 수 있는 형식별 크기 한도를 적용합니다.

64 MiB보다 큰 복구본은 현재·이전 문서를 별도 불변 스냅샷에 저장하고 작은 manifest만 원자적으로 바꿉니다. 두 문서를 하나의 거대한 JSON 문자열로 합쳐 런타임 한계에 걸리지 않도록 했습니다. 기존 인라인 복구본은 계속 읽고 필요한 때에만 이전 형식에서 옮깁니다. 실패 시 기존 manifest·스냅샷을 유지하며, 읽기와 쓰기를 같은 큐로 직렬화합니다. 새 대용량 복구 형식은 구버전 앱에서 인식하지 못하므로 중요한 문서는 Word Studio 파일로도 보관하세요.

원본 뷰어는 휠·트랙패드로 쪽을 이어서 스크롤합니다. Page Up/Down·Home/End, 쪽 번호·이전/다음·찾기가 같은 스크롤 위치와 연동됩니다. 확대 시 읽던 위치를 유지하고 가로 스크롤을 지원합니다. 원본 쪽별 용지 크기로 전체 스크롤 범위를 계산하되 화면 근처의 SVG만 DOM에 넣습니다. 화면 밖 캐시는 8쪽·문자열 64 MiB 이내이며 표시 데이터도 별도 64 MiB 예산을 적용합니다. 한 쪽이 예산보다 클 때는 현재 쪽을 우선합니다. 빠른 이동 시 필요 없어진 대기 렌더링은 취소하고 출력용 요청은 보존합니다. 찾기는 SVG를 만들지 않고 쪽 텍스트만 조회합니다.

실제 PNG 그림을 포함한 약 21MB HWP/HWPX에서 1.4.0의 동일한 24MB 오류를 재현한 기록은 로컬 `.cache/hwp-size-limit/`에 있습니다. 이번에는 원본 HWP **84,843,520바이트**가 편집 HTML **112,180,827바이트**로 늘어나는 12쪽 문서를 편집하고 Word Studio 사본과 HWP로 저장·재열기했습니다. 그림 12개의 SHA-256이 일치하며 렌더러 오류는 0개였습니다. 1,000쪽 합성 문서의 500쪽·끝 쪽 이동과 스크롤·검색, 인접 쪽만 렌더링하는 동작도 확인했습니다.

Mac에서 전체 기능 시험 **102개**와 단위 시험 **21개**를 통과했고, 최종 뷰어 코드의 **12개 회귀 검사**도 통과했습니다. 단위 시험은 기존 복구 형식의 자동 이전·현재/이전 버전 보존·원자적 저장 실패 시 롤백·스냅샷 경로 제한·용량 초과 거부를 포함합니다. 패키징한 1.4.1 앱에서 약 85MB HWP의 그림 12개를 가져오고, **168,270,409바이트** 편집본의 네이티브 열기·저장·자동 복구와 실제 PNG 6개의 데이터 일치를 확인했습니다. 저장 취소·필터·메뉴·HWP/HWPX 재열기도 통과했으며 렌더러 오류는 0개였습니다. 이 수정본의 Windows·Linux 설치 검사는 아직 실행하지 않았습니다.

같은 Mac 패키지에서 원본 보기 8개 파일의 열기·복사·Word 문서 보존·PDF 출력을 확인했습니다. 여러 쪽 문서는 실제 휠 스크롤·끝/처음 이동 후 가까운 SVG만 남는 동작을 검사했습니다. 가로 표는 **4쪽**, 표준취업규칙은 **163쪽**이며 PDF도 같은 쪽 수를 유지했습니다. 근로계약서는 7쪽, 보건소 표는 4쪽입니다. 패키징된 네이티브 앱 두 검사 모두 렌더러 오류는 0개였습니다.

시험 파일은 [tests/large-hwp-fixture.cjs](tests/large-hwp-fixture.cjs)·[tests/hwp-scroll-fixture.cjs](tests/hwp-scroll-fixture.cjs)에서 실행 시 생성하며 Git에 넣지 않습니다. 큰 설치 앱 검사는 `HWP_LARGE_FIXTURE_JSON`에 생성한 `fixture.json` 경로를 지정하고, `HWP_LARGE_STORAGE_COPIES=6`으로 약 168MB 편집본의 열기·저장·복구를 검사할 수 있습니다. 연속 스크롤은 [tests/hwp-reader.spec.ts](tests/hwp-reader.spec.ts), 용량/편집은 [tests/hwp-editing.spec.ts](tests/hwp-editing.spec.ts), 설치 앱은 `npm run test:hwp-native`와 `npm run test:hwp-editing-native`로 검사합니다. 기록은 로컬 `.cache/hwp-scroll-*.log`·`.cache/hwp-scroll-validation/`·`test-results/`에 보관합니다.

이 값은 앱의 허용 상한이며 모든 최대 크기 문서의 처리 성공이나 같은 속도를 보장하지 않습니다. 실제 메모리와 시간은 그림 해상도·개체 수·문서 구조에 따라 달라집니다. 무손실 조판 제한은 아래 1.4.0과 같습니다. 공개 다운로드와 리토스는 1.4.0을 유지하며, 아래 교차 플랫폼 설치 기록도 1.4.0에 대한 것입니다.

## 1.4.0 편집·변환

원본 보기에서 **Word 화면에서 편집 → 편집본 만들기**로 기존 Word 편집기를 사용합니다. 본문·기본 글자/문단 서식·표 병합·그림·용지/여백을 가져오며, 변환 전에 실제 Word 인쇄 쪽 수를 계산해 원본 보기와 비교합니다. 원본은 따로 보관하며 편집 중 **파일 → 한글 원본 보기**로 다시 확인할 수 있습니다. 이전 Word 문서의 복구 저장 실패 시 전환을 취소할 수 있습니다.

내보내기는 같은 엔진의 `pasteHtml`과 `exportHwpWithReport`/`exportHwpxWithReport`를 사용합니다. HWP는 실제 CFB/HWP 5.1 파일, HWPX는 ZIP/XML 파일입니다. 확장자만 바꾸지 않습니다. 직렬화 결과를 다시 파싱하고 저장 전에 Word와 출력 HWP의 쪽 수 및 변환 손실을 표시합니다. 선택 HTML이 표·그림을 빼는 문제를 피하기 위해 본문 컨트롤을 원래 위치에 끼워 넣고, 셀 크기·열 격자·여백·선 굵기도 가져옵니다. 한글과 Word의 본문/머리말 여백 단위 차이, HTML 들여쓰기 줄바꿈, 첫 문단 정렬 손실도 보정합니다. WASM 바이너리는 1.3.1과 같습니다.

큰 표를 하나로 다시 조판하면 쪽 수가 급감하므로 저장에는 Word 인쇄의 쪽 경계를 사용합니다. 이 과정에서 하나의 표·문단이 여러 부분으로 나뉘고 반복 머리행이 각 표에 포함될 수 있습니다. 동일 조판이나 무손실 왕복을 보장하지 않습니다. 떠 있는 개체는 본문 흐름으로 옮깁니다. 수식은 원문 스크립트 표시, 글상자는 가능한 본문 텍스트, 기타 미지원 개체는 위치 표시로 바꿉니다. 검토 메모·자동 필드·첫 쪽/홀짝 머리말·고급 개체 및 일부 표 서식에는 제한이 있으며 내보내기 전에 안내합니다.

### 공개 샘플의 편집·HWP 저장 결과 (Mac, 2026-10-20)

아래는 원본을 읽기 화면으로 확인한 뒤 Word 편집기로 변환하고 HWP 저장 결과를 다시 읽은 값입니다. 기존 1.3.1 원본 보기의 페이지 분할 검증과 구분합니다. **큰 배치 차이가 남아 있는 문서도 있으며 변환 호환성이 완성됐다는 의미가 아닙니다.** 각 문서의 표·그림 수는 편집본에서 확인했습니다.

| 문서 | 원본 보기 | Word 편집/인쇄 | HWP 저장 후 보기 | 편집본 표/그림 |
| --- | ---: | ---: | ---: | --- |
| 근로계약서 | 7 | 7 | 7 | 표 10 |
| 2025 표준취업규칙 | 163 | 177 | 181 | 표 45 |
| 보건소 분장사무 | 4 | 7 | 7 | 표 1 |
| 국어발전계획 | 35 | 39 | 50 | 표 55 · 그림 3 |
| 복합 표 | 4 | 4 | 4 | 표 4 |
| 그림 HWP | 5 | 2 | 2 | 그림 5 |
| 그림 HWPX | 5 | 2 | 2 | 그림 5 |

HWP 저장의 쪽 수 차이가 큰 국어발전계획과 개체 위치가 바뀌는 그림 샘플은 정확한 제출용 조판으로 쓰지 말고 원본을 대조해야 합니다. 편집 가능한 내용 교환 기능으로 사용하고 `.wordstudio` 백업과 원본을 함께 보관하세요. 원본 보기 경로의 163쪽 취업규칙은 그대로 유지합니다.

회귀 검사는 [단위 시험](tests/hwp-editing.test.cjs), [브라우저 편집/변환 시험](tests/hwp-editing.spec.ts), [설치 앱 시험](scripts/hwp-editing-desktop-smoke.cjs)에 있습니다. 실제 HWP/HWPX 시그니처, Unicode/이모지·정렬·글자 서식, 병합/중첩 표, 그림, 용지와 여백, 수동 페이지 나누기, DOCX 양방향 교환, 취소/쓰기 실패/복구 실패 시 원문 보존을 확인합니다. 공개 샘플 시험은 `HWP_EDITING_PUBLIC_SAMPLES=1 npx playwright test tests/hwp-editing.spec.ts`로 실행하며 샘플 파일은 재배포하지 않습니다. Mac 프로덕션 앱에서 한글 저장 메뉴·네이티브 파일 필터·저장 취소·HWP/HWPX 재열기도 통과했습니다.

최종 소스 `122ad23a750aa05ff747d5383d494ded67eea3fb`의 [교차 플랫폼 검증](https://github.com/beaur8-prog/word-studio/actions/runs/39384292326)이 통과했습니다. Mac arm64·Windows x64·Ubuntu 22.04에서 기능 시험 95개와 단위 시험 17개를 확인했습니다. Windows Setup은 설치·재설치·제거와 문서 보존, HWP/HWPX 저장 메뉴·필터·저장 취소·재열기를 통과했습니다. Linux DEB도 Ubuntu 22.04·24.04에서 같은 설치 파일로 동일한 설치·문서 보존·한글 편집 검사를 통과했습니다. 24.04에서는 같은 DEB의 실제 앱을 검사하며 브라우저/단위 시험을 중복 실행하지 않습니다. Mac 앱의 임시 서명 검증도 통과했습니다. 이 검증 빌드를 1.4.0으로 배포합니다.

### 1.4.0 배포 파일

Windows Setup·Linux DEB는 위 성공한 CI의 산출물을 그대로 사용합니다. Mac ZIP은 새 캐시 폴더에 재추출해 버전 1.4.0·압축 CRC·ad-hoc 서명을 확인했습니다. 시험한 앱과 재추출한 앱의 ASAR SHA-256은 모두 `2eb1beed93cdbbf4b95e88bf702803ef5236012e90f67e112d59d2d77cbedd47`입니다. 배포 준비·다운로드 대조 기록은 로컬 `.cache/word-release-1.4.0/`에 보관합니다.

| 파일 | SHA-256 |
| --- | --- |
| `Word-Studio-1.4.0-windows-x64-setup.exe` | `62a8ecd074de70437cc625aa0e572efc5ca8d46189a96ef609c1495dd9f8a975` |
| `Word-Studio-1.4.0-mac-arm64.zip` | `5ce61ce09b7cb678ddbd1b74a952b6d5f32643755d38af45345e7905a4a33dd2` |
| `Word-Studio-1.4.0-linux-x64.deb` | `0772ecd19878518af491addaac8ee470ecb517aa7ee26b1bd2b31bdae0899fa6` |

## 구현과 사용

파일 → 파일 열기에서 HWP 5.x 또는 HWPX를 선택합니다. 읽기 전용 화면에서 쪽 이동·확대·찾기·텍스트 복사·PDF 저장·인쇄를 사용하고, 닫으면 기존 Word 문서로 돌아옵니다. HWP 원본과 Word 편집 이력은 바꾸지 않습니다.

MIT 라이선스의 RHWP 0.8.7을 별도 Worker에서 실행합니다. 이전 1.3.0은 페이지 분할 패치를 적용한 `0.8.7-wordstudio.1`, 1.3.1은 HWPX manifest 경로 수정까지 포함한 `0.8.7-wordstudio.2`입니다. 정확한 소스 커밋·패치·빌드 방법·배포 바이너리 해시는 [vendor/rhwp-core](vendor/rhwp-core/README.md)에 보존합니다. WebAssembly·한국어 Noto 대체 글꼴을 앱에 포함하며 네트워크에 문서를 보내지 않습니다. SVG의 실행 코드·외부 리소스를 제거하고 쪽별 clip ID를 분리합니다. 45초 제한과 취소 시 Worker 종료를 적용합니다. 파일 40 MiB·1,000쪽, HWPX 내부 10,000개·압축 해제 100 MiB, 출력 데이터 100 MiB 제한입니다. HWP 3.x·암호/DRM은 지원하지 않으며 1.3.x에서는 HWP 저장·편집을 지원하지 않습니다.

중첩 표·글상자를 빠뜨리던 원본 텍스트 API 대신 실제 화면의 layout run을 복사합니다. 긴 취업규칙의 텍스트 복사는 5,901자에서 현재 130,180자로 보완했습니다. 페이지마다 독립적인 SVG 자원 ID를 사용해 다중 쪽 PDF의 표 잘림 영역 충돌도 방지합니다.

## 1.3.1: HWPX manifest 상대 경로 수정

`META-INF/container.xml` → `Contents/content.hpf` → `href="section0.xml"` → `Contents/section0.xml` 구조를 1.3.0 엔진에 넣으면 보고된 필수 파일 누락 오류가 재현됩니다. ZIP 루트에 사본을 추가하면 열리며, 루트 사본의 내용이 다르면 그 잘못된 본문을 선택하는 것도 확인했습니다.

manifest의 section·바탕쪽·내장 바이너리 경로를 OPF 폴더 기준으로 한 번 해석한 뒤 파서와 지연 로더에 전달합니다. 상대 경로가 실제로 존재하면 우선하며, 없으면 기존 한컴 파일의 `Contents/section0.xml`·`BinData/image1.bmp`와 같은 ZIP 루트 표기를 지원합니다. `.`·`..`를 ZIP 내부에서 정리하고 아카이브 밖 경로·절대 경로·원격 URL은 거부합니다. 외부 링크 항목은 다시 쓰지 않습니다. `content.hpf`·`header.xml`의 위치 변경이나 조판 규칙 변경은 이번 수정 범위에 포함하지 않습니다.

두 문단만 있는 합성 A4 HWPX에는 ZIP 루트 section 사본이 없습니다. 회귀 검사는 열기, 루트 표기 대조본과 렌더링 동일성, 동명 루트 파일보다 정상 본문 우선, 하위 폴더와 spine 순서, 필수 파일 누락과 잘못된 경로를 확인합니다. fixture 생성 출처는 [fixtures 설명](tests/fixtures/hwp/README.md)에 있습니다.

실제 공개 HWPX 원본 3개(`07-image`, `09-diagonal`, `11-table-text`)의 `version.xml`에는 `Hancom Office Hangul` 13.0.0.3622 / 12.0.0.535가 기록되어 있습니다. 원본들은 ZIP 루트 표기를 사용합니다. 이 원본 대조군과 **manifest 경로만 상대 표기로 바꾼 파생본**을 구분하여 검사합니다. 후자는 실제 한컴 앱에서 새로 저장한 파일이라는 의미가 아닙니다. [한컴의 구조 설명](https://tech.hancom.com/hwpxformat/)도 본문과 패키지 파일이 `Contents/`에 있음을 설명합니다.

1.3.1 검증 결과:

- 새 경로 회귀 5개는 1.3.0 엔진에서 모두 실패하고 수정 엔진에서 모두 통과합니다. 바탕쪽 본문 표시와 지연 로딩한 내장 PNG의 원본 바이트 보존도 검사합니다. 전체 Node 단위 검사 11개와 브라우저 기능 검사 90개가 통과했습니다.
- 지원 공개 원본 11개의 **227쪽 전체 SVG 해시·쪽수·추출 텍스트가 1.3.0과 동일**합니다. HWPX 상대 표기 파생본 3개는 기존 오류가 사라지고 각각 5·1·1쪽이며, 대응 원본과 전체 SVG·텍스트가 같습니다. 암호/구형 샘플의 거부도 유지됩니다. 이는 기존 결과의 보존 검사이며 원본 뷰어와 완전한 조판 일치를 뜻하지 않습니다.
- 최적화 WASM을 새로 빌드했고, WASM 대상 Rust Clippy(`-D warnings`)와 수정 파일 포맷 검사를 통과했습니다. WASM SHA-256은 `20426b9f42bd63059966da66bbdcaedfb626a432af53b846903c570d2a94859d`입니다. upstream 전체 Rust 회귀·Native Rust/Skia 검증은 실행하지 않았습니다.
- Mac 1.3.1 패키지에 같은 WASM이 들어 있음을 ASAR에서 확인하고 로컬 서명을 검증했습니다. 격리 프로필에서 8개 파일의 열기·전체 텍스트 복사·PDF 저장·기존 Word 문서 보존이 통과했습니다. 새 두 문단 HWPX는 화면·PDF 모두 1쪽, 긴 취업규칙은 모두 163쪽이며 렌더러 오류는 0개입니다. 격리 프로필만 사용하여 사용자 앱의 문서는 바꾸지 않았습니다.
- HWPX 공개 원본 3개와 상대 표기 파생본을 브라우저의 실제 읽기 화면에서도 열었습니다. 총 7쪽의 쪽수·텍스트·그림 수·화면 PNG가 대응 원본과 동일했고, 외부 요청·렌더러 오류는 0개였습니다. Word 편집 문서도 유지됐습니다.
- [Windows·Ubuntu 22.04·24.04 CI](https://github.com/beaur8-prog/word-studio/actions/runs/39374731912)는 코드 커밋 `bb12dc5d399a124efaeb856414912c0aa66f29d8`에서 모두 통과했습니다. Windows와 Ubuntu 22.04는 단위 11개·브라우저 90개를 각각 실행했습니다. Windows Setup 및 두 Ubuntu의 실제 설치 앱에서 5개 HWP/HWPX 화면·PDF 쪽수는 `1, 1, 4, 4, 1`이며 마지막은 새 상대 경로 fixture입니다. 재설치·제거·사용자 데이터 보존도 통과했습니다. Ubuntu 24.04는 22.04에서 만든 같은 DEB를 사용했습니다. 세 플랫폼의 결과 JSON도 내려받아 확인했습니다.
- 1.3.1 배포에는 위 CI에서 시험한 Windows Setup·Linux DEB를 그대로 사용합니다. Mac ZIP은 File Provider 밖의 새 캐시 폴더에 재추출해 버전·압축 CRC·ad-hoc 서명과 ASAR SHA-256 `ae8559f2038c0a0865dc6f6a45c38ca85f5eb6d1d9b55401f1a297ea05976b63` 일치를 확인했습니다. 그 재추출 앱으로 일반 편집·저장·PDF 검사와 위 8개 HWP/HWPX 읽기 검사를 다시 통과했습니다. 배포 준비 기록은 로컬 `.cache/word-release-1.3.1/`에 있습니다.

원시 로그·전후 원본/파생본 결과·비교 결과·브라우저 화면·CI 플랫폼별 증적은 로컬 `.cache/hwpx-manifest-fix/`, Mac 실제 앱 화면·PDF·JSON은 `.cache/hwp-native-MTbwaC/`에 있습니다. `node --test tests/hwpx-manifest.test.cjs`로 새 회귀를 실행하며, `RHWP_CORE_DIRECTORY=/path/to/1.3.0-core`를 지정하면 같은 검사를 이전 엔진에 적용할 수 있습니다.

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

## 1.3.0 초기 검사와 재현

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
