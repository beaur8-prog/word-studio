# HWP/HWPX 읽기 검증

2026-10-20 · Word Studio 1.3.0 소스·로컬 변경. 공개 릴리스·리토스 배포는 하지 않았습니다.

## 구현과 사용

파일 → 파일 열기에서 HWP 5.x 또는 HWPX를 선택합니다. 읽기 전용 화면에서 쪽 이동·확대·찾기·텍스트 복사·PDF 저장·인쇄를 사용하고, 닫으면 기존 Word 문서로 돌아옵니다. HWP 원본과 Word 편집 이력은 바꾸지 않습니다.

MIT 라이선스의 `@rhwp/core 0.8.7`을 별도 Worker에서 실행합니다. WebAssembly·한국어 Noto 대체 글꼴을 앱에 포함하며 네트워크에 문서를 보내지 않습니다. SVG의 실행 코드·외부 리소스를 제거하고 쪽별 clip ID를 분리합니다. 45초 제한과 취소 시 Worker 종료를 적용합니다. 파일 40 MiB·1,000쪽, HWPX 내부 10,000개·압축 해제 100 MiB, 출력 데이터 100 MiB 제한입니다. HWP 3.x·암호/DRM·HWP 저장·편집은 지원하지 않습니다.

중첩 표·글상자를 빠뜨리던 원본 텍스트 API 대신 실제 화면의 layout run을 복사합니다. 긴 취업규칙의 텍스트 복사는 5,901자에서 129,913자로 보완했습니다. 페이지마다 독립적인 SVG 자원 ID를 사용해 다중 쪽 PDF의 표 잘림 영역 충돌도 방지합니다.

## 공개 샘플과 비교

파일 출처·정확한 업스트림 커밋·크기·SHA-256은 `tests/hwp-samples.json`에 있습니다. 원본 12개(지원 형식 11개, 거부 확인 1개)와 PDF 참고본 3개를 내려받았습니다. 원본은 Git과 앱 배포물에 포함하지 않습니다.

| 샘플 | 비교 기준 | 기준 쪽수 | Studio 쪽수 | 확인 범위 |
| --- | --- | ---: | ---: | --- |
| 2025 표준근로계약서 HWP | 폴라리스 웹 | 7 | 7 | 전 7쪽 화면 |
| 2025 표준취업규칙 HWP | 폴라리스 웹 | 163 | 154 | 첫·마지막 쪽, 쪽수 차이 확인 |
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

쪽수가 같은 것이 조판 일치를 뜻하지는 않습니다. 설치 글꼴·폭·기준선, 표 행 분할·여백, 복잡한 수식과 도형에는 차이가 남습니다. 특히 163/154쪽 차이가 있는 취업규칙은 정확한 페이지 참조·제출용 출력에 원본 뷰어 확인이 필요합니다. 상세 화면 비교와 측정 자료는 재현 시 `analysis/hwp-validation/report.html`에 생성합니다.

## 검사와 재현

- 기존 브라우저 전체 86개 통과 후, SVG 자원 충돌·저장 실패 창 보호 회귀 2개를 추가해 HWP 검사 7개 통과. 네이티브 단위 시험 4개 통과.
- Mac arm64 패키지에서 실제 파일 열기 IPC·파일 필터, 오프라인 Worker/WASM·한국어 글꼴, OS 클립보드 복사와 원상 복원, PDF 저장, Word 문서 보존을 확인했습니다. 합성 HWP/HWPX 각각 1쪽, 공개 계약서 7쪽, 보건소 표 4쪽은 화면과 실제 PDF 쪽수가 일치합니다. 렌더러 오류는 0개입니다.
- 공개 샘플 11개·화면 37쪽을 검사했고 외부 네트워크 요청·렌더러 오류는 0개입니다. 나머지 암호/구형 1개는 거부합니다.
- Windows와 Ubuntu 설치 검사에도 HWP 열기·복사·PDF 시험을 추가했습니다. 플랫폼별 실행 결과는 후속 CI 기록을 참고하세요.

```sh
npm ci
npm run build
npx playwright test tests/hwp-reader.spec.ts
npm run test:native
# npm run dev로 로컬 서버를 켠 뒤 공개 샘플 다운로드·검사
npm run test:hwp-samples
# 별도 프로필로 실제 Electron/패키지 검사
npm run test:hwp-native
# Mac 실제 배포 앱 검사
WORD_STUDIO_EXECUTABLE='/path/to/Word Studio.app/Contents/MacOS/Word Studio' HWP_PUBLIC_SAMPLES=1 npm run test:hwp-native
```

## 참고와 라이선스

- [한컴 HWP/OWPML 공개 형식](https://download.hancom.com/support/downloadCenter/hwpOwpml)
- [고용노동부 2025 표준근로계약서·취업규칙 원본](https://www.moel.go.kr/info/etc/dataroom/view.do?bbs_seq=20250300356)
- [폴라리스 웹 HWP 뷰어](https://hwp.polarisoffice.com/)
- [RHWP 소스·지원 범위](https://github.com/edwardkim/rhwp) · [MIT 라이선스](electron/licenses/RHWP-MIT.txt)
- [Noto Sans KR 라이선스](electron/licenses/Noto-Sans-KR-OFL.txt) · [Noto Serif KR 라이선스](electron/licenses/Noto-Serif-KR-OFL.txt)
