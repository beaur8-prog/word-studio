# 2026-11-10 레이아웃 검증 증거

제품 a69b0fd, 기준5cb4246. 결과는 ../../WORD_LAYOUT_4H_20261110_RESULT.md, 진행은 progress.md.

- government-final-comparison.json: 같은 실제 Word 입력29개. studioPDF/studioState/sourceDOCX/microsoftPDF 경로와 입력/PDF 해시, 개발17·중간6·최종6.
- render-frozen-baseline/: 동일 현재 폰트 환경의5cb4246 전체29개 기준 PDF.
- render-rows-development/, render-rows-final-groups/, render-rows-final-remaining/: 최종29개 분할 산출물. 이후 a69 변경은 편집 전용이고 인쇄·import 코드 지문은 동일하다.
- government-source-audit-summary.json: 원본 모델 DOM의 문단33402/그림219 보존, 반복 머리행을 분리한 정확한 문자열·src 비교. 모든 잉크가 보인다는 증명은 아니다.
- native-baseline/, native-final2/: Word원본10의 실제 Electron 기준·최종. word-native-final2-summary.json은 고정398/140 비교다.
- native-government-final-fast/: 큰정부3 실제편집·undo/redo·저장·PDF. 글자대응은 생략/null. 편집과인쇄 쪽차는 남아있다.
- roundtrip-word-final/, roundtrip-legacy-final/: 실제 두 저장 파일. SDK입력71개 복제본/도구는 지웠으며 sdk-input-manifest-final.json의 source와SHA로 원본을 추적한다. SDK공식결과는 sdk-final2.log.
- legacy-final-summary.json: 같은 준비HWP의 두 번째 DOCX를 이전·현재제품으로 비교. 새 HWP파싱이나 다른원본160쪽 경로 수치와 혼합하지 않는다.
- next-page-divergence-evidence.json, development-actual-font-pairs.json: 남은 공통원인 후보. 파일명별 제품예외를 만들지 않았다.
- cleanup-manifest.json: 임시도구·중간사본275MiB 정리 목록. 원본/기준/최종·현재최소행높이실험·실패자료는 보존.

## 다시 측정

프로젝트 루트에서 실행한다. 과거 원본과 Microsoft PDF는 manifest가 가리키는 기존 analysis 폴더에 있어야 한다. 현재 지원하는 렌더 모드는 current와 frozen-product이며 이전 실험코드는 Git b45d1a3와 progress.md에 남아 있다.

```sh
LAYOUT_VARIANT=frozen-product LAYOUT_STAGE=frozen-baseline LAYOUT_IDS=all node analysis/word-layout-4h-2026-11-10/tools/render-inputs.cjs
LAYOUT_VARIANT=current LAYOUT_STAGE=current LAYOUT_IDS=all node analysis/word-layout-4h-2026-11-10/tools/render-inputs.cjs
LAYOUT_STAGE=frozen-baseline python3 analysis/word-layout-4h-2026-11-10/tools/paragraph-probe.py
LAYOUT_WORD_STAGE=native-final2 python3 analysis/word-layout-4h-2026-11-10/tools/score-word.py
python3 analysis/word-layout-4h-2026-11-10/tools/verify-retained.py
```

PDF Python 도구는 pdfplumber가 필요하다. 실제앱 도구는 사용자앱과 분리한 native-candidate-app 및 임시프로필을 사용한다. 일반앱/웹 Word 제어는 잠금해제 때만 한다. 공개배포·리토스·버전변경은 이 작업에 포함하지 않았다.
