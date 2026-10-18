# Word 사용 사례 검증 · 1.2.6

> 아래는 1.2.6 당시의 비교 기록입니다. 1.2.10 소스·로컬 빌드에서는 공식 샘플 20개를 기준으로 다단·연속 구역·구역별 편집 폭·머리글 내부 개체·링크·기호·각주 등을 추가 수정했습니다. 중앙 그림의 양쪽 흐름과 긴 다단의 편집 페이지 연결도 수정했습니다. 최신 결과와 검증 범위는 [공식 샘플 수정 기록](OFFICIAL_SAMPLE_FIXES.md)을 참고하세요.


2026-10-18, Windows x64에서 Microsoft 공식 안내의 사용 사례를 Word Studio에 적용했다. 공개 DOCX 10개와 직접 만든 합성 DOCX를 사용했으며 개인 문서는 fixture나 저장소에 포함하지 않았다. Word의 실제 화면·인쇄 미리보기, Studio의 편집 화면·인쇄 미리보기·PDF를 비교했다. 쪽수가 같다는 사실만으로 레이아웃이 같다고 판정하지 않았다.

## 공식 사용 사례와 재현 결과

| 사용 사례 / 기준 | 발견한 문제 | 조치와 검증 |
|---|---|---|
| [사진의 표시 크기 조절](https://support.microsoft.com/en-us/word/change-the-size-of-a-picture-shape-text-box-or-wordart-in-word) | 원본 픽셀 크기로 가져와 375.6×329.47px 그림이 57×50px로 축소됨 | OOXML EMU 크기를 읽어 표시 크기 보존. 재저장·재열기·인쇄에서도 검사 |
| [본문과 같은 줄에 그림 삽입](https://support.microsoft.com/en-us/office/change-default-text-wrap-setting-for-images-3f9f0e39-9114-4138-90f9-5c3e528862bb) | 문단 안 그림을 별도 블록으로 옮겨 빈 문단·추가 간격 발생 | 인라인 그림 노드 추가. 기존 Word Studio 블록 그림 파일도 계속 지원 |
| 표의 생략된 칸·가로/세로 병합 | 공개 문서 2개가 표 열 수 불일치로 전체 열기 실패 | DOCX 격자를 변환 전에 보완하고 가로 병합을 정규화. 세로 병합·본문·재저장 검사. 자체 JSON의 엄격한 검증은 유지 |
| [첫 페이지 머리글·번호 다르게 지정](https://support.microsoft.com/en-us/word/start-page-numbering-later-in-your-document) | 첫 페이지에도 기본 머리글·번호 표시 | 첫쪽 내용과 번호 설정을 독립 보존. 설정 창에서 편집 가능 |
| [홀수·짝수 머리글 구분](https://support.microsoft.com/en-US/Word/create-different-headers-or-footers-for-odd-and-even-pages) | 첫쪽·짝수 내용을 기본 머리글로 통합 | 각 유형의 상속·명시적 빈 내용·인쇄·DOCX 왕복 검사 |
| [탭과 점선 리더로 금액 정렬](https://support.microsoft.com/en-us/word/show-dots-or-leaders-between-tabs) | 탭을 공백으로 바꾸어 정렬·점선 손실 | 탭 노드와 문단 탭 위치 보존. 왼쪽·가운데·오른쪽 정렬 및 리더 렌더링. 입력 후 재측정·PDF 확인 |
| [탭 위치 지정](https://support.microsoft.com/en-us/word/insert-or-add-tab-stops) | 기본 탭 입력·음수 탭과 연속 탭의 정렬 오류 | 표/목록/코드 블록 밖에서 Tab 입력 지원. 음수 위치와 본문 시작점 0의 탭 검증 |
| [제본 여백](https://support.microsoft.com/en-gb/word/set-gutter-margins) | 360twip 제본 여백이 누락되어 본문이 24px 이동 | 왼쪽/위 제본 여백과 머리글·바닥글 거리를 읽음. 저장 왕복 시 중복 적용 방지 |
| [문자 종류별 글꼴](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/aef3c9a6-5d6c-434b-90b7-85e761fd8e62) | 테마 글꼴 누락, 영문에도 동아시아 글꼴 적용 | 테마의 major/minor 글꼴 해석. 한글·영문 혼합 구간을 분리하고 직접 서식 우선순위·왕복 검사 |
| [긴 표의 머리행 반복](https://support.microsoft.com/en-gb/word/repeat-table-header-on-subsequent-pages) | 기본 사례에서 새 결함 없음 | 80행 표의 각 출력 페이지에 머리행 반복, 본문 80행 보존 검사 |
| 다중 셀 서식·행 높이·실행 취소 | 기존 수정에 대한 회귀 검사 | 우클릭, 색·테두리, 병합·분할, 행 높이, 실행 취소, 자체 파일/DOCX/PDF 검사 |
| 찾기·바꾸기, 메모, 문서 전환·복구 | 기존 기능에 대한 회귀 검사 | 찾기/일괄 바꾸기, 메모 유지, 이전 문서 실행 취소 이력 분리, 저장 실패 보호 검사 |

## 공개 샘플 재검사

| 샘플 | Word 쪽수 | Studio 1.2.5 | Studio 1.2.6 미리보기/PDF |
|---|---:|---|---:|
| calibre 종합 데모 | 8 | 열기 실패 | 7 / 7 |
| docx4j 표 모음 | 3 | 열기 실패 | 4 / 4 |
| docx4j 인라인 그림 | 1 | 1쪽, 위치 차이 | 1 / 1 |
| docx4j 번호 목록 | 1 | 1쪽, 번호 형식 차이 | 1 / 1 |
| docx4j 첫쪽/홀짝 머리글 | 3 | 모든 쪽에 기본 머리글 | 3 / 3 |
| python-docx 가로/세로 구역 | 5 | 제본 여백 누락 | 5 / 5 |
| python-docx 탭 | 1 | 공백으로 변환 | 1 / 1 |
| LibreOffice 다단 | 1 | 2쪽, 단 정보 손실 | 2 / 2 |
| docx4j 지정 크기 그림 | 1 | 57×50px | 1 / 1, 약 376×329px |
| LibreOffice 그림 둘러싸기 | 1 | 64×64px, 둘러싸기 손실 | 1 / 1, 186×186px, 둘러싸기 미지원 |

샘플 출처: [calibre 공식 변환 데모](https://manual.calibre-ebook.com/conversion.html#convert-microsoft-word-documents), [docx4j](https://github.com/plutext/docx4j), [python-docx](https://github.com/python-openxml/python-docx), [LibreOffice](https://github.com/LibreOffice/core). 다운로드 URL·버전·SHA-256과 양쪽 화면은 로컬 분석 보고서의 manifest에 보관했다.

## 자동·실행 파일 검증

- `tests/word-use-cases.spec.ts`: 새 회귀 시험 11개. 합성 DOCX는 실행 중 생성하여 제3자 샘플을 배포하지 않는다.
- 브라우저 전체 38개, 네이티브 복구 저장 1개 통과.
- 1.2.6 Windows EXE에서 공개 DOCX 10개를 실제 가져오기·미리보기·PDF로 시험. 열기 실패 0개, 각 미리보기/PDF 쪽수 일치, 렌더러 오류 0개.
- 패키지 데스크톱 시험: 표 메뉴·행 높이, 저장 실패 보호·사본 저장, 샌드박스 유지, 8쪽 미리보기/PDF 일치.
- 화면·PDF·JSON은 Git에서 제외한 `analysis/use-case-validation-2026-10-18/`에 보관한다.

## 당시(1.2.6)의 한계

아래 내용은 이전 버전의 기록입니다. 다단·연속 구역·구역별 용지·머리글 내부 개체·부동 그림의 양쪽 흐름과 다각형 윤곽은 1.2.9~1.2.10에서 추가 처리했습니다.

다단과 연속/홀짝 구역 나눔, 그림의 임의 위치·다각형 둘러싸기, 복잡한 다단계 번호 형식, 표의 정밀한 자동 너비 계산은 Word와 차이가 남는다. 머리글·바닥글의 텍스트/유형/번호 설정은 보존하지만 내부 표·그림·서식과 복잡한 필드를 모두 재현하지는 않는다. 편집 화면의 여러 구역 용지 크기는 여전히 첫 구역 기준이며 출력은 구역별 크기를 사용한다. 양면 제본의 좌우 대칭 여백, 사용자 정의 기본 탭 간격, 모든 언어의 글꼴 선택 규칙도 추가 검증 대상이다.

이번 작업은 Word 전체 기능의 동등성을 검증한 것이 아니다. 위 공개 DOCX의 Word 대조는 Windows에서 수행했다. 최신 Mac·Windows·Linux 배포 검증은 [VALIDATION.md](VALIDATION.md)를 참고하며 실물 프린터는 별도로 검증하지 않았다. 이미 변환된 파일에는 잃어버린 정보가 없으므로 새 가져오기 결과를 보려면 **현재 편집 내용을 사본으로 저장한 뒤 원본 DOCX를 다시 열어야 한다**.
