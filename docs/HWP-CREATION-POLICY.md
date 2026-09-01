# 한글 새 문서 제작 경로 정책

DocBridge의 새 문서 기본값은 `DOCX 우선 → HWPX 변환 → HWP 직접 미세 수정`이다. 기존 HWP/HWPX나 한글 고유 기능이 필요한 작업만 처음부터 HWP 직접 편집을 사용한다. AI 클라이언트는 새 문서를 만들기 전에 읽기 전용 `hwp_plan_creation`을 호출하고 반환된 `mode`를 따른다.

## 결정표

| 조건 | 반환 mode | 실행 경로 |
|---|---|---|
| 새 일반 문서: 문단·표·그림·단순 머리말/꼬리말 | `docx-first` | DOCX 작성·전쪽 렌더 검수 → `hwp_launch` OOXML 가져오기 → 구조/PDF 검증 |
| 기존 HWP/HWPX 중간 수정 | `native-hwp` | 문서 `documentRef` 고정 → 직접 편집 |
| 기존 한글 템플릿 | `native-hwp` | 템플릿 필드·원본 구조 보존 편집 |
| 한글 필드·누름틀·한글 전용 개체 | `native-hwp` | 네이티브 HWP 작업 |
| 복잡한 병합표 또는 원본과 동일한 배치 | `native-hwp` | 네이티브 HWP 작업 |
| 검수 가능한 DOCX 생성 도구가 없음 | `native-hwp` | 빈 문서를 한 번만 만들고 단계별 직접 편집 |

새 일반 문서에서 `docx-first`가 반환됐는데 AI가 편의상 `newDocument=true`로 바꾸면 안 된다. 반대로 기존 HWP 양식이나 한글 고유 기능을 DOCX 변환으로 우회해서도 안 된다.

## 검증 근거

2026-08-13 동일 PC의 대표 A4 표 문서와 기존 DocBridge 감사 로그를 비교했다.

| 측정 항목 | 결과 |
|---|---:|
| DOCX 생성 5회 평균 | 0.264초 |
| DOCX → HWPX 내부 변환 5회 평균 | 0.668초 |
| DOCX 우선 핵심 작업 합계 | 약 0.93초 |
| 별도 프로세스 기동·검증 포함 CLI 평균 | 약 4.45초 |
| HWP 복합 적용 21건 중앙값 | 2.746초 |
| 빈 HWP에서 표·서식을 여러 단계로 구성한 대표 작업 | 약 7.9초 이상 |

직접 HWP 로그는 서로 다른 문서이므로 정확한 동일 문서 배율 비교는 아니다. 그러나 새 일반 문서에서는 DOCX 우선 경로가 COM 호출 수와 반복 서식 조정을 크게 줄인다는 결론에는 충분하다.

대표 DOCX는 한글 `FileOpen(OOXML)`과 `FileSaveAs(HWPX)`를 거쳐 1쪽·표 6개·필수 문구 보존을 4회 연속 통과했다. 변환 중 원본 SHA-256은 바뀌지 않았고 HWP 잔류 프로세스도 없었다.

## DOCX 우선 품질 게이트

1. DOCX의 A4 크기, 여백, 표 너비·행 높이, 글꼴, 문단 간격을 명시한다.
2. DOCX를 PDF/PNG로 렌더해 모든 쪽의 잘림·겹침·빈 페이지를 검사한다.
3. `hwp_launch`에 `creationMode:"docx-first"`, 절대 `sourceFile`, 새 `outputFile`, `expectedPageCount`, `expectedTableCount`, `requiredText`를 전달한다.
4. `sourceUnchanged:true`, `verification.passed:true`, 빈 `warnings`를 모두 확인한다.
5. 변환된 HWP/HWPX를 다시 읽어 본문·표·쪽 수를 확인하고 최종 HWP PDF의 모든 쪽을 시각 검수한다.
6. 실패하면 완료로 보고하지 않는다. DOCX를 조정하고 기존 출력을 덮어쓰지 않는 새 이름으로 다시 변환한다.

운영 변환 경로는 Word COM을 실행하지 않는다. DOCX 파일을 만든 뒤 한글 Automation이 OOXML을 직접 가져오므로 Word 유령 프로세스와 Word 추가기능 상태에 의존하지 않는다.

## HWP 직접 품질 게이트

1. `hwp_doctor`와 `hwp_get_active_context`로 대상 인스턴스와 문서를 확인한다.
2. 새 HWP 전용 양식만 `hwp_launch({"creationMode":"native-hwp","newDocument":true})`로 한 번 시작한다.
3. 기존 문서는 `documentRef`를 고정하고 주변 문맥·표 구조·서식을 읽은 뒤 편집한다.
4. 모든 쓰기는 dry-run → 동일 ops 적용 → `postEditReread` 검증을 거친다.
5. 표·쪽 수·핵심 문구와 최종 PDF의 모든 쪽을 검수한다.
