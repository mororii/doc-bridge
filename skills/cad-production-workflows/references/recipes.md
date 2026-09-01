# CAD 실무 레시피

## 새 객체 작성

1. 대상 도면과 단위 `INSUNITS`를 확인한다.
2. 작성 전 ModelSpace 개수와 대상 영역 객체수를 기록한다.
3. `draw_entities`에 같은 목적의 도형을 최대 1,000개씩 묶는다.
4. 적용 후 기존 시작 인덱스부터 증분 조회하고 유형별 개수와 bbox를 검증한다.

지원 타입: `lwpolyline`, `circle`, `hatch`, `block`, `text`, `line`, `arc`, `ellipse`, `point`, `mtext`, `dim_aligned`, `dim_rotated`.

## 기존 객체 수정

- 반드시 조회에서 얻은 핸들을 사용한다.
- 복사·대칭·offset은 새 객체를 만들므로 적용 전후 ModelSpace 증가량을 확인한다.
- 축척과 이동은 기준점을 명시한다.
- 여러 op 중 하나가 실패하면 자동 롤백 결과를 확인하고 같은 배치를 그대로 재실행하지 않는다.

## 여러 도곽 일괄 검증

```json
{
  "scope": "regions",
  "regions": [
    {"name":"plan","bounds":{"minX":0,"minY":200,"maxX":280,"maxY":260},"minCount":1},
    {"name":"profile","bounds":{"minX":0,"minY":0,"maxX":280,"maxY":195},"minCount":1},
    {"name":"keymap","bounds":{"minX":285,"minY":220,"maxX":310,"maxY":260},"entityTypes":["Polyline"],"minCount":1}
  ]
}
```

`verified:false`인 영역이 있으면 쓰기를 끝냈다고 보고하지 않는다.

## 배치·PDF

- `configure_layout`의 장치명은 설치된 PC3 이름을 사용한다. 일반적으로 `DWG To PDF.pc3`이다.
- 뷰포트 크기는 종이공간 단위, `viewHeight`는 모델공간 단위다.
- `plot_pdf`는 high-risk이며 기존 PDF를 교체할 수 있다.
- 결과 PDF의 페이지 크기·회전·여백·문자 가독성을 렌더링으로 확인한다.
