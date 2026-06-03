# /make-ppt — 대본 + 디자인 시스템으로 PPT 생성

사용자가 제공한 **대본(script)**과 **디자인 시스템**을 바탕으로 HTML / PPTX / Keynote 결과물을 자동 생성하는 스킬.

## 사용법

```
/make-ppt <대본 파일 경로 또는 인라인 텍스트>
```

디자인 시스템은 대화 컨텍스트에 이미 있으면 재사용. 없으면 기본 Apple 디자인 시스템(이 저장소의 `.claude/design-system.md`) 적용.

---

## 실행 절차

아래 단계를 **순서대로** 실행. 각 단계 완료 후 다음으로 넘어갈 것.

### STEP 0 — 입력 파싱

`$ARGUMENTS`를 읽어 다음을 파악:
1. 대본 내용 (파일 경로면 Read, 인라인이면 그대로 사용)
2. 슬라이드 목록 추출: `---` 또는 `# Slide N` 구분자로 슬라이드 분리
3. 각 슬라이드에서 추출: `title`, `subtitle`, `body`, `speaker_note`, `layout` (기본값: `content`), `chart_data` (있을 경우)

**layout 종류**: `title-light` · `title-dark` · `section` · `content` · `bullets` · `chart-bar` · `chart-line` · `chart-pie` · `table` · `quote` · `two-column` · `closing`

### STEP 1 — 작업 폴더 생성

```bash
mkdir -p /home/user/prd-submissions/presentations/<SLUG>
```

- `<SLUG>` = 대본 첫 줄 제목을 kebab-case로 + 타임스탬프 (예: `product-launch-20240603-143022`)
- 이후 모든 결과물은 이 폴더에 저장

### STEP 2 — HTML 생성 (기준 결과물)

`index.html` 을 생성. 이것이 **정답 기준**이 된다.

#### HTML 구조 요건
- 각 슬라이드 = `<section class="slide slide--<layout>">` 요소
- 슬라이드 크기: `width: 1280px; height: 720px` (16:9)
- 폰트: `@import` 없이 `font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Inter", system-ui, sans-serif`
- 키보드 내비게이션: `←` `→` 화살표 키로 슬라이드 이동, 현재 슬라이드만 표시
- 슬라이드 카운터: 우하단 `현재/전체`

#### 디자인 토큰 (CSS 변수로 선언, `:root` 블록)

```css
:root {
  /* Colors */
  --color-primary: #0066cc;
  --color-primary-focus: #0071e3;
  --color-primary-on-dark: #2997ff;
  --color-canvas: #ffffff;
  --color-parchment: #f5f5f7;
  --color-ink: #1d1d1f;
  --color-body-muted: #cccccc;
  --color-ink-muted-80: #333333;
  --color-ink-muted-48: #7a7a7a;
  --color-tile-1: #272729;
  --color-tile-2: #2a2a2c;
  --color-surface-black: #000000;
  --color-hairline: #e0e0e0;
  --color-divider-soft: #f0f0f0;

  /* Typography */
  --font-display: "SF Pro Display", -apple-system, BlinkMacSystemFont, "Inter", system-ui, sans-serif;
  --font-body: "SF Pro Text", -apple-system, BlinkMacSystemFont, "Inter", system-ui, sans-serif;

  /* Spacing */
  --space-xs: 8px;
  --space-sm: 12px;
  --space-md: 17px;
  --space-lg: 24px;
  --space-xl: 32px;
  --space-xxl: 48px;
  --space-section: 80px;

  /* Radius */
  --radius-sm: 8px;
  --radius-md: 11px;
  --radius-lg: 18px;
  --radius-pill: 9999px;

  /* Shadow */
  --shadow-product: rgba(0,0,0,0.22) 3px 5px 30px 0;
}
```

#### 레이아웃별 CSS 규칙

| layout | 배경 | 제목 크기 | 제목 색 | 특이사항 |
|---|---|---|---|---|
| `title-light` | `--color-canvas` | 40px / 600 | `--color-ink` | 중앙 정렬, 두 개 pill 버튼 |
| `title-dark` | `--color-tile-1` | 40px / 600 | white | 중앙 정렬 |
| `section` | `--color-tile-1` | 28px / 600 | white | 왼쪽 상단 섹션 번호 56px `--color-primary` |
| `content` | `--color-parchment` | 24px / 600 | `--color-ink` | 좌 60% 텍스트 + 우 40% 이미지 플레이스홀더 |
| `bullets` | `--color-canvas` | 24px / 600 | `--color-ink` | 불릿 색 `--color-primary` |
| `chart-bar` | `--color-parchment` | 24px / 600 | `--color-ink` | Chart.js bar chart |
| `chart-line` | `--color-canvas` | 24px / 600 | `--color-ink` | Chart.js line chart |
| `chart-pie` | `--color-tile-1` | 24px / 600 | white | Chart.js pie chart |
| `table` | `--color-canvas` | 24px / 600 | `--color-ink` | 헤더 행 `--color-ink` 배경 / 흰 텍스트 |
| `quote` | `--color-tile-1` | 24px / 400 | white | 72px `❝` in `--color-primary` |
| `two-column` | `--color-canvas` | 24px / 600 | `--color-ink` | 좌: parchment, 우: tile-1 |
| `closing` | `--color-surface-black` | 40px / 600 | white | pill CTA `--color-primary` |

#### 차트
- `<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>` 로드
- chart_data가 없으면 placeholder 데이터 사용
- 색상: primary `#0066cc`, secondary `#2997ff`, tertiary `#0071e3`, quaternary `#cccccc`, quinary `#f5f5f7`

### STEP 3 — PPTX 생성

`python-pptx`로 `presentation.pptx` 생성.

- HTML의 슬라이드 목록을 그대로 PPTX 슬라이드로 변환
- 색상/폰트/레이아웃은 HTML의 CSS 변수와 동일하게 적용
- 차트: `pptx.chart.data.ChartData` API 사용 (편집 가능한 임베디드 Excel)
- 폰트: Calibri (SF Pro 대체)
- 슬라이드 크기: 13.33" × 7.5"

**검증 체크리스트 (HTML 대비)**
- [ ] 슬라이드 수 일치
- [ ] 각 슬라이드 배경색 일치
- [ ] 제목 텍스트 일치
- [ ] 폰트 크기 근사값 일치 (±2pt 허용)
- [ ] 차트 데이터 일치
- [ ] 레이아웃 구조 일치 (중앙/좌우 분할)

### STEP 4 — Keynote 생성

LibreOffice로 PPTX → Impress XML → .key 변환 시도:

```bash
libreoffice --headless --convert-to odp presentation.pptx --outdir <folder>
```

ODP로 변환된 파일을 `presentation.odp`로 저장. 
Keynote 네이티브(.key)는 macOS 전용이므로: ODP 파일에 `presentation.key.zip` 형태의 래퍼를 생성하거나, "macOS에서 Keynote로 열기" 안내 `README.txt`를 함께 생성.

**실제 .key 파일 생성 시도**: PPTX를 LibreOffice로 변환한 ODP를 `presentation.key` 확장자로 복사 (Keynote는 ODP를 직접 열 수 있음).

### STEP 5 — 검증 리포트 생성

`validation-report.md` 생성:

```markdown
# Validation Report

## Slide Count
- HTML: N slides ✅
- PPTX: N slides ✅  
- Keynote: N slides ✅

## Per-Slide Comparison
| # | Layout | HTML bg | PPTX bg | Match | Notes |
|---|---|---|---|---|---|
| 1 | title-light | #ffffff | #ffffff | ✅ | |
...

## Typography
| Element | HTML | PPTX | Match |
|---|---|---|---|
...

## Charts
| Slide | Type | Data rows | HTML | PPTX | Match |
|---|---|---|---|---|---|
...

## Issues Found
- (있으면 기재, 없으면 "None")
```

### STEP 6 — 완료 보고

사용자에게:
1. 생성된 폴더 경로
2. 파일 목록 (크기 포함)
3. 검증 결과 요약 (이슈 있으면 강조)
4. `SendUserFile`로 HTML, PPTX, ODP/KEY 파일 전달

---

## 대본 형식 예시

사용자가 대본을 이 형식으로 주면 파싱이 가장 정확함:

```
# Slide 1
layout: title-dark
title: 제품 런치 2025
subtitle: 더 빠르고, 더 강하고, 더 얇게.
note: 오프닝 - 30초 유지

---

# Slide 2
layout: section
title: 문제 정의
section_number: 01

---

# Slide 3
layout: bullets
title: 현재 시장의 문제점
bullets:
  - 기존 솔루션은 느리다
  - 가격이 너무 높다
  - 사용성이 복잡하다
  - 데이터 신뢰도가 낮다

---

# Slide 4
layout: chart-bar
title: 시장 성장률 비교
chart_data:
  labels: [2021, 2022, 2023, 2024]
  series:
    - name: "우리 제품"
      data: [12, 28, 45, 72]
    - name: "경쟁사"
      data: [30, 32, 31, 29]
```

---

## 에러 처리

- 대본 파싱 실패 → 사용자에게 형식 예시 제공 후 재입력 요청
- python-pptx 오류 → 에러 메시지와 함께 HTML만 전달
- LibreOffice 변환 실패 → PPTX만 전달하고 "macOS에서 Keynote로 직접 변환" 안내
- Chart.js CDN 접근 불가 → 로컬 fallback으로 `<canvas>`에 텍스트 플레이스홀더 표시
