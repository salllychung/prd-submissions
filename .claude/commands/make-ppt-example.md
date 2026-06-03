# /make-ppt 사용 예시 대본

아래를 `/make-ppt` 인수로 그대로 붙여 넣으면 동작합니다.

```
# Slide 1
layout: title-dark
title: 2025 Product Vision
subtitle: 더 빠르고, 더 강하고, 더 얇게.
note: 오프닝 — 30초 유지

---

# Slide 2
layout: section
section_number: 01
title: 문제 정의

---

# Slide 3
layout: bullets
title: 현재 시장의 문제점
bullets:
  - 기존 솔루션의 응답 속도가 3초 이상
  - 경쟁사 대비 2배 이상의 비용 구조
  - 복잡한 온보딩으로 이탈률 68%
  - 실시간 데이터 신뢰도 부족

---

# Slide 4
layout: chart-bar
title: 시장 성장률 비교 (YoY %)
chart_data:
  labels: [2021, 2022, 2023, 2024]
  series:
    - name: "우리 제품"
      data: [12, 28, 45, 72]
    - name: "경쟁사 평균"
      data: [30, 32, 31, 29]

---

# Slide 5
layout: content
title: 핵심 차별화 포인트
body: |
  단 하나의 플랫폼으로 기획부터 분석까지.
  기존 워크플로우를 유지하면서 AI 기반
  인사이트를 즉시 활용할 수 있습니다.

---

# Slide 6
layout: chart-line
title: 월별 활성 사용자 추이
chart_data:
  labels: [Jan, Feb, Mar, Apr, May, Jun]
  series:
    - name: "2024"
      data: [1200, 1900, 2400, 3100, 4200, 5800]
    - name: "2023"
      data: [800, 900, 950, 1100, 1300, 1500]

---

# Slide 7
layout: table
title: 플랜별 기능 비교
table:
  headers: [기능, Free, Pro, Enterprise]
  rows:
    - [API 호출 / 월, "1,000", "50,000", "무제한"]
    - [팀 멤버 수, 1명, 10명, 무제한]
    - [SLA, 없음, 99.5%, 99.99%]
    - [전담 지원, ❌, ❌, ✅]

---

# Slide 8
layout: quote
quote: "디자인은 어떻게 보이느냐가 아니라 어떻게 작동하느냐다."
attribution: — Steve Jobs

---

# Slide 9
layout: chart-pie
title: 매출 구성 비율 (2024)
chart_data:
  labels: [Enterprise, Pro, Free→Paid, Partnership, Other]
  data: [45, 30, 15, 7, 3]

---

# Slide 10
layout: two-column
title: Before vs. After
left:
  heading: "기존 방식"
  bullets:
    - 수동 데이터 수집 (3일)
    - Excel 기반 분석
    - 주 1회 리포팅
right:
  heading: "우리 솔루션"
  bullets:
    - 자동 수집 (실시간)
    - AI 인사이트 즉시 제공
    - 실시간 대시보드

---

# Slide 11
layout: closing
title: 함께 만들어 갈 미래
subtitle: 지금 시작하세요.
cta: 무료 체험 시작
```
