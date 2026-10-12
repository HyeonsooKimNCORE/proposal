// KING OF KIM 온라인 판매 활성화 제안서 — NICECORE 템플릿(nicecore_proposal_pt3.pptx) 적용판
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const A = (f) => path.join(__dirname, f);

// ---- Colors ----
// 브리프 블루(primary 계열)를 템플릿 네이비 위에서 읽히도록 조정
const K = {
  primary: "3182F6", primaryDark: "1B64DA", primarySoft: "B0D4FF",
  white: "FFFFFF", body: "D5DDE8", muted: "9FB0C5",
  danger: "FF6F7A",
  navy: "364960", navyLight: "52667F",
  // 표 (네이비 위 단색)
  tRow: "415369", tAlt: "3A4B61", tHi: "2B5596", tTot: "506580", tBorder: "5A6D84",
  grid: "5C6F86",
};
const THEME = {
  name: "NICECORE",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "191F28", lt1: "FFFFFF", dk2: K.navy, lt2: "F2F4F6",
    accent1: K.primary, accent2: K.primaryDark, accent3: "E8F3FF", accent4: K.primarySoft,
    accent5: K.navyLight, accent6: "F04452", hlink: K.primarySoft, folHlink: K.muted,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 (템플릿과 동일)
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "NICECORE";
pres.company = "NICECORE";
pres.title = "KING OF KIM 온라인 판매 활성화 제안서";

// 템플릿 본문 그리드: 좌 1.3" ~ 우 12.03" (템플릿 타이틀 선 끝)
const X0 = 1.3, XR = 12.03, CW = XR - X0, GAP = 0.3, Y0 = 2.1, YB = 6.45;
const W3 = (CW - 2 * GAP) / 3;
const COL3 = [X0, X0 + W3 + GAP, X0 + 2 * (W3 + GAP)];
const R = 0.12;

// ---- Template frame (layouts) ----
const frame = [
  { image: { x: 11.097, y: 6.655, w: 1.703, h: 0.25, path: A("logo_white.png") } },
  { text: { text: "Credible Outcome Reliable Evolution", options: { x: 11.097, y: 7.09, w: 1.95, h: 0.18, fontSize: 7.5, color: K.white, margin: 0, valign: "middle", fontFace: "Arial" } } },
];
pres.defineSlideMaster({
  title: "COVER",
  background: { color: K.navyLight },
  objects: [
    { image: { x: 0, y: 0, w: 13.333, h: 7.5, path: A("tpl_cover_bg.png") } },
    ...frame,
    { text: { text: "©2026 NICECORE. ALL RIGHTS RESERVED.", options: { x: 0.219, y: 7.12, w: 3, h: 0.16, fontSize: 6, color: K.white, margin: 0, valign: "middle", fontFace: "Arial" } } },
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 1.85, w: 10.5, h: 1.6, fontSize: 40, bold: true, color: K.white, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: K.navy },
  margin: [0.5, 1.3, 1.0, 1.3],
  objects: [
    { image: { x: 0, y: 0, w: 13.333, h: 7.5, path: A("tpl_content_bg.png") } },
    ...frame,
    { text: { text: "©2026 NICECORE. ALL RIGHTS RESERVED.  |  KING OF KIM 온라인 판매 활성화 제안", options: { x: 0.219, y: 7.1, w: 6, h: 0.2, fontSize: 8, color: K.white, margin: 0, valign: "middle" } } },
    // 템플릿 '서브타이틀' 아래 짧은 선
    { line: { x: X0, y: 1.85, w: 2.13, h: 0, line: { color: K.white, width: 1.5, transparency: 80 } } },
    { placeholder: { options: { name: "header", type: "body", x: X0, y: 0.52, w: CW, h: 0.45, fontSize: 24, color: K.white, align: "left", valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 1.17, w: CW, h: 0.55, fontSize: 26, bold: false, color: K.white, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 10.35, y: 7.08, w: 0.6, h: 0.2, fontSize: 8, color: K.white, align: "right", margin: 0 },
});

// ---- Helpers ----
let objN = 0;
const glue = (t) => t.replace(/ · /g, " · ").replace(/만 원/g, "만 원");
function T(s, text, o) {
  text = typeof text === "string" ? glue(text) : text.map((r) => Object.assign({}, r, { text: glue(r.text) }));
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: K.body, valign: "top", lineSpacingMultiple: 1.15, objectName: `text-${++objN}` }, o));
}
// kind: glass(반투명 흰 카드) | primary(브리프 블루) | tint(블루 강조 박스) | outline | white
function card(s, x, y, w, h, kind = "glass") {
  const fill = {
    glass: { color: K.white, transparency: 90 },
    primary: { color: K.primary },
    tint: { color: K.primary, transparency: 65 },
    outline: { color: K.white, transparency: 96 },
    white: { color: K.white },
  }[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: R, fill,
    line: kind === "outline" ? { color: K.white, width: 1, transparency: 65 } : { type: "none" },
    objectName: `card-${++objN}`,
  });
}
function badge(s, x, y, d, label, fill = K.primary, color = K.white, size = 16) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: `badge-${++objN}` });
  s.addText(label, { isTextBox: true, x, y, w: d, h: d, margin: 0, align: "center", valign: "middle", fontSize: size, bold: true, color, objectName: `badge-text-${objN}` });
}
// 헤더 텍스트 폭 추정 (타이틀 오른쪽 가로선 시작점 계산용, 보수적으로)
function estW(text, pt, bold) {
  let u = 0;
  for (const ch of text) {
    if (/[가-힣]/.test(ch)) u += 1.0;
    else if (/[0-9]/.test(ch)) u += bold ? 0.62 : 0.58;
    else if (/[A-Z]/.test(ch)) u += 0.75;
    else if (/[a-z]/.test(ch)) u += 0.55;
    else if (ch === " ") u += 0.3;
    else u += 0.5;
  }
  return (u * pt) / 72;
}
function content(section, eyebrow, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  const m = eyebrow.match(/^(\d+)\s+(.*)$/);
  const runs = m
    ? [{ text: m[1], options: { bold: true } }, { text: "  " + m[2] }]
    : [{ text: eyebrow, options: { bold: true } }];
  s.addText(runs, { placeholder: "header" });
  // 템플릿 타이틀 선: 흰 짧은 선 + 20% 긴 선 (헤더 텍스트 오른쪽)
  const tw = m ? estW(m[1], 24, true) + estW("  " + m[2], 24) : estW(eyebrow, 24, true);
  const sx = Math.min(X0 + tw + 0.55, 9.5), ly = 0.9;
  s.addShape(pres.shapes.RECTANGLE, { x: sx, y: ly, w: 0.85, h: 0.03, fill: { color: K.white }, line: { type: "none" }, objectName: "title-rule-strong" });
  s.addShape(pres.shapes.RECTANGLE, { x: sx + 0.85, y: ly, w: XR - sx - 0.85, h: 0.03, fill: { color: K.white, transparency: 80 }, line: { type: "none" }, objectName: "title-rule-faint" });
  s.addText(glue(title), { placeholder: "title" });
  return s;
}
function bullets(items, o = {}) {
  return items.map((t, i) => ({
    text: t,
    options: Object.assign({ bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 6 }, o),
  }));
}
function caption(s, text, y = 6.55, w = 9.5) {
  T(s, text, { x: X0, y, w, h: 0.42, fontSize: 12, color: K.muted });
}
const chartFont = {
  titleFontFace: "+mn-lt", legendFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
};

// =========================================================
// Slide 1. 표지 (템플릿 표지)
// =========================================================
pres.addSection({ title: "표지" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  T(s, "제안서 · 2026.10", { x: X0, y: 1.35, w: 6, h: 0.35, fontSize: 16, bold: true, color: K.primarySoft });
  s.addText([
    { text: "KING OF KIM", options: { breakLine: true } },
    { text: "온라인 판매 활성화 6개월 Pilot" },
  ], { placeholder: "title", lineSpacingMultiple: 1.1 });
  T(s, "네이버 스마트스토어 · 쿠팡 퍼포먼스 마케팅 및 리뷰 기반 구축", { x: X0, y: 3.7, w: 10, h: 0.45, fontSize: 20, color: K.white });
  T(s, "주식회사 양명 귀중 | NICECORE", { x: X0, y: 6.45, w: 6, h: 0.35, fontSize: 14, color: K.white, valign: "middle" });
  s.addNotes("RFP 중 온라인 판매 활성화(6.5)와 검색광고 기반 인지도 확대(6.4 일부)에 집중한 범위 제안임을 먼저 밝힙니다.");
}

// =========================================================
// Slide 2. 제안 요약
// =========================================================
pres.addSection({ title: "제안 개요" });
{
  const s = content("제안 개요", "01 제안 요약", "6개월 동안 온라인에서 \"팔리는 증거\"를 만듭니다");
  const stats = [["월 300만 원", "현재 월 온라인 매출"], ["17개", "스마트스토어 상품당 최대 리뷰 수"], ["300%", "4개월차 이후 목표 ROAS"]];
  stats.forEach(([n, l], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, 1.3);
    T(s, n, { x: x + 0.3, y: Y0 + 0.2, w: W3 - 0.5, h: 0.6, fontSize: 30, bold: true, color: K.primarySoft, valign: "middle", fit: "none" });
    T(s, l, { x: x + 0.3, y: Y0 + 0.86, w: W3 - 0.45, h: 0.3, fontSize: 13 });
  });
  const steps = [
    ["1", "리뷰 확보", ["1~2개월차 체험단 집중", "광고는 최소 수준으로 테스트"]],
    ["2", "구매 전환", ["3개월차부터 네이버 쇼핑검색·쿠팡 광고로 구매 확대"]],
    ["3", "단계 증액", ["목표 ROAS 달성 시에만 다음 단계로 광고비 증액"]],
  ];
  const sy = 3.7, sh = 1.65;
  steps.forEach(([n, t, d], i) => {
    const x = COL3[i];
    card(s, x, sy, W3, sh, "outline");
    badge(s, x + 0.3, sy + 0.25, 0.48, n);
    T(s, t, { x: x + 0.92, y: sy + 0.25, w: W3 - 1.2, h: 0.48, fontSize: 18, bold: true, color: K.white, valign: "middle" });
    T(s, d.map((line, j) => ({ text: line, options: { breakLine: j < d.length - 1 } })), { x: x + 0.3, y: sy + 0.88, w: W3 - 0.55, h: 0.68, fontSize: 14 });
  });
  card(s, X0, 5.65, CW, 0.8, "tint");
  T(s, "6개월 성과(판매량 · 리뷰 · 검색량)는 유통 MD 제안과 조달 홍보에 쓸 수 있는 성과 자료로 정리해 드립니다.", { x: X0 + 0.3, y: 5.65, w: CW - 0.6, h: 0.8, fontSize: 15, bold: true, color: K.white, valign: "middle" });
  s.addNotes("현재 매출 규모에서 예산을 여러 채널로 나누면 어느 채널에서도 검증 가능한 성과가 나오기 어렵다는 점을 먼저 설명합니다.");
}

// =========================================================
// Slide 3. 제안 범위
// =========================================================
{
  const s = content("제안 개요", "02 제안 범위", "예산이 성과로 이어지는 온라인 과업에 집중합니다");
  const h = 4.05;
  const ws = [4.1, 2.75, CW - 4.1 - 2.75 - 2 * GAP];
  const xs = [X0, X0 + ws[0] + GAP, X0 + ws[0] + ws[1] + 2 * GAP];
  const cards = [
    { title: "직접 수행", sub: "NICECORE", kind: "primary", items: [
      "네이버 쇼핑검색 · 파워링크 운영", "쿠팡 광고 운영", "체험단 기획 · 운영 (제품은 양명 제공)",
      "운영 데이터 진단 · 월간 리포트 · 예산 재배분", "6개월 성과 자료 정리 (MD · 조달 제안용)"] },
    { title: "별도 견적", sub: "선택 시 협의", kind: "glass", items: ["상세페이지 개선"] },
    { title: "범위 외", sub: "전문 업체 또는 양명", kind: "glass", pairs: [
      ["소비자 조사(설문 · FGI)", "리서치 전문사"], ["시식 · POP · 오프라인 광고", "홍보 · 오프라인 대행사"], ["유통 MD 영업 · 조달청 영업", "양명 직접"]] },
  ];
  cards.forEach((c, i) => {
    const x = xs[i], w = ws[i], hl = c.kind === "primary";
    card(s, x, Y0, w, h, c.kind);
    T(s, c.title, { x: x + 0.3, y: Y0 + 0.3, w: w - 0.6, h: 0.42, fontSize: 21, bold: true, color: K.white });
    T(s, c.sub, { x: x + 0.3, y: Y0 + 0.76, w: w - 0.6, h: 0.3, fontSize: 14, color: hl ? K.primarySoft : K.muted });
    if (c.items) {
      T(s, bullets(c.items), { x: x + 0.3, y: Y0 + 1.25, w: w - 0.5, h: h - 1.4, fontSize: 14, color: K.white });
    } else {
      const runs = [];
      c.pairs.forEach(([a, b], j) => {
        runs.push({ text: a, options: { bullet: { indent: 14 }, breakLine: true, color: K.white } });
        runs.push({ text: "→ " + b, options: { indentLevel: 1, bullet: false, breakLine: j < c.pairs.length - 1, paraSpaceAfter: 10, color: K.body } });
      });
      T(s, runs, { x: x + 0.3, y: Y0 + 1.25, w: w - 0.5, h: h - 1.4, fontSize: 14 });
    }
  });
  caption(s, "RFP 8항 \"예산 수준별 우선순위 및 제외 항목 제시\" 기준에 따라 범위를 정의했습니다.", 6.35);
  s.addNotes("PT 질문 11번(직접 운영 가능 범위) 대응: 온라인은 직접 운영, 조사·오프라인·유통 영업은 범위 외임을 명확히 합니다.");
}

// =========================================================
// Slide 4. 현황 진단 ① 매출과 리뷰
// =========================================================
pres.addSection({ title: "현황 진단" });
{
  const s = content("현황 진단", "03 현황 진단 ① 매출과 리뷰", "리뷰 부족이 광고 효율을 막는 구조입니다");
  const lw = 4.9, h = YB - Y0;
  card(s, X0, Y0, lw, h);
  s.addChart(pres.charts.DOUGHNUT, [{ name: "온라인 월 매출", labels: ["네이버 스마트스토어", "쿠팡"], values: [200, 100] }], Object.assign({
    x: X0 + 0.15, y: Y0 + 0.1, w: lw - 0.3, h: h - 0.2,
    holeSize: 58, chartColors: [K.primary, K.primarySoft],
    showTitle: true, title: "온라인 월 매출 구성 (만 원)", titleFontSize: 15, titleColor: K.white, titleBold: true,
    showValue: true, showPercent: false, showLabel: false, dataLabelColor: "191F28", dataLabelFontSize: 14, dataLabelFontBold: true,
    showLegend: true, legendPos: "b", legendFontSize: 13, legendColor: K.body,
    dataBorder: { pt: 1.5, color: K.navy }, objectName: "chart-sales-mix",
  }, chartFont));

  const rx = X0 + lw + GAP, rw = CW - lw - GAP;
  card(s, rx, Y0, rw, 1.85);
  T(s, "17개", { x: rx + 0.35, y: Y0 + 0.2, w: 2.6, h: 0.8, fontSize: 46, bold: true, color: K.primarySoft, valign: "middle", fit: "none" });
  T(s, "스마트스토어 상품당 최대 리뷰 수", { x: rx + 0.35, y: Y0 + 1.05, w: rw - 0.7, h: 0.32, fontSize: 15, bold: true, color: K.white });
  T(s, "경쟁 브랜드는 리뷰 수천~수만 건", { x: rx + 0.35, y: Y0 + 1.4, w: rw - 0.7, h: 0.3, fontSize: 14, color: K.muted });

  const iy = Y0 + 1.85 + GAP, ih = YB - iy;
  card(s, rx, iy, rw, ih, "primary");
  T(s, "시사점", { x: rx + 0.35, y: iy + 0.3, w: 2, h: 0.3, fontSize: 14, bold: true, color: K.primarySoft });
  T(s, "지금 광고비를 크게 늘리면 클릭은 늘지만 구매는 따라오지 않습니다. 광고보다 리뷰 확보가 먼저입니다.", { x: rx + 0.35, y: iy + 0.72, w: rw - 0.7, h: ih - 0.9, fontSize: 17, bold: true, color: K.white, lineSpacingMultiple: 1.3 });
  s.addNotes("매출 수치는 광고주 통화 확인 기준(네이버 월 200만 원, 쿠팡 월 100만 원).");
}

// =========================================================
// Slide 5. 현황 진단 ② 자사몰
// =========================================================
{
  const s = content("현황 진단", "03 현황 진단 ② 자사몰", "자사몰에서 확인한 세 가지 과제");
  const items = [
    ["1", "브랜드명 혼재", "사이트명이 \"양명 킹오브김 (해풀잎)\"이고, 상품명에 KING OF KIM · 해풀잎 · 사랑둥이 등이 섞여 있어 브랜드 검색으로 이어지기 어렵습니다."],
    ["2", "상품군 이원화", "일상용(도시락김 · 김자반) 1.8~3.1만 원, 선물세트 5.3~16.8만 원. 객단가 높은 선물세트가 광고 효율에 유리합니다."],
    ["3", "리뷰 · 채널 연결 부족", "상품 리뷰가 거의 없고, 하단 SNS 링크가 실제 계정과 연결되어 있지 않습니다."],
  ];
  const h = 3.25;
  items.forEach(([n, t, d], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, h);
    badge(s, x + 0.3, Y0 + 0.3, 0.5, n);
    T(s, t, { x: x + 0.3, y: Y0 + 0.95, w: W3 - 0.6, h: 0.42, fontSize: 19, bold: true, color: K.white });
    T(s, d, { x: x + 0.3, y: Y0 + 1.47, w: W3 - 0.55, h: h - 1.6, fontSize: 14, lineSpacingMultiple: 1.2 });
  });
  const by = Y0 + h + GAP;
  card(s, X0, by, CW, YB - by, "tint");
  T(s, [
    { text: "활용 가능 자산", options: { bold: true, color: K.primarySoft } },
    { text: "    첫 구매 3,000원 회원가입 쿠폰 운영 중 · 메타 픽셀 설치 완료 (추후 리타겟팅 기반)", options: { color: K.white } },
  ], { x: X0 + 0.3, y: by, w: CW - 0.6, h: YB - by, fontSize: 15, valign: "middle" });
  s.addNotes("ym-kr.com 자사몰 확인 기준.\n\n근거: 사이트 타이틀 '양명 킹오브김 ( 해풀잎 )' / KING OF KIM 도시락김 36봉 22,400원, 72봉 31,400원, 김자반 50g×5봉 18,600원 / 선물세트 사랑둥이 C1호 53,000원 ~ 효자둥이 특품 168,000원 / 첫 쇼핑 3,000원 할인 회원가입 쿠폰 배너 / 메타(페이스북) 픽셀 설치 확인 / 하단 SNS 링크가 instagram.com, youtube.com 등 메인 주소로만 연결.");
}

// =========================================================
// Slide 6. 원인 확인 방법
// =========================================================
{
  const s = content("현황 진단", "03 원인 확인 방법", "\"인지도 부족\"인지 운영 데이터로 구분합니다");
  const cx = { q: X0 + 1.0, m: X0 + 3.45, j: X0 + 7.0 };
  const cw = { q: 2.3, m: 3.35, j: 3.45 };
  T(s, "확인 질문", { x: cx.q, y: Y0, w: cw.q, h: 0.3, fontSize: 12, color: K.muted });
  T(s, "확인 지표", { x: cx.m, y: Y0, w: cw.m, h: 0.3, fontSize: 12, color: K.muted });
  T(s, "판단", { x: cx.j, y: Y0, w: cw.j, h: 0.3, fontSize: 12, color: K.muted });
  const rows = [
    ["1", "노출 문제?", "브랜드 검색량, 쇼핑검색 노출수", "검색량이 거의 없으면 인지도 문제"],
    ["2", "전환 문제?", "유입 대비 구매전환율", "유입은 있는데 구매가 없으면 리뷰 · 가격 · 상세페이지 문제"],
    ["3", "가격 문제?", "경쟁사 동일 규격 가격 · 리뷰 수, 세트별 전환율", "가격 경쟁력과 세트 구성 조정 필요 여부"],
  ];
  const rh = 1.0;
  rows.forEach(([n, q, m, j], i) => {
    const y = Y0 + 0.38 + i * (rh + GAP);
    card(s, X0, y, CW, rh);
    badge(s, X0 + 0.3, y + (rh - 0.48) / 2, 0.48, n);
    T(s, q, { x: cx.q, y, w: cw.q, h: rh, fontSize: 18, bold: true, color: K.white, valign: "middle" });
    T(s, m, { x: cx.m, y, w: cw.m, h: rh, fontSize: 14, valign: "middle" });
    T(s, j, { x: cx.j, y, w: cw.j, h: rh, fontSize: 14, bold: true, color: K.primarySoft, valign: "middle" });
  });
  caption(s, "정량 소비자 조사는 범위 외입니다. 위 데이터는 매월 리포트로 제공해 양명의 추가 조사 · 의사결정 근거로 활용할 수 있게 합니다.", 6.25);
  s.addNotes("PT 질문 1·2번 대응. 인지도 부족을 단정하지 않고 노출·전환·가격 중 어디서 막히는지 데이터로 구분합니다.");
}

// =========================================================
// Slide 7. 전략 방향
// =========================================================
pres.addSection({ title: "전략" });
{
  const s = content("전략", "04 전략 방향", "리뷰 → 검색 → 구매 → 재구매 순서로 갑니다");
  const flow = [["리뷰", "체험단으로 구매 결정 근거 확보"], ["검색", "쇼핑검색 · 브랜드 키워드 노출"], ["구매", "리뷰 갖춘 대표 상품으로 전환"], ["재구매", "알림받기 · 재구매 쿠폰 권장"]];
  const step = (CW + GAP) / 4, d = 1.3, cy = Y0 + 0.05;
  flow.forEach(([t, desc], i) => {
    const x = X0 + i * step;
    badge(s, x, cy, d, t, i === 2 ? K.white : K.primary, i === 2 ? K.primary : K.white, 20);
    T(s, desc, { x, y: cy + d + 0.2, w: step - 0.4, h: 0.6, fontSize: 15, color: K.body });
    if (i < 3) {
      s.addShape(pres.shapes.LINE, { x: x + d + 0.2, y: cy + d / 2, w: step - d - 0.4, h: 0, line: { color: K.primarySoft, width: 2.5, endArrowType: "triangle" }, objectName: `arrow-${i}` });
    }
  });
  const by = 4.5, bh = YB - by;
  card(s, X0, by, CW, bh, "tint");
  T(s, "브랜드명 · 패키지는 그대로 유지합니다", { x: X0 + 0.35, y: by + 0.3, w: CW - 0.7, h: 0.4, fontSize: 19, bold: true, color: K.white });
  T(s, "대신 상품명 · 광고 문구 · 체험단 콘텐츠에서 \"KING OF KIM\"을 일관되게 노출해, 구매 경험이 브랜드 검색으로 이어지게 만듭니다. 리뷰가 없는 상태의 광고는 밑 빠진 독에 물 붓기와 같아서, 리뷰부터 채우고 광고를 늘립니다.", { x: X0 + 0.35, y: by + 0.82, w: CW - 0.7, h: bh - 0.95, fontSize: 15, color: K.white, lineSpacingMultiple: 1.3 });
  s.addNotes("PT 질문 3·4·5번 대응: 브랜드 변경 없이 인지도 확대, 최초 구매 유도, 재구매 전환.");
}

// =========================================================
// Slide 8. 채널 전략
// =========================================================
{
  const s = content("전략", "04 채널 전략", "네이버 쇼핑검색을 주력으로, 쿠팡은 단계 확대");
  const ch = [
    ["50%", "주력", "네이버 쇼핑검색", ["가격 · 리뷰 비교 구매 품목이라 구매와 가장 가까운 지면", "도시락김 · 곱창김 · 김자반 + 용도 키워드", "경쟁이 덜한 세부 키워드부터 시작"]],
    ["17%", "보조", "네이버 파워링크", ["\"킹오브김\" 브랜드 키워드 방어", "선물세트 키워드는 자사몰 선물세트 페이지로 연결", "고단가 상품 판매 담당"]],
    ["33%", "단계 확대", "쿠팡 광고", ["대표 상품(대용량 도시락김 · 김자반) 중심", "리뷰 축적 속도에 맞춰 증액", "3개월차 효율 비교 후 비중 조정"]],
  ];
  const h = 4.0;
  ch.forEach(([p, tag, name, bl], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, h);
    T(s, p, { x: x + 0.3, y: Y0 + 0.25, w: 1.7, h: 0.72, fontSize: 38, bold: true, color: K.primarySoft, valign: "middle", fit: "none" });
    const tw = tag.length > 2 ? 1.15 : 0.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + W3 - 0.3 - tw, y: Y0 + 0.41, w: tw, h: 0.4, rectRadius: 0.2, fill: { color: K.white }, line: { type: "none" }, objectName: `tag-${i}` });
    T(s, tag, { x: x + W3 - 0.3 - tw, y: Y0 + 0.41, w: tw, h: 0.4, fontSize: 13, bold: true, color: K.primary, align: "center", valign: "middle" });
    T(s, name, { x: x + 0.3, y: Y0 + 1.15, w: W3 - 0.6, h: 0.42, fontSize: 19, bold: true, color: K.white });
    T(s, bullets(bl, { paraSpaceAfter: 7 }), { x: x + 0.3, y: Y0 + 1.72, w: W3 - 0.5, h: h - 1.85, fontSize: 14 });
  });
  caption(s, "비중은 전체 광고비 기준 시작값(네이버:쿠팡 = 현재 매출 비중 2:1)이며, 3개월차부터 매체별 ROAS로 재배분합니다.", 6.3);
  s.addNotes("PT 질문 6번 대응: 스마트스토어·쿠팡 판매량 증대 방법.");
}

// =========================================================
// Slide 9. 체험단 · 리뷰 확보
// =========================================================
{
  const s = content("전략", "04 체험단 · 리뷰 확보", "1~2개월차, 체험단으로 리뷰 기반부터 만듭니다");
  const h = YB - Y0, lw = 3.1;
  card(s, X0, Y0, lw, h, "primary");
  T(s, "대표 상품 리뷰 목표", { x: X0 + 0.35, y: Y0 + 0.35, w: lw - 0.7, h: 0.35, fontSize: 15, bold: true, color: K.primarySoft });
  [["50건+", "2개월차까지", 0.95], ["150건+", "6개월차까지", 2.55]].forEach(([n, l, oy]) => {
    T(s, n, { x: X0 + 0.35, y: Y0 + oy, w: lw - 0.5, h: 0.8, fontSize: 44, bold: true, color: K.white, valign: "middle", fit: "none" });
    T(s, l, { x: X0 + 0.35, y: Y0 + oy + 0.85, w: lw - 0.7, h: 0.35, fontSize: 15, color: K.white });
  });
  const rx = X0 + lw + GAP, rw = CW - lw - GAP;
  const rows = [
    ["1", "역할 분담", "제품은 양명이 제공하고, 체험단 모집 · 운영 · 콘텐츠 관리는 NICECORE가 담당합니다."],
    ["2", "확보 채널", "스마트스토어 리뷰와 네이버 블로그 후기 중심. 블로그 후기는 \"킹오브김\" 검색 결과에도 노출됩니다."],
    ["3", "쿠팡 운영", "쿠팡은 외부 체험단 리뷰에 제한이 있어, 플랫폼 정책에 맞는 방식(쿠팡 체험단 프로그램 등)을 별도 검토합니다."],
  ];
  const rh = (h - 2 * GAP) / 3;
  rows.forEach(([n, t, d], i) => {
    const y = Y0 + i * (rh + GAP);
    card(s, rx, y, rw, rh);
    badge(s, rx + 0.3, y + 0.2, 0.45, n);
    T(s, t, { x: rx + 0.95, y: y + 0.18, w: rw - 1.25, h: 0.48, fontSize: 18, bold: true, color: K.white, valign: "middle" });
    T(s, d, { x: rx + 0.95, y: y + 0.66, w: rw - 1.2, h: rh - 0.72, fontSize: 14 });
  });
  s.addNotes("체험단 운영비는 별도 견적. 제품은 광고주 제공 조건.");
}

// =========================================================
// Slide 10. 시즌 전략
// =========================================================
{
  const s = content("전략", "04 시즌 전략", "2027년 설, 선물세트로 첫 성공사례를 만듭니다");
  const tl = [
    ["12월", "준비", "선물세트 키워드 · 광고 소재 세팅, 자사몰 선물세트 페이지 점검"],
    ["1월", "집중", "파워링크 · 쇼핑검색으로 선물세트 집중 노출, 광고비 증액 구간"],
    ["2월 초", "피크 · 분석", "명절 직전 피크 운영, 종료 후 상품 · 키워드별 성과 분석"],
  ];
  const h = 2.15;
  tl.forEach(([m, st, d], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, h);
    T(s, m, { x: x + 0.3, y: Y0 + 0.22, w: 1.8, h: 0.48, fontSize: 24, bold: true, color: K.primarySoft, valign: "middle" });
    T(s, st, { x: x + 0.3, y: Y0 + 0.75, w: W3 - 0.6, h: 0.35, fontSize: 17, bold: true, color: K.white });
    T(s, d, { x: x + 0.3, y: Y0 + 1.18, w: W3 - 0.55, h: h - 1.25, fontSize: 14 });
    if (i < 2) {
      const ax = x + W3 + GAP / 2 - 0.22;
      s.addShape(pres.shapes.OVAL, { x: ax, y: Y0 + h / 2 - 0.22, w: 0.44, h: 0.44, fill: { color: K.primary }, line: { color: K.navy, width: 1.5 }, objectName: `tl-arrow-${i}` });
      T(s, "›", { x: ax, y: Y0 + h / 2 - 0.25, w: 0.44, h: 0.44, fontSize: 20, bold: true, color: K.white, align: "center", valign: "middle" });
    }
  });
  const by = Y0 + h + GAP, bh = YB - by, bw = (CW - GAP) / 2;
  [["5~16만 원대", "선물세트 가격대. 같은 광고비로 일상용 상품보다 큰 매출을 기대할 수 있습니다."],
   ["Sales Case", "명절 실적은 6개월 성과 자료의 핵심 사례로, 유통 MD · 기업 특판 제안에 활용됩니다."]].forEach(([n, d], i) => {
    const x = X0 + i * (bw + GAP);
    card(s, x, by, bw, bh, "tint");
    T(s, n, { x: x + 0.35, y: by + 0.28, w: bw - 0.7, h: 0.7, fontSize: 34, bold: true, color: K.white, valign: "middle", fit: "none" });
    T(s, d, { x: x + 0.35, y: by + 1.08, w: bw - 0.7, h: bh - 1.2, fontSize: 15, color: K.white });
  });
  s.addNotes("Pilot 기간을 2026년 11월~2027년 4월로 가정. 설 연휴는 2027년 2월 초.");
}

// =========================================================
// Slide 11. 6개월 로드맵
// =========================================================
pres.addSection({ title: "실행 계획" });
{
  const s = content("실행 계획", "05 6개월 로드맵", "매출 반응을 보며 단계적으로 증액합니다");
  const rm = [
    ["11월", "기반 구축", "계정 세팅 · Baseline 측정 · 체험단 1차", "참고", "100%"],
    ["12월", "리뷰 확보", "체험단 2차 · 광고 소재 테스트", "참고", "150%"],
    ["1월", "확대 ①", "설 선물세트 집중 · 광고 증액", "참고", "200%"],
    ["2월", "확대 ②", "명절 성과 분석 · 효율 채널 재배분", "목표", "300%"],
    ["3월", "집중", "효율 상품 · 채널에 예산 집중", "목표", "300%"],
    ["4월", "정리", "성과 유지 · 6개월 성과 자료 작성", "목표", "300%"],
  ];
  const g = 0.2, cw = (CW - 5 * g) / 6, h = 2.5;
  rm.forEach(([m, st, act, kind, roas], i) => {
    const x = X0 + i * (cw + g);
    card(s, x, Y0, cw, h);
    T(s, m, { x: x + 0.18, y: Y0 + 0.2, w: cw - 0.36, h: 0.42, fontSize: 20, bold: true, color: K.primarySoft, valign: "middle" });
    T(s, st, { x: x + 0.18, y: Y0 + 0.68, w: cw - 0.36, h: 0.35, fontSize: 16, bold: true, color: K.white });
    T(s, bullets(act.split(" · "), { bullet: { indent: 10 }, paraSpaceAfter: 4 }), { x: x + 0.16, y: Y0 + 1.12, w: cw - 0.26, h: h - 1.2, fontSize: 14 });
    const goal = kind === "목표";
    const cy = Y0 + h + 0.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: 0.68, rectRadius: 0.15, fill: goal ? { color: K.primary } : { color: K.white, transparency: 85 }, line: goal ? { type: "none" } : { color: K.primarySoft, width: 1 }, objectName: `roas-chip-${i}` });
    T(s, [
      { text: `${kind} ROAS`, options: { fontSize: 12, bold: false, breakLine: true } },
      { text: roas, options: { fontSize: 17, bold: true } },
    ], { x, y: cy, w: cw, h: 0.68, color: goal ? K.white : K.primarySoft, align: "center", valign: "middle", lineSpacingMultiple: 1.0 });
  });
  const by = 5.65;
  card(s, X0, by, CW, YB - by, "outline");
  T(s, [
    { text: "증액 조건", options: { bold: true, color: K.primarySoft } },
    { text: "    1~2개월차는 리뷰 목표 달성 시 다음 단계로. 3개월차부터는 전월 ROAS 기준 달성 시 증액, 미달 시 동결 후 키워드 · 상품 정리", options: { color: K.white } },
  ], { x: X0 + 0.3, y: by, w: CW - 0.6, h: YB - by, fontSize: 14, valign: "middle" });
  s.addNotes("1~3개월차는 리뷰를 쌓는 투자 구간이라 ROAS는 참고치. 4개월차부터 목표 ROAS 300% 적용. ROAS = 광고 매출 ÷ 광고비(수수료 제외).");
}

// =========================================================
// Slide 12. 예산 옵션
// =========================================================
{
  const s = content("실행 계획", "05 예산 옵션", "최대 도달 예산에 따라 3가지 안을 제안합니다");
  const months = ["1개월", "2개월", "3개월", "4개월", "5개월", "6개월"];
  const lw = 6.6, h = 4.05;
  card(s, X0, Y0, lw, h);
  s.addChart(pres.charts.BAR, [
    { name: "A안", labels: months, values: [150, 200, 300, 400, 500, 500] },
    { name: "B안 (권장)", labels: months, values: [200, 300, 450, 600, 750, 750] },
    { name: "C안", labels: months, values: [250, 400, 600, 800, 1000, 1000] },
  ], Object.assign({
    x: X0 + 0.1, y: Y0 + 0.08, w: lw - 0.2, h: h - 0.16, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: ["8EA1B9", K.primary, K.primarySoft],
    showTitle: true, title: "월별 광고비 (만 원, 네이버 + 쿠팡)", titleFontSize: 14, titleColor: K.white, titleBold: true,
    showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: K.body,
    catAxisLabelColor: K.body, catAxisLabelFontSize: 12, catAxisLineShow: true, catAxisLineColor: K.grid,
    valAxisLabelColor: K.muted, valAxisLabelFontSize: 11, valAxisLabelFormatCode: "#,##0", valAxisLineShow: false,
    valAxisMaxVal: 1000, valAxisMajorUnit: 250,
    valGridLine: { color: K.grid, size: 0.75 }, catGridLine: { style: "none" },
    objectName: "chart-budget-options",
  }, chartFont));
  const rx = X0 + lw + GAP, rw = CW - lw - GAP, ch = (h - 2 * GAP) / 3;
  [["A안", "최대 월 500만 원", "2,050만 원", false], ["B안 · 권장", "최대 월 750만 원", "3,050만 원", true], ["C안", "최대 월 1,000만 원", "4,050만 원", false]].forEach(([t, sub, tot, hl], i) => {
    const y = Y0 + i * (ch + GAP);
    card(s, rx, y, rw, ch, hl ? "primary" : "glass");
    T(s, t, { x: rx + 0.3, y: y + 0.22, w: 1.7, h: 0.4, fontSize: 17, bold: true, color: K.white, valign: "middle" });
    T(s, tot, { x: rx + rw - 0.3 - 1.75, y: y + 0.22, w: 1.75, h: 0.4, fontSize: 20, bold: true, color: hl ? K.white : K.primarySoft, align: "right", valign: "middle" });
    T(s, sub, { x: rx + 0.3, y: y + 0.7, w: 1.9, h: 0.3, fontSize: 14, color: hl ? K.white : K.body });
    T(s, "6개월 합계", { x: rx + rw - 0.3 - 1.5, y: y + 0.72, w: 1.5, h: 0.28, fontSize: 12, color: hl ? K.primarySoft : K.muted, align: "right" });
  });
  caption(s, "6개월 광고비 합계 기준 · VAT 별도 | 대행수수료: 쿠팡 광고비의 15% (네이버 없음) | 체험단 운영비 별도 견적 | 상세페이지 개선 선택 시 별도", 6.3);
  s.addNotes("PT 질문 10번 대응. 매체별 상세 배분은 부록 참조. 컨설팅·조사·오프라인 비용은 범위 외로 해당 없음.");
}

// ---- table helpers ----
const tb = { type: "solid", pt: 0.75, color: K.tBorder };
function cell(text, o = {}) {
  return { text, options: Object.assign({ border: [tb, tb, tb, tb] }, o) };
}
const head = (h, o = {}) => cell(h, Object.assign({ bold: true, color: K.white, fill: { color: K.primary } }, o));

// =========================================================
// Slide 13. 목표 ROAS 300% 기준 예상 매출
// =========================================================
{
  const s = content("실행 계획", "05 목표 ROAS 300% 기준 예상 매출", "4개월차부터 광고비의 3배 매출을 목표로 합니다");
  [["A안", "1,500만 원", false], ["B안 · 권장", "2,250만 원", true], ["C안", "3,000만 원", false]].forEach(([o, v, hl], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, 1.4, hl ? "primary" : "glass");
    T(s, o, { x: x + 0.3, y: Y0 + 0.15, w: W3 - 0.6, h: 0.3, fontSize: 14, bold: true, color: K.white });
    T(s, v, { x: x + 0.3, y: Y0 + 0.45, w: W3 - 0.6, h: 0.55, fontSize: 30, bold: true, color: hl ? K.white : K.primarySoft, valign: "middle", fit: "none" });
    T(s, "6개월차 월 광고 매출", { x: x + 0.3, y: Y0 + 1.03, w: W3 - 0.6, h: 0.28, fontSize: 13, color: hl ? K.primarySoft : K.muted });
  });
  const hd = ["구분", "1개월", "2개월", "3개월", "4개월", "5개월", "6개월", "6개월 합계", "평균 ROAS"];
  const data = [
    ["목표 ROAS", "100%", "150%", "200%", "300%", "300%", "300%", "-", "-"],
    ["A안", "150", "300", "600", "1,200", "1,500", "1,500", "5,250", "256%"],
    ["B안 (권장)", "200", "450", "900", "1,800", "2,250", "2,250", "7,850", "257%"],
    ["C안", "250", "600", "1,200", "2,400", "3,000", "3,000", "10,450", "258%"],
  ];
  const rows = [hd.map((h, j) => head(h, { align: j === 0 ? "left" : "center", fontSize: 13 }))];
  data.forEach((r, i) => {
    const isGoal = i === 0, isB = i === 2;
    const fill = isGoal ? K.tTot : isB ? K.tHi : K.tRow;
    rows.push(r.map((v, j) => cell(v, {
      fill: { color: fill }, bold: isB || j === 0 || (j === 8 && !isGoal),
      color: j === 8 && !isGoal ? K.primarySoft : (j === 0 || isB ? K.white : K.body),
      align: j === 0 ? "left" : "center",
    })));
  });
  const fc = 1.35, sc = 1.3, oc = (CW - fc - 2 * sc) / 6;
  s.addTable(rows, { x: X0, y: 3.75, w: CW, colW: [fc, ...Array(6).fill(oc), sc, sc], rowH: 0.42, fontSize: 14, valign: "middle", margin: [0, 0.12, 0, 0.12], objectName: "table-roas" });
  T(s, "단위: 만 원. 광고 기여 매출 = 광고비 × 목표 ROAS. 현재 자연 매출(월 약 300만 원)은 별도이며, ROAS는 대행수수료를 제외한 광고비 기준입니다. 1~3개월차 ROAS는 리뷰 확보 단계의 참고치입니다.", { x: X0, y: 5.98, w: CW, h: 0.45, fontSize: 12, color: K.muted });
  s.addNotes("C안 6개월차 월 광고 매출 3,000만 원은 현재 온라인 매출의 약 10배로 공격적인 목표. 3개월차 점검 결과에 따라 옵션 간 이동 가능함을 함께 설명.\n\n계산: 광고 매출 = 월 광고비 × 해당 월 ROAS. 평균 ROAS = 6개월 광고 매출 합계 ÷ 6개월 광고비 합계 (A 5,250÷2,050, B 7,850÷3,050, C 10,450÷4,050, 반올림).");
}

// =========================================================
// Slide 14. KPI와 성과 측정
// =========================================================
{
  const s = content("실행 계획", "05 KPI와 성과 측정", "온라인 지표로 매월 측정하고 보고합니다");
  const kpi = [
    ["리뷰", "대표 상품 리뷰 수", "2개월차 50건+, 6개월차 150건+"],
    ["브랜드", "\"킹오브김\" 브랜드 검색량", "Baseline 대비 월별 증가율"],
    ["유입", "스토어 방문자 수", "월별 추이"],
    ["전환", "구매전환율 · 신규 구매자 수", "월별 추이, 매체별 비교"],
    ["매출", "온라인 매출 · 광고 ROAS", null],
    ["재구매", "재구매 고객 비율", "4개월차 이후 측정"],
  ];
  const rows = [["영역", "KPI", "목표 · 측정 방법"].map((h) => head(h))];
  kpi.forEach(([a, b, c], i) => {
    const fill = i % 2 ? K.tAlt : K.tRow;
    const third = c === null
      ? [{ text: "Baseline 월 300만 원 대비, ", options: {} }, { text: "4개월차 이후 ROAS 300%", options: { bold: true, color: K.primarySoft } }]
      : c;
    rows.push([
      cell(a, { fill: { color: fill }, bold: true, color: K.white }),
      cell(b, { fill: { color: fill }, color: K.white }),
      cell(third, { fill: { color: fill }, color: K.body }),
    ]);
  });
  s.addTable(rows, { x: X0, y: Y0, w: CW, colW: [1.4, 3.2, CW - 4.6], rowH: 0.46, fontSize: 14, valign: "middle", margin: [0, 0.2, 0, 0.2], objectName: "table-kpi" });
  const by = 5.65;
  card(s, X0, by, CW, YB - by, "tint");
  T(s, [
    { text: "월간 리포트", options: { bold: true, color: K.primarySoft } },
    { text: "    지표 추이, 매체별 효율, 예산 재배분 내역과 다음 달 운영 계획을 함께 공유합니다.", options: { color: K.white } },
  ], { x: X0 + 0.3, y: by, w: CW - 0.6, h: YB - by, fontSize: 15, valign: "middle" });
  s.addNotes("KPI는 온라인 지표로 한정. 오프라인 판매·신규 입점·조달 매출은 범위 외.");
}

// =========================================================
// Slide 15. 기대 효과와 리스크
// =========================================================
pres.addSection({ title: "기대 효과" });
{
  const s = content("기대 효과", "06 기대 효과와 리스크", "온라인 성과가 유통 확대의 근거가 됩니다");
  const lw = 4.3, h = 3.95;
  card(s, X0, Y0, lw, h, "primary");
  T(s, "6개월 후 결과물", { x: X0 + 0.35, y: Y0 + 0.35, w: lw - 0.7, h: 0.45, fontSize: 21, bold: true, color: K.white });
  ["리뷰가 갖춰진 대표 상품", "측정 가능한 브랜드 검색량", "MD · 조달 담당자에게 보여줄 온라인 판매 성과 자료"].forEach((t, i) => {
    const y = Y0 + 1.1 + i * 0.88;
    badge(s, X0 + 0.35, y, 0.42, "✓", K.white, K.primary, 15);
    T(s, t, { x: X0 + 0.95, y: y - 0.03, w: lw - 1.25, h: 0.75, fontSize: 16, bold: true, color: K.white });
  });
  const rx = X0 + lw + GAP, rw = CW - lw - GAP, rh = (h - 2 * GAP) / 3;
  const half = (rw - 0.6) / 2;
  [["리뷰 확보 지연", "증액 시점을 늦추고 체험단 추가 운영"], ["낮은 객단가로 효율 저조", "세트 · 대용량 위주로 전환, 묶음 구성 제안"], ["쿠팡 효율 저조", "네이버로 예산 이동"]].forEach(([r, a], i) => {
    const y = Y0 + i * (rh + GAP);
    card(s, rx, y, rw, rh);
    T(s, "리스크", { x: rx + 0.3, y: y + 0.17, w: 1.2, h: 0.28, fontSize: 13, bold: true, color: K.danger });
    T(s, r, { x: rx + 0.3, y: y + 0.48, w: half - 0.15, h: rh - 0.55, fontSize: 16, bold: true, color: K.white });
    T(s, "대응", { x: rx + 0.3 + half + 0.15, y: y + 0.17, w: 1.2, h: 0.28, fontSize: 13, bold: true, color: K.primarySoft });
    T(s, a, { x: rx + 0.3 + half + 0.15, y: y + 0.48, w: half - 0.15, h: rh - 0.55, fontSize: 15, color: K.body });
  });
  caption(s, "본 제안의 KPI는 온라인 지표에 한정되며, 오프라인 판매 · 신규 입점 · 조달 매출은 범위 외입니다.", 6.3);
  s.addNotes("PT 질문 7·8번 대응: 오프라인 입점 확대의 근거를 온라인 판매 실적으로 먼저 만든다는 논리.");
}

// =========================================================
// Slide 16. 부록
// =========================================================
pres.addSection({ title: "부록" });
{
  const s = content("부록", "부록", "옵션별 · 매체별 월 광고비 상세");
  const hd = ["옵션", "매체", "1개월", "2개월", "3개월", "4개월", "5개월", "6개월", "합계"];
  const groups = [
    ["A안", [["네이버 쇼핑검색", 80, 100, 150, 200, 250, 250, 1030], ["네이버 파워링크", 20, 30, 50, 70, 80, 80, 330], ["쿠팡 광고", 50, 70, 100, 130, 170, 170, 690], ["광고비 합계", 150, 200, 300, 400, 500, 500, 2050]]],
    ["B안\n(권장)", [["네이버 쇼핑검색", 100, 150, 230, 300, 380, 380, 1540], ["네이버 파워링크", 30, 50, 70, 100, 120, 120, 490], ["쿠팡 광고", 70, 100, 150, 200, 250, 250, 1020], ["광고비 합계", 200, 300, 450, 600, 750, 750, 3050]]],
    ["C안", [["네이버 쇼핑검색", 130, 200, 300, 400, 500, 500, 2030], ["네이버 파워링크", 40, 70, 100, 130, 170, 170, 680], ["쿠팡 광고", 80, 130, 200, 270, 330, 330, 1340], ["광고비 합계", 250, 400, 600, 800, 1000, 1000, 4050]]],
  ];
  const fmt = (n) => n.toLocaleString("en-US");
  const rows = [hd.map((h, j) => head(h, { align: j < 2 ? "left" : "center" }))];
  groups.forEach(([opt, rs], gi) => {
    const isB = gi === 1;
    rs.forEach((r, ri) => {
      const total = ri === 3;
      const fill = total ? K.tTot : isB ? K.tHi : K.tRow;
      const row = [];
      if (ri === 0) row.push(cell(opt, { rowspan: 4, fill: { color: isB ? K.tHi : K.tRow }, bold: true, color: K.white, valign: "middle" }));
      r.forEach((v, j) => {
        row.push(cell(typeof v === "number" ? fmt(v) : v, {
          fill: { color: fill }, bold: total || (isB && j === 7),
          color: total && j === 7 ? K.primarySoft : total || isB ? K.white : K.body,
          align: j === 0 ? "left" : "center",
        }));
      });
      rows.push(row);
    });
  });
  const c0 = 0.95, c1 = 1.9, cl = 1.3, cm = (CW - c0 - c1 - cl) / 6;
  s.addTable(rows, { x: X0, y: Y0, w: CW, colW: [c0, c1, ...Array(6).fill(cm), cl], rowH: 0.3, fontSize: 11, valign: "middle", margin: [0, 0.1, 0, 0.1], objectName: "table-appendix" });
  caption(s, "단위: 만 원, VAT 별도. 쿠팡 대행수수료(광고비의 15%) 6개월 합계: A안 103.5 · B안 153 · C안 201", 6.25);
  s.addNotes("각 월의 세 매체 합 = 해당 옵션 월 광고비. 쿠팡 수수료 = 쿠팡 광고 합계 × 0.15 (A 690×0.15=103.5, B 1,020×0.15=153, C 1,340×0.15=201).");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
