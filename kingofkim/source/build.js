const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const A = (f) => path.join(__dirname, f);

// ---- Theme (brief §2-2) ----
const HEX = {
  primary: "3182F6", primaryDark: "1B64DA", primaryLight: "E8F3FF", primarySoft: "B0D4FF",
  strong: "191F28", body: "4E5968", muted: "8B95A1", surface: "F2F4F6", border: "E5E8EB",
  danger: "F04452", white: "FFFFFF",
};
const THEME = {
  name: "NICECORE KING OF KIM",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: HEX.strong, lt1: HEX.white, dk2: HEX.body, lt2: HEX.surface,
    accent1: HEX.primary, accent2: HEX.primaryDark, accent3: HEX.primaryLight, accent4: HEX.primarySoft,
    accent5: HEX.muted, accent6: HEX.danger, hlink: HEX.primary, folHlink: HEX.primaryDark,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "NICECORE";
pres.company = "NICECORE";
pres.title = "KING OF KIM 온라인 판매 활성화 제안서";

const C = pres.SchemeColor;
const K = {
  primary: C.accent1, primaryDark: C.accent2, primaryLight: C.accent3, primarySoft: C.accent4,
  strong: C.text1, body: C.text2, muted: C.accent5, surface: C.background2, white: C.background1,
  danger: C.accent6, border: HEX.border,
};

const X0 = 0.6, CW = 12.133, GAP = 0.3, Y0 = 1.85;
const COL3 = [0.6, 4.75, 8.9], W3 = 3.85;
const R = 0.12;

// ---- Layouts ----
pres.defineSlideMaster({
  title: "COVER",
  background: { color: K.primary },
  objects: [
    { image: { x: 0, y: 0, w: 13.333, h: 7.5, path: A("cover_bg.png") } },
    { image: { x: 11.03, y: 6.75, w: 1.7, h: 0.25, path: A("logo_white.png") } },
    { placeholder: { options: { name: "title", type: "title", x: 1.0, y: 2.3, w: 10, h: 1.7, fontSize: 44, bold: true, color: K.white, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: K.white },
  margin: [0.45, 0.6, 0.5, 0.6],
  objects: [
    { placeholder: { options: { name: "eyebrow", type: "body", x: X0, y: 0.45, w: CW, h: 0.3, fontSize: 13, bold: true, color: K.primary, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 0.82, w: CW, h: 0.6, fontSize: 28, bold: true, color: K.strong, align: "left", valign: "top", margin: 0 }, text: "" } },
    { text: { text: "NICECORE | KING OF KIM 온라인 판매 활성화 제안", options: { x: X0, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: K.muted, margin: 0, valign: "middle" } } },
    { image: { x: 11.37, y: 7.05, w: 1.365, h: 0.2, path: A("logo_gray.png") } },
  ],
  slideNumber: { x: 10.55, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: HEX.muted, align: "right", margin: 0 },
});

// ---- Helpers ----
let objN = 0;
const glue = (t) => t.replace(/ · /g, "\u00A0·\u00A0").replace(/만 원/g, "만\u00A0원");
function T(s, text, o) {
  text = typeof text === "string" ? glue(text) : text.map((r) => Object.assign({}, r, { text: glue(r.text) }));
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: K.body, valign: "top", lineSpacingMultiple: 1.15, objectName: `text-${++objN}` }, o));
}
function card(s, x, y, w, h, fill, opts = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({
    x, y, w, h, rectRadius: R, fill: { color: fill },
    line: opts.outline ? { color: K.border, width: 1 } : { type: "none" },
    objectName: opts.name || `card-${++objN}`,
  }, {}));
}
function badge(s, x, y, d, label, fill = K.primary, color = K.white, size = 16) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: `badge-${++objN}` });
  s.addText(label, { isTextBox: true, x, y, w: d, h: d, margin: 0, align: "center", valign: "middle", fontSize: size, bold: true, color, objectName: `badge-text-${objN}` });
}
function content(section, eyebrow, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(items, o = {}) {
  return items.map((t, i) => ({
    text: t,
    options: Object.assign({ bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 6 }, o),
  }));
}
const chartFont = {
  titleFontFace: "+mn-lt", legendFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
};

// =========================================================
// Slide 1. 표지
// =========================================================
pres.addSection({ title: "표지" });
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  T(s, "제안서 · 2026.10", { x: 1.0, y: 1.75, w: 6, h: 0.35, fontSize: 16, bold: true, color: K.primarySoft });
  s.addText([
    { text: "KING OF KIM", options: { breakLine: true } },
    { text: "온라인 판매 활성화 6개월 Pilot" },
  ], { placeholder: "title", lineSpacingMultiple: 1.1 });
  T(s, "네이버 스마트스토어 · 쿠팡 퍼포먼스 마케팅 및 리뷰 기반 구축", { x: 1.0, y: 4.25, w: 10, h: 0.45, fontSize: 20, color: K.white });
  T(s, "주식회사 양명 귀중 | NICECORE", { x: 1.0, y: 6.7, w: 6, h: 0.35, fontSize: 14, color: K.white, valign: "middle" });
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
    card(s, x, Y0, W3, 1.45, K.surface);
    T(s, n, { x: x + 0.35, y: Y0 + 0.25, w: W3 - 0.7, h: 0.6, fontSize: 36, bold: true, color: K.primary, valign: "middle", fit: "none" });
    T(s, l, { x: x + 0.35, y: Y0 + 0.92, w: W3 - 0.7, h: 0.32, fontSize: 14 });
  });
  const steps = [
    ["1", "리뷰 확보", ["1~2개월차 체험단 집중", "광고는 최소 수준으로 테스트"]],
    ["2", "구매 전환", ["3개월차부터 네이버 쇼핑검색·쿠팡 광고로 구매 확대"]],
    ["3", "단계 증액", ["목표 ROAS 달성 시에만 다음 단계로 광고비 증액"]],
  ];
  const sy = 3.6;
  steps.forEach(([n, t, d], i) => {
    const x = COL3[i];
    card(s, x, sy, W3, 1.85, K.white, { outline: true });
    badge(s, x + 0.3, sy + 0.3, 0.5, n);
    T(s, t, { x: x + 0.95, y: sy + 0.3, w: W3 - 1.25, h: 0.5, fontSize: 19, bold: true, color: K.strong, valign: "middle" });
    T(s, d.map((line, j) => ({ text: line, options: { breakLine: j < d.length - 1 } })), { x: x + 0.3, y: sy + 1.0, w: W3 - 0.6, h: 0.7, fontSize: 14 });
  });
  card(s, X0, 5.75, CW, 0.8, K.primaryLight);
  T(s, "6개월 성과(판매량 · 리뷰 · 검색량)는 유통 MD 제안과 조달 홍보에 쓸 수 있는 성과 자료로 정리해 드립니다.", { x: X0 + 0.35, y: 5.75, w: CW - 0.7, h: 0.8, fontSize: 15, bold: true, color: K.primaryDark, valign: "middle" });
  s.addNotes("현재 매출 규모에서 예산을 여러 채널로 나누면 어느 채널에서도 검증 가능한 성과가 나오기 어렵다는 점을 먼저 설명합니다.");
}

// =========================================================
// Slide 3. 제안 범위
// =========================================================
{
  const s = content("제안 개요", "02 제안 범위", "예산이 성과로 이어지는 온라인 과업에 집중합니다");
  const h = 4.3;
  const cards = [
    { title: "직접 수행", sub: "NICECORE", fill: K.primaryLight, tColor: K.primary, items: [
      "네이버 쇼핑검색 · 파워링크 운영", "쿠팡 광고 운영", "체험단 기획 · 운영 (제품은 양명 제공)",
      "운영 데이터 진단 · 월간 리포트 · 예산 재배분", "6개월 성과 자료 정리 (MD · 조달 제안용)"] },
    { title: "별도 견적", sub: "선택 시 협의", fill: K.surface, tColor: K.strong, items: ["상세페이지 개선"] },
    { title: "범위 외", sub: "전문 업체 또는 양명", fill: K.surface, tColor: K.strong, pairs: [
      ["소비자 조사(설문 · FGI)", "리서치 전문사"], ["시식 · POP · 오프라인 광고", "홍보 · 오프라인 대행사"], ["유통 MD 영업 · 조달청 영업", "양명 직접"]] },
  ];
  cards.forEach((c, i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, h, c.fill);
    T(s, c.title, { x: x + 0.35, y: Y0 + 0.35, w: W3 - 0.7, h: 0.45, fontSize: 22, bold: true, color: c.tColor });
    T(s, c.sub, { x: x + 0.35, y: Y0 + 0.85, w: W3 - 0.7, h: 0.3, fontSize: 14, color: K.muted });
    if (c.items) {
      T(s, bullets(c.items), { x: x + 0.35, y: Y0 + 1.4, w: W3 - 0.6, h: h - 1.6, fontSize: 14, color: K.strong });
    } else {
      const runs = [];
      c.pairs.forEach(([a, b], j) => {
        runs.push({ text: a, options: { bullet: { indent: 14 }, breakLine: true, color: K.strong } });
        runs.push({ text: "→ " + b, options: { indentLevel: 1, bullet: false, breakLine: j < c.pairs.length - 1, paraSpaceAfter: 10, color: K.body } });
      });
      T(s, runs, { x: x + 0.35, y: Y0 + 1.4, w: W3 - 0.6, h: h - 1.6, fontSize: 14 });
    }
  });
  T(s, "RFP 8항 \"예산 수준별 우선순위 및 제외 항목 제시\" 기준에 따라 범위를 정의했습니다.", { x: X0, y: 6.4, w: CW, h: 0.3, fontSize: 12, color: K.muted });
  s.addNotes("PT 질문 11번(직접 운영 가능 범위) 대응: 온라인은 직접 운영, 조사·오프라인·유통 영업은 범위 외임을 명확히 합니다.");
}

// =========================================================
// Slide 4. 현황 진단 ① 매출과 리뷰
// =========================================================
pres.addSection({ title: "현황 진단" });
{
  const s = content("현황 진단", "03 현황 진단 ① 매출과 리뷰", "리뷰 부족이 광고 효율을 막는 구조입니다");
  const lw = 5.5, h = 4.55;
  card(s, X0, Y0, lw, h, K.surface);
  s.addChart(pres.charts.DOUGHNUT, [{ name: "온라인 월 매출", labels: ["네이버 스마트스토어", "쿠팡"], values: [200, 100] }], Object.assign({
    x: X0 + 0.2, y: Y0 + 0.15, w: lw - 0.4, h: h - 0.3,
    holeSize: 58, chartColors: [HEX.primary, HEX.primarySoft],
    showTitle: true, title: "온라인 월 매출 구성 (만 원)", titleFontSize: 15, titleColor: HEX.strong, titleBold: true,
    showValue: true, showPercent: false, showLabel: false, dataLabelColor: HEX.strong, dataLabelFontSize: 14, dataLabelFontBold: true,
    showLegend: true, legendPos: "b", legendFontSize: 13, legendColor: HEX.body,
    dataBorder: { pt: 2, color: HEX.surface }, objectName: "chart-sales-mix",
  }, chartFont));

  const rx = X0 + lw + GAP, rw = CW - lw - GAP;
  card(s, rx, Y0, rw, 2.0, K.surface);
  T(s, "17개", { x: rx + 0.4, y: Y0 + 0.25, w: 2.6, h: 0.85, fontSize: 48, bold: true, color: K.primary, valign: "middle", fit: "none" });
  T(s, "스마트스토어 상품당 최대 리뷰 수", { x: rx + 0.4, y: Y0 + 1.15, w: rw - 0.8, h: 0.32, fontSize: 15, bold: true, color: K.strong });
  T(s, "경쟁 브랜드는 리뷰 수천~수만 건", { x: rx + 0.4, y: Y0 + 1.5, w: rw - 0.8, h: 0.3, fontSize: 14, color: K.muted });

  const iy = Y0 + 2.0 + GAP, ih = h - 2.0 - GAP;
  card(s, rx, iy, rw, ih, K.primary);
  T(s, "시사점", { x: rx + 0.4, y: iy + 0.35, w: 2, h: 0.3, fontSize: 14, bold: true, color: K.primarySoft });
  T(s, "지금 광고비를 크게 늘리면 클릭은 늘지만 구매는 따라오지 않습니다. 광고보다 리뷰 확보가 먼저입니다.", { x: rx + 0.4, y: iy + 0.8, w: rw - 0.8, h: ih - 1.05, fontSize: 17, bold: true, color: K.white, lineSpacingMultiple: 1.3 });
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
  const h = 3.55;
  items.forEach(([n, t, d], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, h, K.surface);
    badge(s, x + 0.35, Y0 + 0.35, 0.55, n);
    T(s, t, { x: x + 0.35, y: Y0 + 1.15, w: W3 - 0.7, h: 0.45, fontSize: 20, bold: true, color: K.strong });
    T(s, d, { x: x + 0.35, y: Y0 + 1.7, w: W3 - 0.7, h: h - 1.9, fontSize: 14, lineSpacingMultiple: 1.25 });
  });
  const by = Y0 + h + GAP;
  card(s, X0, by, CW, 0.8, K.primaryLight);
  T(s, [
    { text: "활용 가능 자산", options: { bold: true, color: K.primary } },
    { text: "    첫 구매 3,000원 회원가입 쿠폰 운영 중 · 메타 픽셀 설치 완료 (추후 리타겟팅 기반)", options: { color: K.strong } },
  ], { x: X0 + 0.35, y: by, w: CW - 0.7, h: 0.8, fontSize: 15, valign: "middle" });
  s.addNotes("ym-kr.com 자사몰 확인 기준.\n\n근거: 사이트 타이틀 '양명 킹오브김 ( 해풀잎 )' / KING OF KIM 도시락김 36봉 22,400원, 72봉 31,400원, 김자반 50g×5봉 18,600원 / 선물세트 사랑둥이 C1호 53,000원 ~ 효자둥이 특품 168,000원 / 첫 쇼핑 3,000원 할인 회원가입 쿠폰 배너 / 메타(페이스북) 픽셀 설치 확인 / 하단 SNS 링크가 instagram.com, youtube.com 등 메인 주소로만 연결.");
}

// =========================================================
// Slide 6. 원인 확인 방법
// =========================================================
{
  const s = content("현황 진단", "03 원인 확인 방법", "\"인지도 부족\"인지 운영 데이터로 구분합니다");
  const cx = { q: X0 + 1.15, m: X0 + 4.05, j: X0 + 8.05 };
  const cw = { q: 2.7, m: 3.7, j: 3.75 };
  T(s, "확인 질문", { x: cx.q, y: Y0, w: cw.q, h: 0.3, fontSize: 12, color: K.muted });
  T(s, "확인 지표", { x: cx.m, y: Y0, w: cw.m, h: 0.3, fontSize: 12, color: K.muted });
  T(s, "판단", { x: cx.j, y: Y0, w: cw.j, h: 0.3, fontSize: 12, color: K.muted });
  const rows = [
    ["1", "노출 문제?", "브랜드 검색량, 쇼핑검색 노출수", "검색량이 거의 없으면 인지도 문제"],
    ["2", "전환 문제?", "유입 대비 구매전환율", "유입은 있는데 구매가 없으면 리뷰 · 가격 · 상세페이지 문제"],
    ["3", "가격 문제?", "경쟁사 동일 규격 가격 · 리뷰 수, 세트별 전환율", "가격 경쟁력과 세트 구성 조정 필요 여부"],
  ];
  const rh = 1.1;
  rows.forEach(([n, q, m, j], i) => {
    const y = Y0 + 0.4 + i * (rh + GAP);
    card(s, X0, y, CW, rh, K.surface);
    badge(s, X0 + 0.35, y + (rh - 0.5) / 2, 0.5, n);
    T(s, q, { x: cx.q, y, w: cw.q, h: rh, fontSize: 18, bold: true, color: K.strong, valign: "middle" });
    T(s, m, { x: cx.m, y, w: cw.m, h: rh, fontSize: 14, valign: "middle" });
    T(s, j, { x: cx.j, y, w: cw.j, h: rh, fontSize: 14, bold: true, color: K.primaryDark, valign: "middle" });
  });
  T(s, "정량 소비자 조사는 범위 외입니다. 위 데이터는 매월 리포트로 제공해 양명의 추가 조사 · 의사결정 근거로 활용할 수 있게 합니다.", { x: X0, y: 6.4, w: CW, h: 0.3, fontSize: 12, color: K.muted });
  s.addNotes("PT 질문 1·2번 대응. 인지도 부족을 단정하지 않고 노출·전환·가격 중 어디서 막히는지 데이터로 구분합니다.");
}

// =========================================================
// Slide 7. 전략 방향
// =========================================================
pres.addSection({ title: "전략" });
{
  const s = content("전략", "04 전략 방향", "리뷰 → 검색 → 구매 → 재구매 순서로 갑니다");
  const flow = [["리뷰", "체험단으로 구매 결정 근거 확보"], ["검색", "쇼핑검색 · 브랜드 키워드 노출"], ["구매", "리뷰 갖춘 대표 상품으로 전환"], ["재구매", "알림받기 · 재구매 쿠폰 권장"]];
  const step = (CW + GAP) / 4, d = 1.4, cy = Y0 + 0.15;
  flow.forEach(([t, desc], i) => {
    const x = X0 + i * step;
    badge(s, x, cy, d, t, i === 2 ? K.primaryDark : K.primary, K.white, 20);
    T(s, desc, { x, y: cy + d + 0.25, w: step - 0.5, h: 0.6, fontSize: 15, color: K.body });
    if (i < 3) {
      s.addShape(pres.shapes.LINE, { x: x + d + 0.25, y: cy + d / 2, w: step - d - 0.5, h: 0, line: { color: K.primarySoft, width: 2.5, endArrowType: "triangle" }, objectName: `arrow-${i}` });
    }
  });
  const by = 4.45, bh = 2.0;
  card(s, X0, by, CW, bh, K.primaryLight);
  T(s, "브랜드명 · 패키지는 그대로 유지합니다", { x: X0 + 0.4, y: by + 0.35, w: CW - 0.8, h: 0.4, fontSize: 19, bold: true, color: K.primaryDark });
  T(s, "대신 상품명 · 광고 문구 · 체험단 콘텐츠에서 \"KING OF KIM\"을 일관되게 노출해, 구매 경험이 브랜드 검색으로 이어지게 만듭니다. 리뷰가 없는 상태의 광고는 밑 빠진 독에 물 붓기와 같아서, 리뷰부터 채우고 광고를 늘립니다.", { x: X0 + 0.4, y: by + 0.9, w: CW - 0.8, h: bh - 1.1, fontSize: 15, color: K.strong, lineSpacingMultiple: 1.3 });
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
  const h = 4.3;
  ch.forEach(([p, tag, name, bl], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, h, K.surface);
    T(s, p, { x: x + 0.35, y: Y0 + 0.3, w: 1.8, h: 0.75, fontSize: 40, bold: true, color: K.primary, valign: "middle", fit: "none" });
    const tw = tag.length > 2 ? 1.2 : 0.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + W3 - 0.35 - tw, y: Y0 + 0.47, w: tw, h: 0.4, rectRadius: 0.2, fill: { color: K.white }, line: { type: "none" }, objectName: `tag-${i}` });
    T(s, tag, { x: x + W3 - 0.35 - tw, y: Y0 + 0.47, w: tw, h: 0.4, fontSize: 13, bold: true, color: K.primary, align: "center", valign: "middle" });
    T(s, name, { x: x + 0.35, y: Y0 + 1.3, w: W3 - 0.7, h: 0.45, fontSize: 20, bold: true, color: K.strong });
    T(s, bullets(bl, { paraSpaceAfter: 8 }), { x: x + 0.35, y: Y0 + 1.95, w: W3 - 0.6, h: h - 2.15, fontSize: 14 });
  });
  T(s, "비중은 전체 광고비 기준 시작값(네이버:쿠팡 = 현재 매출 비중 2:1)이며, 3개월차부터 매체별 ROAS로 재배분합니다.", { x: X0, y: 6.4, w: CW, h: 0.3, fontSize: 12, color: K.muted });
  s.addNotes("PT 질문 6번 대응: 스마트스토어·쿠팡 판매량 증대 방법.");
}

// =========================================================
// Slide 9. 체험단 · 리뷰 확보
// =========================================================
{
  const s = content("전략", "04 체험단 · 리뷰 확보", "1~2개월차, 체험단으로 리뷰 기반부터 만듭니다");
  const h = 4.6;
  card(s, X0, Y0, W3, h, K.primary);
  T(s, "대표 상품 리뷰 목표", { x: X0 + 0.4, y: Y0 + 0.4, w: W3 - 0.8, h: 0.35, fontSize: 15, bold: true, color: K.primarySoft });
  [["50건+", "2개월차까지", 1.05], ["150건+", "6개월차까지", 2.75]].forEach(([n, l, oy]) => {
    T(s, n, { x: X0 + 0.4, y: Y0 + oy, w: W3 - 0.8, h: 0.8, fontSize: 44, bold: true, color: K.white, valign: "middle", fit: "none" });
    T(s, l, { x: X0 + 0.4, y: Y0 + oy + 0.85, w: W3 - 0.8, h: 0.35, fontSize: 15, color: K.white });
  });
  const rx = COL3[1], rw = CW - (rx - X0);
  const rows = [
    ["1", "역할 분담", "제품은 양명이 제공하고, 체험단 모집 · 운영 · 콘텐츠 관리는 NICECORE가 담당합니다."],
    ["2", "확보 채널", "스마트스토어 리뷰와 네이버 블로그 후기 중심. 블로그 후기는 \"킹오브김\" 검색 결과에도 노출됩니다."],
    ["3", "쿠팡 운영", "쿠팡은 외부 체험단 리뷰에 제한이 있어, 플랫폼 정책에 맞는 방식(쿠팡 체험단 프로그램 등)을 별도 검토합니다."],
  ];
  const rh = (h - 2 * GAP) / 3;
  rows.forEach(([n, t, d], i) => {
    const y = Y0 + i * (rh + GAP);
    card(s, rx, y, rw, rh, K.surface);
    badge(s, rx + 0.35, y + 0.22, 0.45, n);
    T(s, t, { x: rx + 1.05, y: y + 0.2, w: rw - 1.4, h: 0.5, fontSize: 18, bold: true, color: K.strong, valign: "middle" });
    T(s, d, { x: rx + 1.05, y: y + 0.72, w: rw - 1.4, h: rh - 0.8, fontSize: 14 });
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
  const h = 2.35;
  tl.forEach(([m, st, d], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, h, K.surface);
    T(s, m, { x: x + 0.35, y: Y0 + 0.3, w: 1.8, h: 0.5, fontSize: 26, bold: true, color: K.primary, valign: "middle" });
    T(s, st, { x: x + 0.35, y: Y0 + 0.88, w: W3 - 0.7, h: 0.35, fontSize: 17, bold: true, color: K.strong });
    T(s, d, { x: x + 0.35, y: Y0 + 1.35, w: W3 - 0.7, h: h - 1.5, fontSize: 14 });
    if (i < 2) {
      const ax = x + W3 + GAP / 2 - 0.22;
      s.addShape(pres.shapes.OVAL, { x: ax, y: Y0 + h / 2 - 0.22, w: 0.44, h: 0.44, fill: { color: K.white }, line: { color: K.border, width: 1 }, objectName: `tl-arrow-${i}` });
      T(s, "›", { x: ax, y: Y0 + h / 2 - 0.25, w: 0.44, h: 0.44, fontSize: 20, bold: true, color: K.primary, align: "center", valign: "middle" });
    }
  });
  const by = Y0 + h + GAP, bh = 6.45 - by, bw = (CW - GAP) / 2;
  [["5~16만 원대", "선물세트 가격대. 같은 광고비로 일상용 상품보다 큰 매출을 기대할 수 있습니다."],
   ["Sales Case", "명절 실적은 6개월 성과 자료의 핵심 사례로, 유통 MD · 기업 특판 제안에 활용됩니다."]].forEach(([n, d], i) => {
    const x = X0 + i * (bw + GAP);
    card(s, x, by, bw, bh, K.primaryLight);
    T(s, n, { x: x + 0.4, y: by + 0.35, w: bw - 0.8, h: 0.7, fontSize: 36, bold: true, color: K.primary, valign: "middle", fit: "none" });
    T(s, d, { x: x + 0.4, y: by + 1.2, w: bw - 0.8, h: bh - 1.35, fontSize: 15, color: K.strong });
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
    ["11월", "기반 구축", "계정 세팅 · Baseline 측정 · 체험단 1차", "참고 ROAS 100%"],
    ["12월", "리뷰 확보", "체험단 2차 · 광고 소재 테스트", "참고 ROAS 150%"],
    ["1월", "확대 ①", "설 선물세트 집중 · 광고 증액", "참고 ROAS 200%"],
    ["2월", "확대 ②", "명절 성과 분석 · 효율 채널 재배분", "목표 ROAS 300%"],
    ["3월", "집중", "효율 상품 · 채널에 예산 집중", "목표 ROAS 300%"],
    ["4월", "정리", "성과 유지 · 6개월 성과 자료 작성", "목표 ROAS 300%"],
  ];
  const cw = (CW - 5 * GAP) / 6, h = 2.85;
  rm.forEach(([m, st, act, roas], i) => {
    const x = X0 + i * (cw + GAP);
    card(s, x, Y0, cw, h, K.surface);
    T(s, m, { x: x + 0.22, y: Y0 + 0.25, w: cw - 0.44, h: 0.45, fontSize: 22, bold: true, color: K.primary, valign: "middle" });
    T(s, st, { x: x + 0.22, y: Y0 + 0.8, w: cw - 0.44, h: 0.35, fontSize: 17, bold: true, color: K.strong });
    T(s, bullets(act.split(" · "), { bullet: { indent: 10 }, paraSpaceAfter: 4 }), { x: x + 0.2, y: Y0 + 1.3, w: cw - 0.3, h: h - 1.45, fontSize: 14 });
    const goal = roas.startsWith("목표");
    const cy = Y0 + h + 0.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: 0.5, rectRadius: 0.25, fill: { color: goal ? K.primary : K.primaryLight }, line: { type: "none" }, objectName: `roas-chip-${i}` });
    T(s, roas, { x, y: cy, w: cw, h: 0.5, fontSize: 14, bold: true, color: goal ? K.white : K.primary, align: "center", valign: "middle" });
  });
  const by = 5.75;
  card(s, X0, by, CW, 0.8, K.white, { outline: true });
  T(s, [
    { text: "증액 조건", options: { bold: true, color: K.primary } },
    { text: "    1~2개월차는 리뷰 목표 달성 시 다음 단계로. 3개월차부터는 전월 ROAS 기준 달성 시 증액, 미달 시 동결 후 키워드 · 상품 정리", options: { color: K.strong } },
  ], { x: X0 + 0.35, y: by, w: CW - 0.7, h: 0.8, fontSize: 14, valign: "middle" });
  s.addNotes("1~3개월차는 리뷰를 쌓는 투자 구간이라 ROAS는 참고치. 4개월차부터 목표 ROAS 300% 적용. ROAS = 광고 매출 ÷ 광고비(수수료 제외).");
}

// =========================================================
// Slide 12. 예산 옵션
// =========================================================
{
  const s = content("실행 계획", "05 예산 옵션", "최대 도달 예산에 따라 3가지 안을 제안합니다");
  const months = ["1개월", "2개월", "3개월", "4개월", "5개월", "6개월"];
  const lw = 7.75, h = 4.3;
  card(s, X0, Y0, lw, h, K.white, { outline: true });
  s.addChart(pres.charts.BAR, [
    { name: "A안", labels: months, values: [150, 200, 300, 400, 500, 500] },
    { name: "B안 (권장)", labels: months, values: [200, 300, 450, 600, 750, 750] },
    { name: "C안", labels: months, values: [250, 400, 600, 800, 1000, 1000] },
  ], Object.assign({
    x: X0 + 0.15, y: Y0 + 0.1, w: lw - 0.3, h: h - 0.2, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HEX.primarySoft, HEX.primary, HEX.primaryDark],
    showTitle: true, title: "월별 광고비 (만 원, 네이버 + 쿠팡)", titleFontSize: 15, titleColor: HEX.strong, titleBold: true,
    showLegend: true, legendPos: "b", legendFontSize: 13, legendColor: HEX.body,
    catAxisLabelColor: HEX.body, catAxisLabelFontSize: 12, catAxisLineShow: true, catAxisLineColor: HEX.border,
    valAxisLabelColor: HEX.muted, valAxisLabelFontSize: 11, valAxisLabelFormatCode: "#,##0", valAxisLineShow: false,
    valAxisMaxVal: 1000, valAxisMajorUnit: 250,
    valGridLine: { color: HEX.border, size: 0.75 }, catGridLine: { style: "none" },
    objectName: "chart-budget-options",
  }, chartFont));
  const rx = X0 + lw + GAP, rw = CW - lw - GAP, ch = (h - 2 * GAP) / 3;
  [["A안", "최대 월 500만 원", "2,050만 원", false], ["B안 · 권장", "최대 월 750만 원", "3,050만 원", true], ["C안", "최대 월 1,000만 원", "4,050만 원", false]].forEach(([t, sub, tot, hl], i) => {
    const y = Y0 + i * (ch + GAP);
    card(s, rx, y, rw, ch, hl ? K.primary : K.surface);
    T(s, t, { x: rx + 0.35, y: y + 0.28, w: 2.2, h: 0.38, fontSize: 18, bold: true, color: hl ? K.white : K.strong });
    T(s, sub, { x: rx + 0.35, y: y + 0.7, w: 2.2, h: 0.3, fontSize: 14, color: hl ? K.primaryLight : K.body });
    T(s, [{ text: "6개월 합계", options: { fontSize: 12, bold: false, breakLine: true, color: hl ? K.primaryLight : K.muted } }, { text: tot }],
      { x: rx + rw - 0.35 - 1.9, y: y + 0.22, w: 1.9, h: ch - 0.44, fontSize: 20, bold: true, color: hl ? K.white : K.primary, align: "right", valign: "middle" });
  });
  T(s, "6개월 광고비 합계 기준 · VAT 별도 | 대행수수료: 쿠팡 광고비의 15% (네이버 없음) | 체험단 운영비 별도 견적 | 상세페이지 개선 선택 시 별도", { x: X0, y: 6.35, w: CW, h: 0.3, fontSize: 12, color: K.muted });
  s.addNotes("PT 질문 10번 대응. 매체별 상세 배분은 부록 참조. 컨설팅·조사·오프라인 비용은 범위 외로 해당 없음.");
}

// ---- table helpers ----
const tb = { type: "solid", pt: 0.75, color: HEX.border };
function cell(text, o = {}) {
  return { text, options: Object.assign({ border: [tb, tb, tb, tb] }, o) };
}

// =========================================================
// Slide 13. 목표 ROAS 300% 기준 예상 매출
// =========================================================
{
  const s = content("실행 계획", "05 목표 ROAS 300% 기준 예상 매출", "4개월차부터 광고비의 3배 매출을 목표로 합니다");
  [["A안", "1,500만 원", false], ["B안 · 권장", "2,250만 원", true], ["C안", "3,000만 원", false]].forEach(([o, v, hl], i) => {
    const x = COL3[i];
    card(s, x, Y0, W3, 1.5, hl ? K.primary : K.surface);
    T(s, `${o} · 6개월차 월 광고 매출`, { x: x + 0.35, y: Y0 + 0.28, w: W3 - 0.7, h: 0.3, fontSize: 14, bold: hl, color: hl ? K.white : K.body });
    T(s, v, { x: x + 0.35, y: Y0 + 0.68, w: W3 - 0.7, h: 0.6, fontSize: 36, bold: true, color: hl ? K.white : K.primary, valign: "middle", fit: "none" });
  });
  const head = ["구분", "1개월", "2개월", "3개월", "4개월", "5개월", "6개월", "6개월 합계", "평균 ROAS"];
  const data = [
    ["목표 ROAS", "100%", "150%", "200%", "300%", "300%", "300%", "-", "-"],
    ["A안", "150", "300", "600", "1,200", "1,500", "1,500", "5,250", "256%"],
    ["B안 (권장)", "200", "450", "900", "1,800", "2,250", "2,250", "7,850", "257%"],
    ["C안", "250", "600", "1,200", "2,400", "3,000", "3,000", "10,450", "258%"],
  ];
  const rows = [head.map((h, j) => cell(h, { bold: true, color: K.white, fill: { color: K.primary }, align: j === 0 ? "left" : "center" }))];
  data.forEach((r, i) => {
    const isGoal = i === 0, isB = i === 2;
    const fill = isGoal ? K.surface : isB ? K.primaryLight : K.white;
    rows.push(r.map((v, j) => cell(v, {
      fill: { color: fill }, bold: isB || j === 0 || (j === 8 && !isGoal),
      color: j === 8 && !isGoal ? K.primaryDark : (j === 0 || isB ? K.strong : K.body),
      align: j === 0 ? "left" : "center",
    })));
  });
  const fc = 1.5, sc = 1.5, oc = (CW - fc - 2 * sc) / 6;
  s.addTable(rows, { x: X0, y: 3.65, w: CW, colW: [fc, ...Array(6).fill(oc), sc, sc], rowH: 0.44, fontSize: 14, valign: "middle", margin: [0, 0.15, 0, 0.15], objectName: "table-roas" });
  T(s, "단위: 만 원. 광고 기여 매출 = 광고비 × 목표 ROAS. 현재 자연 매출(월 약 300만 원)은 별도이며, ROAS는 대행수수료를 제외한 광고비 기준입니다. 1~3개월차 ROAS는 리뷰 확보 단계의 참고치입니다.", { x: X0, y: 6.05, w: CW, h: 0.5, fontSize: 12, color: K.muted });
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
  const rows = [["영역", "KPI", "목표 · 측정 방법"].map((h) => cell(h, { bold: true, color: K.white, fill: { color: K.primary } }))];
  kpi.forEach(([a, b, c], i) => {
    const fill = i % 2 ? K.surface : K.white;
    const third = c === null
      ? [{ text: "Baseline 월 300만 원 대비, ", options: {} }, { text: "4개월차 이후 ROAS 300%", options: { bold: true, color: K.primaryDark } }]
      : c;
    rows.push([
      cell(a, { fill: { color: fill }, bold: true, color: K.strong }),
      cell(b, { fill: { color: fill }, color: K.strong }),
      cell(third, { fill: { color: fill }, color: K.body }),
    ]);
  });
  s.addTable(rows, { x: X0, y: Y0, w: CW, colW: [2.0, 4.4, CW - 6.4], rowH: 0.5, fontSize: 14, valign: "middle", margin: [0, 0.25, 0, 0.25], objectName: "table-kpi" });
  const by = 5.65;
  card(s, X0, by, CW, 0.8, K.primaryLight);
  T(s, [
    { text: "월간 리포트", options: { bold: true, color: K.primary } },
    { text: "    지표 추이, 매체별 효율, 예산 재배분 내역과 다음 달 운영 계획을 함께 공유합니다.", options: { color: K.strong } },
  ], { x: X0 + 0.35, y: by, w: CW - 0.7, h: 0.8, fontSize: 15, valign: "middle" });
  s.addNotes("KPI는 온라인 지표로 한정. 오프라인 판매·신규 입점·조달 매출은 범위 외.");
}

// =========================================================
// Slide 15. 기대 효과와 리스크
// =========================================================
pres.addSection({ title: "기대 효과" });
{
  const s = content("기대 효과", "06 기대 효과와 리스크", "온라인 성과가 유통 확대의 근거가 됩니다");
  const lw = 4.9, h = 4.3;
  card(s, X0, Y0, lw, h, K.primary);
  T(s, "6개월 후 결과물", { x: X0 + 0.4, y: Y0 + 0.4, w: lw - 0.8, h: 0.45, fontSize: 22, bold: true, color: K.white });
  ["리뷰가 갖춰진 대표 상품", "측정 가능한 브랜드 검색량", "MD · 조달 담당자에게 보여줄 온라인 판매 성과 자료"].forEach((t, i) => {
    const y = Y0 + 1.2 + i * 0.95;
    badge(s, X0 + 0.4, y, 0.45, "✓", K.white, K.primary, 16);
    T(s, t, { x: X0 + 1.05, y: y - 0.05, w: lw - 1.4, h: 0.8, fontSize: 16, bold: true, color: K.white });
  });
  const rx = X0 + lw + GAP, rw = CW - lw - GAP, rh = (h - 2 * GAP) / 3;
  const half = (rw - 0.7) / 2;
  [["리뷰 확보 지연", "증액 시점을 늦추고 체험단 추가 운영"], ["낮은 객단가로 효율 저조", "세트 · 대용량 위주로 전환, 묶음 구성 제안"], ["쿠팡 효율 저조", "네이버로 예산 이동"]].forEach(([r, a], i) => {
    const y = Y0 + i * (rh + GAP);
    card(s, rx, y, rw, rh, K.surface);
    T(s, "리스크", { x: rx + 0.35, y: y + 0.25, w: 1.2, h: 0.28, fontSize: 13, bold: true, color: K.danger });
    T(s, r, { x: rx + 0.35, y: y + 0.58, w: half - 0.2, h: rh - 0.7, fontSize: 16, bold: true, color: K.strong });
    T(s, "대응", { x: rx + 0.35 + half + 0.2, y: y + 0.25, w: 1.2, h: 0.28, fontSize: 13, bold: true, color: K.primary });
    T(s, a, { x: rx + 0.35 + half + 0.2, y: y + 0.58, w: half - 0.2, h: rh - 0.7, fontSize: 15, color: K.body });
  });
  T(s, "본 제안의 KPI는 온라인 지표에 한정되며, 오프라인 판매 · 신규 입점 · 조달 매출은 범위 외입니다.", { x: X0, y: 6.4, w: CW, h: 0.3, fontSize: 12, color: K.muted });
  s.addNotes("PT 질문 7·8번 대응: 오프라인 입점 확대의 근거를 온라인 판매 실적으로 먼저 만든다는 논리.");
}

// =========================================================
// Slide 16. 부록
// =========================================================
pres.addSection({ title: "부록" });
{
  const s = content("부록", "부록", "옵션별 · 매체별 월 광고비 상세");
  const head = ["옵션", "매체", "1개월", "2개월", "3개월", "4개월", "5개월", "6개월", "합계"];
  const groups = [
    ["A안", [["네이버 쇼핑검색", 80, 100, 150, 200, 250, 250, 1030], ["네이버 파워링크", 20, 30, 50, 70, 80, 80, 330], ["쿠팡 광고", 50, 70, 100, 130, 170, 170, 690], ["광고비 합계", 150, 200, 300, 400, 500, 500, 2050]]],
    ["B안\n(권장)", [["네이버 쇼핑검색", 100, 150, 230, 300, 380, 380, 1540], ["네이버 파워링크", 30, 50, 70, 100, 120, 120, 490], ["쿠팡 광고", 70, 100, 150, 200, 250, 250, 1020], ["광고비 합계", 200, 300, 450, 600, 750, 750, 3050]]],
    ["C안", [["네이버 쇼핑검색", 130, 200, 300, 400, 500, 500, 2030], ["네이버 파워링크", 40, 70, 100, 130, 170, 170, 680], ["쿠팡 광고", 80, 130, 200, 270, 330, 330, 1340], ["광고비 합계", 250, 400, 600, 800, 1000, 1000, 4050]]],
  ];
  const fmt = (n) => n.toLocaleString("en-US");
  const rows = [head.map((h, j) => cell(h, { bold: true, color: K.white, fill: { color: K.primary }, align: j < 2 ? "left" : "center" }))];
  groups.forEach(([opt, rs], gi) => {
    const isB = gi === 1;
    rs.forEach((r, ri) => {
      const total = ri === 3;
      const fill = total ? K.surface : isB ? K.primaryLight : K.white;
      const row = [];
      if (ri === 0) row.push(cell(opt, { rowspan: 4, fill: { color: isB ? K.primaryLight : K.white }, bold: true, color: isB ? K.primaryDark : K.strong, valign: "middle" }));
      r.forEach((v, j) => {
        row.push(cell(typeof v === "number" ? fmt(v) : v, {
          fill: { color: fill }, bold: total || isB && j === 7,
          color: total && j === 7 ? K.primaryDark : total ? K.strong : K.body,
          align: j === 0 ? "left" : "center",
        }));
      });
      rows.push(row);
    });
  });
  const c0 = 1.1, c1 = 2.2, cl = 1.45, cm = (CW - c0 - c1 - cl) / 6;
  s.addTable(rows, { x: X0, y: Y0, w: CW, colW: [c0, c1, ...Array(6).fill(cm), cl], rowH: 0.32, fontSize: 11, valign: "middle", margin: [0, 0.12, 0, 0.12], objectName: "table-appendix" });
  T(s, "단위: 만 원, VAT 별도. 쿠팡 대행수수료(광고비의 15%) 6개월 합계: A안 103.5 · B안 153 · C안 201", { x: X0, y: 6.3, w: CW, h: 0.3, fontSize: 12, color: K.muted });
  s.addNotes("각 월의 세 매체 합 = 해당 옵션 월 광고비. 쿠팡 수수료 = 쿠팡 광고 합계 × 0.15 (A 690×0.15=103.5, B 1,020×0.15=153, C 1,340×0.15=201).");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
