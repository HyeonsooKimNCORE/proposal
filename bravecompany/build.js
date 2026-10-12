// 브레이브컴퍼니 네이버 광고 진단 제안서 생성 스크립트
// 실행: node build.js  →  브레이브컴퍼니_네이버광고_진단_제안서.pptx
// 숫자 기준: 브레이브컴퍼니_PPT_작업지시서.md (+ 사용자 확인: 증액 효율 239%/2.3배, 건담닭 스크립트 미설치, 올라잇그린은 '제품군'으로만 표기)

const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply_theme.js");

const OUT = "브레이브컴퍼니_네이버광고_진단_제안서.pptx";
const A = (f) => path.join(__dirname, "assets", f);

// ---------- 테마 (NICECORE 템플릿 색 + 작업지시서 텍스트 색) ----------
const HEX = {
  text: "191F28",
  white: "FFFFFF",
  navy: "100442",
  gray: "F2F4F6",
  brand: "F13518", // 템플릿 오렌지레드 = 위험/강조
  sub: "4E5968",
  cap: "8B95A1",
  good: "03B26C",
  warn: "F08A00",
  slate: "6F7D97",
  line: "D1D6DB",
};
const THEME = {
  name: "NICECORE Diagnosis",
  headFontFace: "SUIT",
  bodyFontFace: "SUIT",
  colors: {
    dk1: HEX.text, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.gray,
    accent1: HEX.brand, accent2: HEX.sub, accent3: HEX.cap,
    accent4: HEX.good, accent5: HEX.warn, accent6: HEX.slate,
    hlink: "3182F6", folHlink: "6F7D97",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "NICECORE 퍼포먼스마케팅팀";
pres.company = "NICECORE";
pres.title = "브레이브컴퍼니 네이버 광고 개선 방향";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

const S = pres.SchemeColor;
const C = {
  text: S.text1, white: S.background1, navy: S.text2, gray: S.background2,
  brand: S.accent1, sub: S.accent2, cap: S.accent3, good: S.accent4, warn: S.accent5, slate: S.accent6,
};

const W = 13.333;
const MX = 0.6; // 좌우 여백
const CW = W - MX * 2; // 콘텐츠 폭 12.13

// ---------- 레이아웃 ----------
const darkFooter = [
  { image: { path: A("logo_white.png"), x: 11.1, y: 6.62, w: 1.7, h: 0.25 } },
  { text: { text: "Credible Outcome Reliable Evolution", options: { x: 11.1, y: 6.95, w: 2.0, h: 0.2, fontSize: 8, color: C.white, margin: 0 } } },
  { text: { text: "©2026 NICECORE. ALL RIGHTS RESERVED.", options: { x: 0.25, y: 7.1, w: 3.5, h: 0.2, fontSize: 7, color: C.white, margin: 0 } } },
];

pres.defineSlideMaster({
  title: "NC_COVER",
  background: { path: A("bg_cover.jpg") },
  objects: darkFooter,
});
pres.defineSlideMaster({
  title: "NC_CONTENTS",
  background: { path: A("bg_contents.jpg") },
  objects: darkFooter,
});
pres.defineSlideMaster({
  title: "NC_SECTION",
  background: { path: A("bg_section.jpg") },
  objects: [
    ...darkFooter,
    { placeholder: { options: { name: "title", type: "title", x: 1.3, y: 0.7, w: 10.5, h: 0.8, fontSize: 36, bold: true, color: C.white, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "NC_CONTENT",
  background: { color: C.white },
  objects: [
    { image: { path: A("logo_navy.png"), x: 11.43, y: 6.98, w: 1.3, h: 0.19 } },
    { text: { text: "©2026 NICECORE. ALL RIGHTS RESERVED.", options: { x: MX, y: 7.0, w: 4, h: 0.18, fontSize: 8, color: C.cap, margin: 0 } } },
    { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.38, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.brand, valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.72, w: CW, h: 0.7, fontSize: 28, bold: true, color: C.text, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 10.9, y: 6.98, w: 0.4, h: 0.2, fontSize: 8, color: C.cap, align: "right" },
});

// ---------- 공통 도우미 ----------
const T = (s, text, o) =>
  s.addText(text, { isTextBox: true, margin: 0, fontSize: 14, color: C.text, valign: "top", ...o });

const card = (s, x, y, w, h, fill = C.gray, extra = {}) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill, ...(extra.fill || {}) }, line: extra.line || { type: "none" }, ...(extra.opts || {}) });

const tint = (s, x, y, w, h, color, transparency = 90) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color, transparency }, line: { type: "none" } });

const badge = (s, x, y, n, fill = C.navy, d = 0.42, size = 14) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  T(s, String(n), { x, y, w: d, h: d, fontSize: size, bold: true, color: C.white, align: "center", valign: "middle" });
};

const content = (sec, title, kicker) => {
  const s = pres.addSlide({ masterName: "NC_CONTENT", sectionTitle: sec });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
};

const TOP = 1.7; // 콘텐츠 시작 y
const tableHead = (cells, fill = C.navy) =>
  cells.map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: fill }, fontSize: 14 } }));

const SEC = { A: "01 진단 요약", B: "02 검색광고", C: "03 GFA", D: "04 개선 제안", X: "부록" };

// =====================================================================
// 표지
// =====================================================================
pres.addSection({ title: "표지" });
{
  const s = pres.addSlide({ masterName: "NC_COVER", sectionTitle: "표지" });
  T(s, "NICECORE 광고 진단 제안", { x: 1.3, y: 1.2, w: 7, h: 0.35, fontSize: 16, bold: true, color: C.white });
  T(s, "브레이브컴퍼니\n네이버 광고 개선 방향", { x: 1.3, y: 1.7, w: 8.5, h: 1.6, fontSize: 40, bold: true, color: C.white, lineSpacingMultiple: 1.05 });
  T(s, "분석 기간 2026.07.01 ~ 09.30 (GFA 추이는 5월부터)  |  구매완료 전환 기준  |  2026년 10월", { x: 1.3, y: 3.45, w: 9, h: 0.35, fontSize: 14, color: C.white });
  s.addNotes(
    "브레이브컴퍼니 네이버 광고 진단 결과를 말씀드리겠습니다.\n" +
    "분석 기간은 2026년 7월 1일부터 9월 30일까지 3개월이고, GFA는 흐름을 보기 위해 5월부터 봤습니다.\n" +
    "모든 숫자는 총전환이 아니라 구매완료 전환 기준입니다."
  );
}

// =====================================================================
// 목차
// =====================================================================
{
  const s = pres.addSlide({ masterName: "NC_CONTENTS", sectionTitle: "표지" });
  T(s, "CONTENTS", { x: 0.9, y: 1.1, w: 4, h: 0.6, fontSize: 28, bold: true, color: C.white });
  const items = [
    ["01", "진단 요약", "성과가 아니라 구조의 문제"],
    ["02", "검색광고", "측정 · 브랜드검색 · 키워드 · 제외키워드"],
    ["03", "GFA", "증액 후 효율 · ADVoost 분할 · 역할 혼재"],
    ["04", "개선 제안", "실행 방향 · 실행 계획 · 성과 지표"],
  ];
  items.forEach(([n, t, d], i) => {
    const y = 1.15 + i * 0.95;
    T(s, [
      { text: n + "   ", options: { bold: true, fontSize: 20 } },
      { text: t, options: { bold: true, fontSize: 20 } },
    ], { x: 6.4, y, w: 5.6, h: 0.42, color: C.white, valign: "middle" });
    T(s, d, { x: 7.05, y: y + 0.42, w: 5.0, h: 0.3, fontSize: 13, color: C.white });
    s.addShape(pres.shapes.LINE, { x: 6.4, y: y + 0.82, w: 5.6, h: 0, line: { color: C.white, width: 1 } });
  });
  s.addNotes(
    "순서는 네 부분입니다. 먼저 계정 전체를 요약하고, 검색광고와 GFA를 차례로 본 뒤, 개선 제안과 실행 계획으로 마무리하겠습니다."
  );
}

const divider = (sec, num, title, sub, body, notes) => {
  pres.addSection({ title: sec });
  const s = pres.addSlide({ masterName: "NC_SECTION", sectionTitle: sec });
  s.addText(num + "  " + title, { placeholder: "title" });
  T(s, sub, { x: 1.3, y: 1.85, w: 10, h: 0.6, fontSize: 26, color: C.white });
  T(s, body, { x: 1.3, y: 2.75, w: 10, h: 0.4, fontSize: 16, color: C.white });
  s.addNotes(notes);
};

// =====================================================================
// 섹션 A. 진단 요약
// =====================================================================
divider(SEC.A, "01", "진단 요약", "성과가 나쁜 게 아니라, 구조의 문제입니다", "계정 전체 숫자  ·  광고비 구조  ·  문제 7가지",
  "먼저 계정 전체를 한눈에 보겠습니다.");

// --- 2. 핵심 메시지
{
  const s = content(SEC.A, "성과가 나쁜 게 아니라, 볼 수 없고 키울 수 없는 상태입니다", "01  진단 요약");
  const rows = [
    ["브랜드 고객만 받는 검색광고", "검색광고비 98.5%가 이미 브랜드를 아는 고객에게 쓰였습니다"],
    ["증액 후 꺾인 GFA", "예산을 2.3배 늘린 뒤 ROAS가 4개월 연속 떨어졌습니다"],
    ["측정 누락과 계약 미노출", "구매가 잡히지 않는 파워링크, 꺼져 있던 브랜드검색이 있습니다"],
  ];
  rows.forEach(([h, d], i) => {
    const y = TOP + 0.15 + i * 1.55;
    badge(s, MX, y, i + 1);
    T(s, h, { x: MX + 0.6, y: y - 0.02, w: 5.0, h: 0.42, fontSize: 18, bold: true, valign: "middle" });
    T(s, d, { x: MX + 0.6, y: y + 0.45, w: 5.0, h: 0.75, fontSize: 14, color: C.sub });
  });
  const stats = [
    ["590만 원", "브랜드검색 미노출 46일치 계약금", C.brand],
    ["98.5%", "검색광고비 중 자사 브랜드명 검색 비중", C.navy],
    ["241%", "9월 GFA ROAS. 364%에서 4개월 연속 하락", C.brand],
    ["121만 원", "구매가 측정되지 않는 파워링크 광고비", C.navy],
  ];
  const gx = 6.55, gw = 2.97, gh = 2.3, gap = 0.22;
  stats.forEach(([n, l, col], i) => {
    const x = gx + (i % 2) * (gw + gap), y = TOP + Math.floor(i / 2) * (gh + gap);
    card(s, x, y, gw, gh);
    T(s, n, { x: x + 0.3, y: y + 0.45, w: gw - 0.5, h: 0.7, fontSize: 30, bold: true, color: col, valign: "middle" });
    T(s, l, { x: x + 0.3, y: y + 1.3, w: gw - 0.55, h: 0.75, fontSize: 14, color: C.sub });
  });
  s.addNotes(
    "결론부터 말씀드리면, 성과가 나쁜 계정이 아닙니다. 다만 지금 구조로는 어디서 성과가 나는지 볼 수 없고, 더 키울 수도 없는 상태입니다.\n" +
    "세 가지입니다. 측정이 안 되는 곳, 돈이 새는 곳, 새 고객을 데려오는 구조가 없는 것.\n" +
    "오른쪽 네 숫자가 이 세 가지를 보여줍니다. 브랜드검색은 계약 기간 중 46일 동안 노출이 없었고, 그 기간 계약금이 약 590만 원입니다. " +
    "검색광고비의 98.5%는 이미 브랜드를 아는 고객의 검색에 쓰였습니다. GFA ROAS는 364%에서 241%까지 4개월 연속 내려왔고, " +
    "파워링크 121만 원은 구매가 아예 측정되지 않고 있습니다."
  );
}

// --- 3. 광고비 구조
{
  const s = content(SEC.A, "광고비의 88%가 GFA, 그중 90%가 ADVoost입니다", "01  진단 요약");
  const body = (t, o = {}) => ({ text: t, options: { fontSize: 14, color: C.text, ...o } });
  const rows = [
    tableHead(["구분", "광고비", "ROAS", "한 줄 진단"]),
    [body("검색광고 SA"), body("499만 원", { align: "right" }), body("10,875%", { align: "right" }), body("대부분 브랜드명 검색. 계약금 반영 시 4,069%")],
    [body("브랜드검색 계약금 (일할)"), body("834만 원", { align: "right" }), body("1,065%", { align: "right" }), body("계약 65일 중 46일 노출 0", { color: C.brand, bold: true })],
    [body("GFA ADVoost"), body("3,174만 원", { align: "right" }), body("306%", { align: "right" }), body("9개 캠페인 분할, 신규·기존 고객 혼재")],
    [body("GFA 수동 캠페인"), body("339만 원", { align: "right" }), body("약 175%", { align: "right" }), body("트래픽 목적, 소재 정체, 카탈로그 없음")],
    [body("합계 (계약금 포함)", { bold: true }), body("약 4,846만 원", { align: "right", bold: true }), body("—", { align: "right" }), body("보고서 기준 4,012만 원 + 계약금 해당분")],
  ];
  s.addTable(rows, {
    x: MX, y: TOP, w: CW, colW: [2.9, 1.8, 1.35, 6.08], rowH: 0.42,
    border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0.04, 0.12, 0.04, 0.12],
  });
  // 광고비 비중 (계약금 포함 4,846만 원 기준)
  s.addChart(pres.charts.BAR, [
    { name: "GFA ADVoost", labels: ["광고비"], values: [3174] },
    { name: "GFA 수동", labels: ["광고비"], values: [339] },
    { name: "검색광고 SA", labels: ["광고비"], values: [499] },
    { name: "브랜드검색 계약금", labels: ["광고비"], values: [834] },
  ], {
    x: MX - 0.1, y: 4.35, w: CW + 0.2, h: 1.25, barDir: "bar", barGrouping: "percentStacked", barGapWidthPct: 30,
    chartColors: [HEX.navy, HEX.slate, HEX.brand, HEX.warn],
    showValue: true, dataLabelPosition: "ctr", dataLabelColor: HEX.white, dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "#,##0",
    showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.sub,
    catAxisHidden: true, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showTitle: true, title: "광고비 구성 (만 원, 계약금 포함 약 4,846만 원)", titleFontSize: 12, titleFontFace: "+mn-lt", titleColor: HEX.sub,
  });
  tint(s, MX, 5.8, CW, 0.75, C.brand, 90);
  T(s, [
    { text: "SA ROAS 10,875%는 브랜드검색 계약금이 빠진 숫자입니다.  ", options: { bold: true } },
    { text: "계약금을 넣으면 4,069% — 실제보다 2.7배 부풀려 보이고 있습니다." },
  ], { x: MX + 0.3, y: 5.8, w: CW - 0.6, h: 0.75, fontSize: 15, valign: "middle" });
  s.addNotes(
    "3개월 광고비는 보고서 기준 4,012만 원이고, 보고서에 0원으로 찍히는 브랜드검색 계약금 해당분을 더하면 약 4,846만 원입니다.\n" +
    "보고서 기준 광고비의 88%가 GFA이고, 그중 90%가 ADVoost입니다. 사실상 ADVoost 한 상품에 계정이 기대고 있습니다.\n" +
    "검색광고 ROAS 10,875%는 계약금이 빠진 숫자입니다. 브랜드검색 계약금을 하루 단위로 나눠 분석 기간만큼 넣으면 4,069%로, 2.7배 차이가 납니다.\n" +
    "ADVoost 306% 자체는 나쁜 숫자가 아닙니다. 뒤에서 왜 더 좋아질 여지가 막혀 있는지 말씀드리겠습니다."
  );
}

// --- 4. 문제 7가지
{
  const s = content(SEC.A, "구조에서 생긴 문제 7가지", "01  진단 요약");
  const items = [
    ["위험", "측정 누락", "파워링크 121만 원, 구매 0건"],
    ["위험", "브랜드검색 미노출", "계약 65일 중 46일 노출 0"],
    ["위험", "신규 키워드 부재", "검색광고비 98.5%가 브랜드명 검색"],
    ["위험", "ADVoost 분할·혼재", "9개로 분할, 새 고객·기존 고객 섞임"],
    ["위험", "증액 후 효율 하락", "GFA ROAS 364% → 241%"],
    ["주의", "올라잇그린 이름 충돌", "클릭당 비용이 꼬기다의 4.7배"],
    ["주의", "기본 관리 부재", "제외키워드 0개, 새는 비용 217만 원"],
  ];
  const cw = (CW - 3 * 0.25) / 4, ch = 2.25;
  items.forEach(([lv, h, d], i) => {
    const x = MX + (i % 4) * (cw + 0.25), y = TOP + 0.05 + Math.floor(i / 4) * (ch + 0.25);
    const col = lv === "위험" ? C.brand : C.warn;
    card(s, x, y, cw, ch);
    badge(s, x + 0.25, y + 0.25, i + 1, col);
    T(s, lv, { x: x + 0.8, y: y + 0.25, w: 1.0, h: 0.42, fontSize: 12, bold: true, color: col, valign: "middle" });
    T(s, h, { x: x + 0.25, y: y + 0.9, w: cw - 0.45, h: 0.45, fontSize: 17, bold: true });
    T(s, d, { x: x + 0.25, y: y + 1.38, w: cw - 0.45, h: 0.7, fontSize: 14, color: C.sub });
  });
  const x = MX + 3 * (cw + 0.25), y = TOP + 0.05 + ch + 0.25;
  card(s, x, y, cw, ch, C.navy);
  T(s, "3가지 방향으로\n해결합니다", { x: x + 0.25, y: y + 0.3, w: cw - 0.45, h: 0.95, fontSize: 18, bold: true, color: C.white });
  T(s, "보이게 만들고\n새는 곳을 막고\n키웁니다", { x: x + 0.25, y: y + 1.25, w: cw - 0.45, h: 0.9, fontSize: 14, color: C.white });
  s.addNotes(
    "세부적으로는 일곱 가지 문제가 보였습니다. 다섯 개는 위험, 두 개는 주의 단계로 분류했습니다.\n" +
    "중요한 건, 이 문제들이 누군가 잘못해서 생긴 게 아니라 지금 구조에서 자연스럽게 나오는 숫자라는 점입니다.\n" +
    "그래서 해결도 구조로 합니다. 보이게 만들고, 새는 곳을 막고, 키우는 순서입니다."
  );
}

// =====================================================================
// 섹션 B. 검색광고
// =====================================================================
divider(SEC.B, "02", "검색광고", "검색광고는 이미 아는 고객만 받고 있습니다", "측정 누락  ·  브랜드검색 미노출  ·  신규 키워드  ·  제외키워드",
  "검색광고부터 보겠습니다. 숫자는 좋아 보이지만, 대부분 이미 브랜드를 아는 고객의 검색입니다.");

// --- 5. 측정 누락
{
  const s = content(SEC.B, "파워링크 121만 원, 구매가 하나도 집계되지 않았습니다", "02  검색광고");
  card(s, MX, TOP, 4.9, 4.85);
  T(s, "건담닭 · 올라잇그린 파워링크 (3개월)", { x: MX + 0.35, y: TOP + 0.3, w: 4.3, h: 0.35, fontSize: 14, bold: true, color: C.sub });
  const big = [["121만 원", "광고비", C.text], ["7,672회", "클릭", C.text], ["0건", "구매", C.brand]];
  big.forEach(([n, l, col], i) => {
    const y = TOP + 0.9 + i * 1.25;
    T(s, n, { x: MX + 0.35, y, w: 3.2, h: 0.8, fontSize: 40, bold: true, color: col, valign: "middle" });
    T(s, l, { x: MX + 3.55, y, w: 1.0, h: 0.8, fontSize: 16, color: C.sub, valign: "middle" });
  });
  const rx = 5.85, rw = W - MX - rx;
  const pts = [
    ["확인된 원인", "올라잇그린·건담닭 파워링크 모두 네이버 전환 스크립트 미설치 (올라잇그린 랜딩: 자사몰 allrightgreen.co.kr)", C.brand],
    ["구매가 없는 게 아니라 잡히지 않는 것", "같은 브랜드의 쇼핑검색에서는 구매 225건이 집계됐습니다. 두 캠페인 검색어의 99%가 자사 브랜드명입니다", C.navy],
    ["효과를 모르는 채 증액", "올라잇그린 파워링크는 9월에 68% 증액 (8월 21만 원 → 9월 35만 원)", C.navy],
    ["SA 광고비의 24%", "성과를 알 수 없는 상태로 집행되고 있습니다", C.navy],
  ];
  pts.forEach(([h, d, col], i) => {
    const y = TOP + i * 1.24;
    badge(s, rx, y + 0.02, i + 1, col, 0.4, 13);
    T(s, h, { x: rx + 0.6, y, w: rw - 0.6, h: 0.4, fontSize: 16, bold: true, valign: "middle" });
    T(s, d, { x: rx + 0.6, y: y + 0.43, w: rw - 0.6, h: 0.7, fontSize: 14, color: C.sub });
  });
  s.addNotes(
    "건담닭과 올라잇그린 파워링크에 3개월 동안 121만 원을 썼고 클릭은 7,672회였는데, 구매는 0건으로 나옵니다.\n" +
    "두 캠페인 검색어의 99%가 자사 브랜드명입니다. 브랜드를 검색하고 들어온 고객이 하나도 안 샀다는 건 성과 문제가 아니라 측정 문제입니다. " +
    "실제로 같은 브랜드의 쇼핑검색에서는 구매가 225건 잡혔습니다.\n" +
    "확인해 보니 두 캠페인 모두 랜딩 페이지에 네이버 전환 스크립트가 설치돼 있지 않았습니다.\n" +
    "그런데 올라잇그린 파워링크는 9월에 오히려 68% 증액됐습니다. 이 돈이 효과가 있는지 없는지, 지금은 판단할 수 없습니다. " +
    "검색광고비의 24%가 이런 상태입니다."
  );
}

// --- 6. 브랜드검색 미노출
{
  const s = content(SEC.B, "계약 65일 중 46일, 브랜드검색이 꺼져 있었습니다", "02  검색광고");
  // 7/1 ~ 9/30 타임라인 (92일)
  const segs = [
    { d: 13, c: C.slate, top: "노출", bot: "7/1~7/13\n이전 계약분" },
    { d: 14, c: C.line, top: "노출 0", bot: "7/14~7/27" },
    { d: 4, c: C.navy, top: "", bot: "" },
    { d: 46, c: C.brand, top: "노출 0  ·  46일 (8월 전체 포함)", bot: "8/1~9/15" },
    { d: 15, c: C.navy, top: "노출", bot: "9/16~9/30" },
  ];
  const tx = MX, tw = CW, ty = 2.1, th = 0.5;
  let cx = tx;
  segs.forEach((g) => {
    const w = (g.d / 92) * tw;
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: ty, w, h: th, fill: { color: g.c === C.line ? "D1D6DB" : g.c }, line: { color: C.white, width: 1.5 } });
    if (g.top) T(s, g.top, { x: cx + 0.08, y: ty, w: w - 0.1, h: th, fontSize: 12, bold: true, color: g.c === C.line ? C.sub : C.white, valign: "middle" });
    if (g.bot) T(s, g.bot, { x: cx, y: ty + th + 0.08, w: Math.max(w, 1.2), h: 0.45, fontSize: 11, color: C.sub });
    cx += w;
  });
  // 계약 시작 표시 (7/28 = 27일째)
  const sx = tx + (27 / 92) * tw;
  s.addShape(pres.shapes.LINE, { x: sx, y: ty - 0.3, w: 0, h: 0.3, line: { color: C.navy, width: 1.5 } });
  T(s, "7/28 새 계약 시작 (90일, 1,155만 원, 하루 128,333원)", { x: sx + 0.08, y: ty - 0.4, w: 6, h: 0.28, fontSize: 12, bold: true, color: C.navy, valign: "middle" });

  // 블랙키위 캡처
  const iy = 3.3, iw = 6.2, ih = iw * 344 / 1201;
  s.addImage({ path: A("cap_blackkiwi.png"), x: MX, y: iy, w: iw, h: ih, altText: "블랙키위 '꼬기다' 월간 검색량 추이" });
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: iy, w: iw, h: ih, fill: { type: "none" }, line: { color: "D1D6DB", width: 0.75 } });
  const ax = MX + (1055 / 1201) * iw, ay = iy + (40 / 344) * ih;
  s.addShape(pres.shapes.OVAL, { x: ax - 0.17, y: ay - 0.17, w: 0.34, h: 0.34, fill: { type: "none" }, line: { color: C.brand, width: 2.25 } });
  T(s, "8월 약 23,900건", { x: ax - 1.85, y: ay - 0.15, w: 1.6, h: 0.3, fontSize: 12, bold: true, color: C.brand, align: "right", valign: "middle" });
  T(s, "'꼬기다' 월간 검색량 (블랙키위, 2025.10~2026.09) · 그래프 판독 근사치", { x: MX, y: iy + ih + 0.08, w: iw, h: 0.25, fontSize: 11, color: C.cap });
  tint(s, MX, 5.85, iw, 0.7, C.good, 88);
  T(s, "상품 선택은 적정: '꼬기다' 모바일 월 21,300건 → 라이트형 모바일이 맞음. 문제는 노출 관리", { x: MX + 0.25, y: 5.85, w: iw - 0.4, h: 0.7, fontSize: 13, valign: "middle" });

  const rx = 7.15, rw = W - MX - rx;
  const cards = [
    ["약 590만 원", "노출 0인 46일치 계약금 (5,903,333원)"],
    ["약 23,900건", "8월 '꼬기다' 검색량, 최근 12개월 최고. 그 한 달 내내 노출 0"],
    ["198 → 471회", "파워링크 '꼬기다' 하루 클릭. 브랜드검색 공백을 파워링크가 대신 메움"],
  ];
  cards.forEach(([n, l], i) => {
    const y = 3.3 + i * 1.1;
    card(s, rx, y, rw, 0.98);
    T(s, n, { x: rx + 0.25, y, w: 2.2, h: 0.98, fontSize: 22, bold: true, color: C.brand, valign: "middle" });
    T(s, l, { x: rx + 2.5, y, w: rw - 2.7, h: 0.98, fontSize: 13, color: C.sub, valign: "middle" });
  });
  s.addNotes(
    "먼저 잘하신 부분부터 말씀드리면, 브랜드검색 상품 선택은 적정했습니다. '꼬기다' 모바일 월 검색량이 21,300건이라 라이트형 모바일이 맞습니다.\n" +
    "문제는 노출 관리입니다. 7월 28일부터 90일, 1,155만 원 계약인데, 분석 기간에 해당하는 65일 중 실제로 노출된 날은 19일뿐이고 46일은 노출이 0이었습니다. " +
    "8월 1일부터 9월 15일까지입니다. 이 기간 계약금이 약 590만 원입니다.\n" +
    "하필 8월은 '꼬기다' 검색량이 약 23,900건으로 최근 12개월 중 가장 많았던 달입니다.\n" +
    "노출이 0인 날에도 보고서에 브랜드검색 구매가 찍히는데, 이건 그 전에 클릭한 고객의 지연 전환입니다. 그래서 노출수 기준으로 판단했습니다.\n" +
    "브랜드검색이 꺼진 동안 파워링크 '꼬기다'가 대신 클릭을 받아, 하루 클릭이 198회에서 471회로, 하루 비용이 18,767원에서 34,576원으로 늘었습니다. " +
    "반대로 브랜드검색이 켜진 날에도 파워링크가 평균 1.1위를 유지하고 있어, 같은 고객에게 이중으로 지출하고 있습니다.\n" +
    "참고로 계약 기간 브랜드검색 실질 ROAS는 1,065%, 클릭당 약 1,050원인데, 이건 미노출 때문에 높아진 숫자입니다.\n" +
    "미노출 원인은 광고 OFF, 소재 검수 등 여러 가능성이 있어 계정 이력 확인이 필요합니다."
  );
}

// --- 7. 신규 키워드 부재
{
  const s = content(SEC.B, "검색광고비 98.5%가 이미 브랜드를 아는 고객에게 쓰였습니다", "02  검색광고");
  s.addChart(pres.charts.BAR, [
    { name: "브랜드명 검색", labels: ["꼬기다", "건담닭", "올라잇그린", "통영숙수"], values: [278.6, 58.9, 83.2, 0.1] },
    { name: "일반 검색", labels: ["꼬기다", "건담닭", "올라잇그린", "통영숙수"], values: [2.8, 1.6, 1.9, 49.1] },
  ], {
    x: MX - 0.1, y: TOP, w: 6.7, h: 4.35, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 45,
    chartColors: [HEX.navy, HEX.brand],
    catAxisOrientation: "maxMin", valAxisHidden: true,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.text, dataLabelFormatCode: "#,##0.0",
    catAxisLabelColor: HEX.text, catAxisLabelFontSize: 14, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
    valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt", legendColor: HEX.sub,
    showTitle: true, title: "브랜드별 검색어 광고비 (만 원, 3개월)", titleFontSize: 14, titleFontFace: "+mn-lt", titleColor: HEX.text,
  });
  T(s, "꼬기다·건담닭·올라잇그린 기준 98.5%, 통영숙수 포함 전체 88%", { x: MX, y: 6.15, w: 6.4, h: 0.3, fontSize: 12, color: C.cap });

  const rx = 7.35, rw = W - MX - rx;
  T(s, "하지 않았는데도 팔렸습니다", { x: rx, y: TOP, w: rw, h: 0.45, fontSize: 20, bold: true });
  T(s, "닭가슴살·닭다리살 같은 카테고리 키워드는 파워링크에 등록돼 있지 않습니다", { x: rx, y: TOP + 0.5, w: rw, h: 0.6, fontSize: 14, color: C.sub });
  card(s, rx, TOP + 1.25, rw, 1.75);
  T(s, "'닭가슴살' 포함 검색어 898개가 쇼핑·확장검색으로 우연히 유입", { x: rx + 0.25, y: TOP + 1.38, w: rw - 0.5, h: 0.55, fontSize: 14, bold: true });
  [["3.8만 원", "비용"], ["34건", "구매"], ["191만 원", "매출"]].forEach(([n, l], i) => {
    const x = rx + 0.25 + i * ((rw - 0.5) / 3);
    T(s, n, { x, y: TOP + 2.0, w: (rw - 0.5) / 3, h: 0.5, fontSize: 22, bold: true, color: C.navy, valign: "middle" });
    T(s, l, { x, y: TOP + 2.5, w: (rw - 0.5) / 3, h: 0.3, fontSize: 12, color: C.sub });
  });
  card(s, rx, TOP + 3.15, rw, 1.0);
  T(s, [
    { text: "타사 브랜드 검색에서 구매 27건", options: { bold: true, breakLine: true } },
    { text: "'고기남자 닭가슴살', '워크맨 닭가슴살' 등", options: { color: C.sub, fontSize: 13 } },
  ], { x: rx + 0.25, y: TOP + 3.15, w: rw - 0.5, h: 1.0, fontSize: 14, valign: "middle" });
  card(s, rx, TOP + 4.3, rw, 0.55, C.navy);
  T(s, "제대로 하면 더 팔 수 있습니다", { x: rx + 0.25, y: TOP + 4.3, w: rw - 0.5, h: 0.55, fontSize: 15, bold: true, color: C.white, valign: "middle" });
  s.addNotes(
    "꼬기다, 건담닭, 올라잇그린 세 브랜드의 검색광고비 중 98.5%가 자사 브랜드명 검색에 쓰였습니다. 통영숙수까지 넣어도 88%입니다.\n" +
    "브랜드명 검색은 이미 우리를 아는 고객입니다. 새 고객을 데려오는 키워드, 예를 들어 닭가슴살이나 닭다리살 같은 카테고리 키워드는 파워링크에 등록돼 있지 않습니다.\n" +
    "그런데도 쇼핑검색과 확장검색을 통해 '닭가슴살'이 들어간 검색어 898개가 우연히 들어왔고, 3.8만 원으로 34건, 191만 원어치가 팔렸습니다. " +
    "'고기남자 닭가슴살', '워크맨 닭가슴살' 같은 타사 브랜드 검색에서도 27건이 팔렸습니다.\n" +
    "하지 않았는데도 팔렸다는 건, 제대로 하면 더 팔 수 있다는 뜻입니다."
  );
}

// --- 8-1. 올라잇그린 숫자
{
  const s = content(SEC.B, "'올라잇그린'은 브랜드 키워드처럼 작동하지 않습니다", "02  검색광고");
  T(s, "파워링크 브랜드 키워드 그룹 비교", { x: MX, y: TOP, w: 6, h: 0.35, fontSize: 16, bold: true });
  const hl = { fill: { color: "FDE4DF" } };
  const c = (t, o = {}) => ({ text: t, options: { fontSize: 15, ...o } });
  s.addTable([
    tableHead(["브랜드 키워드 그룹", "클릭률", "클릭당 비용"]),
    [c("꼬기다"), c("25.8%", { align: "right" }), c("77원", { align: "right" })],
    [c("건담닭"), c("24.4%", { align: "right" }), c("79원", { align: "right" })],
    [c("올라잇그린", { bold: true, ...hl }), c("9.8%", { align: "right", bold: true, color: C.brand, ...hl }), c("361원", { align: "right", bold: true, color: C.brand, ...hl })],
  ], { x: MX, y: TOP + 0.5, w: 6.0, colW: [2.6, 1.6, 1.8], rowH: 0.5, border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0.04, 0.15, 0.04, 0.15] });
  card(s, MX, 4.45, 2.9, 2.05);
  T(s, "4.7배", { x: MX + 0.25, y: 4.6, w: 2.5, h: 0.8, fontSize: 40, bold: true, color: C.brand, valign: "middle" });
  T(s, "올라잇그린 클릭당 비용\n(꼬기다 대비)", { x: MX + 0.25, y: 5.45, w: 2.5, h: 0.8, fontSize: 14, color: C.sub });
  tint(s, MX + 3.1, 4.45, 2.9, 2.05, C.good, 88);
  T(s, "품질지수 10/10", { x: MX + 3.35, y: 4.6, w: 2.5, h: 0.8, fontSize: 24, bold: true, color: C.good, valign: "middle" });
  T(s, "광고 품질은 만점.\n원인은 경쟁 입찰과 이름 충돌", { x: MX + 3.35, y: 5.45, w: 2.5, h: 0.8, fontSize: 14, color: C.text });

  const rx = 7.1, rw = W - MX - rx;
  T(s, "검색어 단위로 보면 두 갈래입니다", { x: rx, y: TOP, w: rw, h: 0.35, fontSize: 16, bold: true });
  const two = [
    ["'올라잇그린' 자체", [["22.3%", "클릭률 (양호)"], ["407원", "클릭당 비용"]], "타사 비오틴 제품 '올라잇' 셀러와 입찰 경쟁"],
    ["'올라잇' 단독", [["7,456회", "노출"], ["2.0%", "클릭률"]], "그룹 평균 클릭률을 끌어내림"],
  ];
  two.forEach(([h, nums, d], i) => {
    const y = TOP + 0.5 + i * 2.2;
    card(s, rx, y, rw, 2.05);
    T(s, h, { x: rx + 0.3, y: y + 0.2, w: rw - 0.6, h: 0.4, fontSize: 16, bold: true });
    nums.forEach(([n, l], j) => {
      const x = rx + 0.3 + j * 2.4;
      T(s, n, { x, y: y + 0.65, w: 2.3, h: 0.55, fontSize: 26, bold: true, color: C.navy, valign: "middle" });
      T(s, l, { x, y: y + 1.2, w: 2.3, h: 0.28, fontSize: 12, color: C.sub });
    });
    T(s, "→ " + d, { x: rx + 0.3, y: y + 1.55, w: rw - 0.6, h: 0.35, fontSize: 14, bold: true, color: C.brand });
  });
  s.addNotes(
    "올라잇그린은 브랜드 키워드인데도 브랜드 키워드처럼 작동하지 않습니다.\n" +
    "꼬기다와 건담닭은 클릭률이 25% 안팎이고 클릭당 비용이 80원 이하인데, 올라잇그린은 클릭률 9.8%, 클릭당 361원으로 꼬기다의 4.7배입니다.\n" +
    "광고 품질 문제는 아닙니다. 품질지수는 10점 만점입니다. 원인은 경쟁 입찰입니다.\n" +
    "검색어로 쪼개 보면 두 갈래입니다. '올라잇그린'이라고 정확히 검색한 경우는 클릭률 22.3%로 괜찮은데, 클릭당 407원입니다. 같은 이름의 다른 회사 제품을 파는 셀러들과 경쟁하기 때문입니다. " +
    "'올라잇'만 검색한 경우는 노출이 7,456회인데 클릭률이 2%라 그룹 평균을 끌어내립니다.\n" +
    "올라잇그린을 검색하는 사람 중 상당수는 다른 회사 비오틴 제품을 찾는 사람입니다. 참고로 구매 0건은 앞서 말씀드린 스크립트 미설치 때문입니다."
  );
}

// --- 8-2. 올라잇그린 캡처
{
  const s = content(SEC.B, "'올라잇그린' 검색 결과 상단은 다른 회사 제품입니다", "02  검색광고");
  const ih = 3.75, iy = TOP;
  const caps = [
    { f: "cap_search.png", r: 707 / 692, cap: "통합검색: 타사 비오틴 제품 '올라잇' 셀러(지혜홈쿡·미쁘다) 광고가 공식 광고보다 위에 노출" },
    { f: "cap_shopping.png", r: 432 / 795, cap: "쇼핑검색: 상단은 올라잇뷰티의 왁싱 제품. 공식 브랜드 그룹 평균 4.6위" },
    { f: "cap_ext_copy.png", r: 538 / 498, cap: "확장소재: 추가홍보문구가 맛·성분 위주라 제품군이 드러나지 않음. 옆 셀러는 '본품 증정' 이벤트 노출" },
  ];
  const ws = caps.map((c) => ih * c.r);
  const gap = (CW - ws.reduce((a, b) => a + b, 0)) / 2;
  let x = MX;
  caps.forEach((c, i) => {
    const w = ws[i];
    s.addImage({ path: A(c.f), x, y: iy, w, h: ih, altText: c.cap });
    s.addShape(pres.shapes.RECTANGLE, { x, y: iy, w, h: ih, fill: { type: "none" }, line: { color: "D1D6DB", width: 0.75 } });
    badge(s, x - 0.12, iy - 0.12, i + 1, C.brand, 0.38, 13);
    T(s, c.cap, { x, y: iy + ih + 0.15, w: Math.max(w, 2.3), h: 1.0, fontSize: 12, color: C.sub });
    x += w + gap;
  });
  s.addNotes(
    "실제 검색 화면입니다.\n" +
    "첫 번째, 통합검색에서 '올라잇그린'을 치면 지혜홈쿡, 미쁘다 같은 셀러 광고가 공식 광고보다 위에 나옵니다. 이 셀러들은 다른 회사의 비오틴 제품 '올라잇'을 파는 곳입니다. 그린, 블루 패키지로 나와서 이름이 겹칩니다.\n" +
    "두 번째, 쇼핑검색에서는 올라잇뷰티의 왁싱 제품이 상단을 차지하고, 공식 브랜드 그룹은 평균 4.6위입니다.\n" +
    "세 번째, 공식 광고의 추가홍보문구는 '속세의 맛', '19종 유산균'처럼 맛과 성분 위주라, 이게 어떤 제품군인지가 드러나지 않습니다. 설명문구 뒤에 슬래시로 이어 붙는 형태입니다. " +
    "옆의 셀러들은 '본품 증정' 같은 이벤트 확장소재를 보여주고 있습니다.\n" +
    "그래서 제안은 제품군 키워드를 따로 묶고, 문구에 제품군과 혜택을 넣고, '비오틴', '블루', '공구' 같은 검색어를 제외하는 것입니다."
  );
}

// --- 9. 제외키워드 / 새는 곳
{
  const s = content(SEC.B, "검색어 7,602개가 들어오는데 제외키워드는 0개입니다", "02  검색광고");
  card(s, MX, TOP, 3.9, 4.25, C.navy);
  T(s, "등록된 제외키워드", { x: MX + 0.3, y: TOP + 0.3, w: 3.3, h: 0.35, fontSize: 15, color: C.white });
  T(s, "0개", { x: MX + 0.3, y: TOP + 0.7, w: 3.3, h: 1.1, fontSize: 60, bold: true, color: C.white, valign: "middle" });
  T(s, [
    { text: "3개월간 들어온 서로 다른 검색어", options: { breakLine: true } },
    { text: "7,602개", options: { bold: true, fontSize: 26 } },
  ], { x: MX + 0.3, y: TOP + 1.95, w: 3.3, h: 0.9, fontSize: 14, color: C.white });
  T(s, "통영숙수 전복죽 캠페인에 타사 전복죽 브랜드(여수해적, 장우준 등), '닭죽', '떡볶이 밀키트'까지 유입", { x: MX + 0.3, y: TOP + 3.0, w: 3.3, h: 1.1, fontSize: 13, color: C.white });

  const rx = 4.8, rw = W - MX - rx;
  const c = (t, o = {}) => ({ text: t, options: { fontSize: 14, ...o } });
  s.addTable([
    tableHead(["새는 곳", "3개월 비용", "내용"]),
    [c("GFA 트래픽 목적 캠페인"), c("91만 원", { align: "right" }), c("ROAS 59%, 구매 16건")],
    [c("GFA 19세 미만 노출"), c("57만 원", { align: "right" }), c("ROAS 143%")],
    [c("GFA 스마트채널·커뮤니케이션 지면"), c("32만 원", { align: "right" }), c("ROAS 75% · 0%")],
    [c("SA 콘텐츠 매체"), c("22만 원", { align: "right" }), c("클릭률 0.36%, ROAS 113%")],
    [c("SA 구매 0건 검색어"), c("15만 원", { align: "right" }), c("395개 (측정 누락 캠페인 제외)")],
    [c("합계", { bold: true, fill: { color: "FDE4DF" } }), c("약 217만 원", { align: "right", bold: true, color: C.brand, fill: { color: "FDE4DF" } }), c("월 약 70만 원 재배치 가능", { bold: true, color: C.brand, fill: { color: "FDE4DF" } })],
  ], { x: rx, y: TOP, w: rw, colW: [3.2, 1.35, 3.38], rowH: 0.5, border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0.04, 0.12, 0.04, 0.12] });
  card(s, MX, 6.1, CW, 0.5);
  T(s, "GFA 소재 87개 중 51개 구매 0건, 8~9월 신규 소재 없음  ·  통영숙수 전복죽 15일, 물회 39일 만에 종료 (판단 데이터가 쌓이기 전 중단)", { x: MX + 0.25, y: 6.1, w: CW - 0.5, h: 0.5, fontSize: 13, color: C.sub, valign: "middle" });
  s.addNotes(
    "3개월 동안 서로 다른 검색어가 7,602개 들어왔는데, 등록된 제외키워드는 하나도 없습니다.\n" +
    "예를 들어 통영숙수 전복죽 캠페인에는 다른 회사 전복죽 브랜드, '닭죽', '떡볶이 밀키트' 같은 검색어까지 들어왔습니다.\n" +
    "이렇게 새는 곳을 모아 보면 3개월 약 217만 원, 한 달에 약 70만 원입니다. 이 돈은 정리만 해도 성과가 나는 곳으로 옮길 수 있습니다.\n" +
    "소재도 87개 중 51개가 구매 0건이고, 8~9월에는 새 소재가 없었습니다. 통영숙수 전복죽은 15일, 물회는 39일 만에 꺼서, 판단할 데이터가 쌓이기 전에 멈췄습니다."
  );
}

// =====================================================================
// 섹션 C. GFA
// =====================================================================
divider(SEC.C, "03", "GFA", "늘린 예산과 쪼개진 캠페인", "증액 후 효율 하락  ·  ADVoost 분할  ·  새 고객과 기존 고객 혼재",
  "다음은 광고비의 88%를 차지하는 GFA입니다.");

// --- 10. 증액 후 하락
{
  const s = content(SEC.C, "예산을 2.3배 늘린 뒤 ROAS가 4개월 연속 떨어졌습니다", "03  GFA");
  s.addChart(pres.charts.BAR, [
    { name: "ROAS", labels: ["5~6월", "7월", "8월", "9월"], values: [364, 321, 312, 241] },
  ], {
    x: MX - 0.1, y: TOP, w: 6.9, h: 4.2, barDir: "col", barGapWidthPct: 55,
    chartColors: [HEX.navy, HEX.navy, HEX.navy, HEX.brand],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.text, dataLabelFormatCode: '0"%"',
    catAxisLabelColor: HEX.text, catAxisLabelFontSize: 14, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 420,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "GFA 월별 ROAS (구매완료 기준)", titleFontSize: 14, titleFontFace: "+mn-lt", titleColor: HEX.text,
  });
  T(s, "9월: 광고비는 8월과 비슷한데 노출은 절반, 구매는 32% 감소", { x: MX, y: 5.95, w: 6.6, h: 0.3, fontSize: 13, bold: true, color: C.brand });

  const rx = 7.35, rw = W - MX - rx;
  const cs = [
    ["월 광고비", "510만 → 1,171만 원", "2.3배"],
    ["CPM (노출 1,000회당 단가)", "4,377 → 11,990원", "2.7배"],
    ["늘린 돈의 효율 (추정)", "ROAS 239%", "추가 월 광고비 약 660만 원 → 추가 월 매출 약 1,580만 원"],
  ];
  cs.forEach(([l, n, d], i) => {
    const y = TOP + i * 1.45;
    card(s, rx, y, rw, 1.3, i === 2 ? C.navy : C.gray);
    const fg = i === 2 ? C.white : C.text;
    T(s, l, { x: rx + 0.3, y: y + 0.15, w: rw - 0.6, h: 0.3, fontSize: 13, color: i === 2 ? C.white : C.sub });
    T(s, n, { x: rx + 0.3, y: y + 0.45, w: rw - 1.5, h: 0.5, fontSize: 22, bold: true, color: fg, valign: "middle" });
    if (i < 2) T(s, d, { x: rx + rw - 1.3, y: y + 0.45, w: 1.0, h: 0.5, fontSize: 20, bold: true, color: C.brand, align: "right", valign: "middle" });
    else T(s, d, { x: rx + 0.3, y: y + 0.95, w: rw - 0.6, h: 0.3, fontSize: 12, color: C.white });
  });
  T(s, "5~6월은 5/1~9 중단을 반영해 5/10~6/30(52일) 실적을 월 환산. 7~9월 월 광고비는 3개월 단순 평균. 늘린 돈의 효율은 월 환산값으로 계산한 추정치", { x: rx, y: TOP + 4.4, w: rw, h: 0.6, fontSize: 11, color: C.cap });
  s.addNotes(
    "GFA는 7월부터 월 예산이 약 510만 원에서 1,171만 원으로 2.3배 늘었습니다. 그때부터 ROAS가 364%, 321%, 312%, 241%로 4개월 연속 떨어졌고, 한 번도 반등하지 않았습니다.\n" +
    "같은 기간 노출 1,000회당 단가인 CPM은 4,377원에서 11,990원으로 2.7배가 됐습니다. 9월은 광고비가 8월과 비슷한데 노출이 절반이고, 구매는 32% 줄었습니다.\n" +
    "늘린 돈만 떼어서 보면, 추가로 쓴 월 약 660만 원이 벌어온 추가 매출은 약 1,580만 원으로 ROAS 239% 정도로 추정됩니다.\n" +
    "단가가 왜 올랐는지는 추정입니다. 쪼개진 캠페인들이 같은 지면을 두고 단가만 올렸을 가능성이 있습니다.\n" +
    "반대로 말하면 5~6월에는 364%가 나오던 구간이 있었습니다. 거기서부터 다시 찾겠습니다."
  );
}

// --- 11. ADVoost 9개 분할
{
  const s = content(SEC.C, "AI가 찾아야 할 고객을 사람이 9개로 나눠 놓았습니다", "03  GFA");
  const c = (t, o = {}) => ({ text: t, options: { fontSize: 14, ...o } });
  const roas = (v) => c(v + "%", { align: "right", bold: v < 300, color: v < 300 ? C.brand : C.text });
  const tw = 3.3;
  const tbl = (x, title, rows) => {
    T(s, title, { x, y: TOP, w: tw, h: 0.4, fontSize: 16, bold: true, valign: "middle" });
    s.addTable([
      tableHead(["타겟", "광고비", "ROAS"]),
      ...rows.map(([t, w, r]) => [c(t), c(w, { align: "right" }), roas(r)]),
    ], { x, y: TOP + 0.5, w: tw, colW: [1.45, 1.0, 0.85], rowH: 0.45, border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0.03, 0.1, 0.03, 0.1] });
  };
  tbl(MX, "네이버 판매처 (4개)", [["25~39세", "672만", 332], ["꼬기밥 단독", "434만", 353], ["서브 연령대", "408만", 292], ["40~54세", "310만", 265]]);
  tbl(MX + tw + 0.3, "자사몰 (3개)", [["25~49세", "653만", 311], ["타겟 미표기", "480만", 325], ["서브 연령대", "9만", 343]]);
  card(s, MX, 4.6, 2 * tw + 0.3, 0.75);
  T(s, "삭제된 통영숙수 전복죽 ADVoost 2개 (네이버·자사몰): 209만 원, ROAS 156%", { x: MX + 0.25, y: 4.6, w: 2 * tw - 0.2, h: 0.75, fontSize: 14, color: C.sub, valign: "middle" });
  tint(s, MX, 5.5, 2 * tw + 0.3, 0.95, C.good, 88);
  T(s, [
    { text: "ADVoost 전체 3,174만 원 · ROAS 306%", options: { bold: true, breakLine: true } },
    { text: "나쁜 숫자가 아닙니다. 더 좋아질 여지가 막혀 있는 게 문제입니다", options: { fontSize: 13 } },
  ], { x: MX + 0.25, y: 5.5, w: 2 * tw - 0.2, h: 0.95, fontSize: 14, valign: "middle" });

  const rx = 7.75, rw = W - MX - rx;
  const pts = [
    ["합치면 결국 전 연령", "나눈 캠페인을 더하면 전 연령입니다. 학습할 구매 데이터만 나뉩니다"],
    ["보여주는 상품은 모두 같음", "연령을 나눴지만 상품은 전부 '전체상품'입니다"],
    ["경계가 효율과 안 맞음", "ROAS 426%인 45~49세 여성이 효율 최저(265%) 40~54세 캠페인에 섞여 있고, 연령 기준도 판매처마다 다릅니다"],
    ["서로 입찰할 수 있음", "꼬기밥 단독과 전체상품 캠페인의 상품이 겹쳐 같은 지면에서 경쟁할 수 있습니다"],
  ];
  pts.forEach(([h, d], i) => {
    const y = TOP + i * 1.2;
    badge(s, rx, y, i + 1, C.navy, 0.4, 13);
    T(s, h, { x: rx + 0.55, y, w: rw - 0.55, h: 0.4, fontSize: 16, bold: true, valign: "middle" });
    T(s, d, { x: rx + 0.55, y: y + 0.42, w: rw - 0.55, h: 0.75, fontSize: 13, color: C.sub });
  });
  s.addNotes(
    "ADVoost는 AI가 살 사람을 찾아주는 상품입니다. 그런데 지금은 판매처와 연령대로 9개 캠페인을 나눠 놓았습니다. 브랜드별로 나눈 것도 아닙니다. 꼬기밥 단독, 통영숙수 전복죽 정도만 예외이고 나머지는 전부 '전체상품'입니다.\n" +
    "네 가지가 걸립니다. 나눈 캠페인을 합치면 결국 전 연령이라, 학습할 구매 데이터만 쪼개집니다. 연령마다 보여주는 상품도 똑같습니다. " +
    "경계도 효율과 맞지 않습니다. ROAS 426%가 나오는 45~49세 여성이 효율이 가장 낮은 40~54세 캠페인 안에 섞여 있고, 연령 기준도 네이버는 25~39/40~54세, 자사몰은 25~49세로 다릅니다. " +
    "그리고 꼬기밥 단독 캠페인과 전체상품 캠페인이 같은 상품으로 같은 지면에서 서로 입찰할 수 있습니다.\n" +
    "ROAS 306%는 나쁜 숫자가 아닙니다. 더 좋아질 여지가 막혀 있고, 판단할 비교군이 없다는 게 문제입니다."
  );
}

// --- 12. 역할 혼재
{
  const s = content(SEC.C, "새 고객용 광고에 기존 고객이 섞여 있습니다", "03  GFA");
  const iw = 4.4, ih = iw * 383 / 480;
  s.addImage({ path: A("cap_signal.png"), x: MX, y: TOP, w: iw, h: ih, altText: "자사몰 ADVoost 잠재고객 시그널 화면" });
  s.addShape(pres.shapes.RECTANGLE, { x: MX, y: TOP, w: iw, h: ih, fill: { type: "none" }, line: { color: "D1D6DB", width: 0.75 } });
  const bx = MX + (68 / 480) * iw, by = TOP + (284 / 383) * ih;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w: (200 / 480) * iw, h: (30 / 383) * ih, rectRadius: 0.05, fill: { type: "none" }, line: { color: C.brand, width: 2.25 } });
  T(s, "기존 구매자", { x: bx + (210 / 480) * iw, y: by, w: 1.4, h: (30 / 383) * ih, fontSize: 13, bold: true, color: C.brand, valign: "middle" });
  T(s, "자사몰 ADVoost 시그널(AI에게 주는 힌트): 방문 60일 · 장바구니 60일 · 구매 30일", { x: MX, y: TOP + ih + 0.1, w: iw + 0.4, h: 0.55, fontSize: 12, color: C.sub });
  tint(s, MX, 5.75, iw + 0.4, 0.75, C.brand, 90);
  T(s, [
    { text: "라이브 중인 카탈로그 광고 0개", options: { bold: true, color: C.brand, breakLine: true } },
    { text: "구매 직전 고객만 다시 부르는 전용 광고가 없습니다", options: { fontSize: 13 } },
  ], { x: MX + 0.25, y: 5.75, w: iw - 0.1, h: 0.75, fontSize: 14, valign: "middle" });

  const rx = 5.85, rw = W - MX - rx;
  T(s, "광고별 역할을 나눕니다", { x: rx, y: TOP, w: rw, h: 0.4, fontSize: 16, bold: true, valign: "middle" });
  const c = (t, o = {}) => ({ text: t, options: { fontSize: 14, ...o } });
  s.addTable([
    tableHead(["광고", "역할", "대상"]),
    [c("ADVoost", { bold: true }), c("새 고객 획득"), c("판매처별 2개로 통합, 연령 분할 해제, 시그널에서 구매자 제외")],
    [c("카탈로그 (신규)", { bold: true }), c("구매 직전 고객 회수"), c("자사몰 방문 60일, 장바구니 60일")],
    [c("재구매 캠페인", { bold: true }), c("기존 고객 재구매"), c("구매자 대상, 닭가슴살 재구매 주기에 맞춤")],
  ], { x: rx, y: TOP + 0.5, w: rw, colW: [1.95, 1.85, rw - 3.8], rowH: [0.5, 0.85, 0.75, 0.75], border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0.05, 0.12, 0.05, 0.12] });
  tint(s, rx, 5.3, rw, 1.2, C.navy, 92);
  T(s, [
    { text: "지금은 기존 구매자를 학습 출발점으로 준 상태", options: { bold: true, breakLine: true } },
    { text: "ROAS 306% 중 새 고객 매출이 얼마인지 알 수 없습니다. 역할을 나누면 신규 ROAS와 리타겟 ROAS를 따로 볼 수 있습니다", options: { fontSize: 13, color: C.sub } },
  ], { x: rx + 0.3, y: 5.3, w: rw - 0.6, h: 1.2, fontSize: 15, valign: "middle" });
  s.addNotes(
    "자사몰 ADVoost의 잠재고객 시그널 화면입니다. 시그널은 AI에게 '이런 사람과 비슷한 사람을 찾아라' 하고 주는 힌트입니다.\n" +
    "자사몰 방문 60일, 장바구니 60일, 그리고 구매 30일이 들어가 있습니다. 즉 기존 구매자를 학습 출발점으로 준 상태입니다.\n" +
    "반면 지금 라이브 중인 카탈로그 광고는 0개라, 구매 직전 고객만 다시 부르는 전용 광고가 없습니다. 새 고객용 광고 하나가 새 고객, 구매 직전 고객, 기존 고객을 모두 맡고 있는 셈입니다.\n" +
    "그래서 ROAS 306% 중 새 고객이 만든 매출이 얼마인지 알 수 없습니다.\n" +
    "역할을 나누겠습니다. ADVoost는 새 고객 획득에 집중하고, 카탈로그 광고를 새로 만들어 구매 직전 고객을 회수하고, 재구매 캠페인으로 기존 고객을 따로 관리합니다."
  );
}

// =====================================================================
// 섹션 D. 개선 제안
// =====================================================================
divider(SEC.D, "04", "개선 제안", "순서대로 실행하는 개선 계획", "실행 방향  ·  실행 계획  ·  성과 지표",
  "이제 어떻게 바꿀지 말씀드리겠습니다.");

// --- 13. 3가지 방향
{
  const s = content(SEC.D, "보이게 만들고, 새는 곳을 막고, 키웁니다", "04  개선 제안");
  const cols = [
    ["보이게 만든다", C.navy, [
      "올라잇그린·건담닭 자사몰 전환 스크립트 설치",
      "계약금을 포함한 실제 SA ROAS로 보고",
      "브랜드/비브랜드, 신규/리타겟 매출 분리",
    ]],
    ["새는 곳을 막는다", C.brand, [
      "브랜드검색 노출 정상화",
      "브랜드검색 노출 기간엔 파워링크 브랜드 입찰 하향",
      "제외키워드 등록, 콘텐츠 매체·저효율 지면·19세 미만 제외",
      "올라잇그린: 제품군 키워드 분리, 문구에 제품군과 혜택, '비오틴'·'블루'·'공구' 제외",
    ]],
    ["키운다", C.good, [
      "ADVoost 통합 후 5~6월 효율 구간에서 단계적 증액",
      "카탈로그·재구매 캠페인 신설",
      "꼬기다 카테고리 키워드, 타사 브랜드 키워드 테스트",
      "핵심 타겟 수동 전환 캠페인 (GFA 예산의 20%)",
    ]],
  ];
  const gap = 0.35, cw = (CW - 2 * gap) / 3;
  cols.forEach(([h, col, items], i) => {
    const x = MX + i * (cw + gap);
    card(s, x, TOP, cw, 3.95);
    badge(s, x + 0.3, TOP + 0.3, i + 1, col, 0.5, 16);
    T(s, h, { x: x + 0.95, y: TOP + 0.3, w: cw - 1.1, h: 0.5, fontSize: 20, bold: true, valign: "middle" });
    T(s, items.map((t, k) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: k < items.length - 1, paraSpaceAfter: 8 } })),
      { x: x + 0.3, y: TOP + 1.05, w: cw - 0.55, h: 2.8, fontSize: 14, color: C.text });
    if (i < 2) T(s, "›", { x: x + cw, y: TOP + 1.7, w: gap, h: 0.6, fontSize: 28, bold: true, color: C.cap, align: "center", valign: "middle" });
  });
  card(s, MX, TOP + 4.15, CW, 0.7, C.navy);
  T(s, "순서가 중요합니다. 보이지 않는 상태에서 키우면 같은 결과가 반복됩니다", { x: MX + 0.3, y: TOP + 4.15, w: CW - 0.6, h: 0.7, fontSize: 16, bold: true, color: C.white, valign: "middle" });
  s.addNotes(
    "개선 방향은 세 가지이고, 순서가 중요합니다.\n" +
    "첫째, 보이게 만듭니다. 올라잇그린과 건담닭 자사몰에 전환 스크립트를 설치하고, 계약금까지 넣은 실제 ROAS로 보고, 브랜드와 비브랜드, 신규와 리타겟 매출을 나눠 봅니다.\n" +
    "둘째, 새는 곳을 막습니다. 브랜드검색 노출을 정상화하고, 노출되는 동안에는 파워링크 브랜드 입찰을 낮춰 이중 지출을 줄입니다. 제외키워드와 저효율 지면을 정리하고, 올라잇그린은 제품군 키워드와 문구를 정비합니다.\n" +
    "셋째, 키웁니다. ADVoost를 통합한 뒤 5~6월 효율이 나오던 구간에서부터 단계적으로 늘리고, 카탈로그와 재구매 캠페인을 새로 만들고, 카테고리 키워드와 타사 브랜드 키워드를 테스트합니다.\n" +
    "보이지 않는 상태에서 키우면 다시 같은 결과가 나옵니다. 그래서 보이게 만드는 것부터 시작합니다."
  );
}

// --- 14. 실행 계획
{
  const s = content(SEC.D, "이번 주부터 바로 시작할 수 있습니다", "04  개선 제안");
  const list = (arr, o = {}) => ({
    text: arr.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < arr.length - 1 } })),
    options: { fontSize: 13, valign: "middle", ...o },
  });
  const when = (t, col) => ({ text: t, options: { bold: true, fontSize: 15, color: C.white, fill: { color: col }, align: "center", valign: "middle" } });
  s.addTable([
    tableHead(["시기", "실행 항목", "확인 기준"]),
    [when("이번 주", HEX.brand),
      list(["브랜드검색 미노출 원인 확인·정상화", "올라잇그린·건담닭 자사몰 전환 스크립트 설치", "전체 제외키워드 등록", "콘텐츠 매체·저효율 지면·19세 미만 제외"]),
      list(["브랜드검색 일 노출 정상화", "파워링크 구매 집계 시작"], { color: C.sub })],
    [when("1개월", HEX.navy),
      list(["ADVoost 9개 → 2개 통합 (2주 비교)", "카탈로그 리타겟 신규 세팅", "올라잇그린 키워드·문구 정비", "카테고리·타사 브랜드 키워드 테스트", "브랜드검색 기간 파워링크 입찰 하향", "트래픽 캠페인은 전환 목적으로 통합, 신규 소재 월 2~3개"]),
      list(["통합 ADVoost ROAS·CPM 유지 또는 개선", "올라잇그린 클릭당 비용 하락", "신규 키워드 구매당 비용"], { color: C.sub })],
    [when("2~3개월", HEX.slate),
      list(["ADVoost 주간 예산 단계 테스트 (5~6월 수준에서 출발)", "재구매 캠페인, 핵심 타겟 수동 전환 캠페인 20%", "시즌 상품 최소 4주 운영", "분리 리포트 정착"]),
      list(["ROAS가 꺾이는 예산 구간 확인", "신규 고객 매출 비중 증가"], { color: C.sub })],
  ], { x: MX, y: TOP, w: CW, colW: [1.5, 6.5, 4.13], border: { type: "solid", pt: 0.75, color: HEX.line }, margin: [0.08, 0.15, 0.08, 0.15] });
  s.addNotes(
    "실행은 세 단계입니다.\n" +
    "이번 주에는 돈이 새는 곳과 측정을 먼저 막습니다. 브랜드검색 미노출 원인을 확인해 정상화하고, 올라잇그린과 건담닭 자사몰에 전환 스크립트를 설치하고, 제외키워드와 저효율 지면을 정리합니다. 확인 기준은 브랜드검색이 매일 노출되는지, 파워링크 구매가 잡히기 시작하는지입니다.\n" +
    "1개월 안에는 구조를 바꿉니다. ADVoost를 9개에서 2개로 합치고 2주간 비교합니다. 카탈로그 리타겟을 새로 세팅하고, 올라잇그린 키워드와 문구, 카테고리 키워드 테스트를 시작합니다.\n" +
    "2~3개월에는 키웁니다. 5~6월 예산 수준에서 출발해 주 단위로 예산을 올려 보면서 ROAS가 꺾이는 지점을 찾습니다. 시즌 상품은 최소 4주는 운영해서 판단할 데이터를 쌓습니다."
  );
}

// --- 15. KPI
{
  const s = content(SEC.D, "하나의 ROAS가 아니라 다섯 개의 숫자로 봅니다", "04  개선 제안");
  const kpis = [
    ["계약금 포함 SA ROAS", "브랜드검색 계약금까지 넣은 실제 검색광고 효율"],
    ["비브랜드 검색 매출 · 구매당 비용", "새 고객을 데려오는 검색의 성과"],
    ["GFA 신규 ROAS / 리타겟 ROAS", "새 고객과 기존 고객의 효율을 따로"],
    ["GFA CPM (노출 단가)", "늘린 예산이 단가만 올리고 있지 않은지"],
    ["측정 누락 캠페인 0개", "모든 광고비의 성과가 보이는 상태"],
  ];
  kpis.forEach(([h, d], i) => {
    const y = TOP + i * 0.97;
    card(s, MX, y, 6.75, 0.85);
    badge(s, MX + 0.22, y + 0.2, i + 1, C.navy, 0.45, 14);
    T(s, h, { x: MX + 0.85, y: y + 0.08, w: 5.7, h: 0.38, fontSize: 16, bold: true, valign: "middle" });
    T(s, d, { x: MX + 0.85, y: y + 0.46, w: 5.7, h: 0.32, fontSize: 13, color: C.sub, valign: "middle" });
  });
  const rx = 7.65, rw = W - MX - rx;
  card(s, rx, TOP, rw, 4.73, C.navy);
  T(s, "정리만으로", { x: rx + 0.35, y: TOP + 0.3, w: rw - 0.7, h: 0.35, fontSize: 15, color: C.white });
  T(s, "월 약 70만 원", { x: rx + 0.35, y: TOP + 0.65, w: rw - 0.7, h: 0.75, fontSize: 36, bold: true, color: C.white, valign: "middle" });
  T(s, "재배치 가능 (3개월 217만 원)", { x: rx + 0.35, y: TOP + 1.4, w: rw - 0.7, h: 0.35, fontSize: 14, color: C.white });
  const more = ["남은 브랜드검색 계약(10/25까지) 정상 노출", "올라잇그린 클릭당 비용이 내려갈 여지", "측정 누락 121만 원의 효과 판단 가능"];
  T(s, more.map((t, k) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: k < more.length - 1, paraSpaceAfter: 10 } })),
    { x: rx + 0.35, y: TOP + 2.1, w: rw - 0.7, h: 2.4, fontSize: 14, color: C.white });
  s.addNotes(
    "앞으로는 ROAS 하나가 아니라 다섯 개 숫자로 보시길 제안드립니다. 계약금을 포함한 검색광고 ROAS, 비브랜드 검색의 매출과 구매당 비용, GFA의 신규와 리타겟 ROAS, GFA 노출 단가, 그리고 측정이 안 되는 캠페인이 0개인지입니다.\n" +
    "효과를 몇 퍼센트 올리겠다고 약속드리기보다, 확인 가능한 것만 말씀드리겠습니다. 정리만으로 한 달 약 70만 원을 성과 나는 곳으로 옮길 수 있습니다. " +
    "10월 25일까지 남은 브랜드검색 계약을 정상 노출시키고, 올라잇그린 클릭당 비용이 내려갈 여지가 있고, 지금 측정이 안 되는 121만 원의 효과도 판단할 수 있게 됩니다."
  );
}

// --- 16. 필요한 것 (어두운 장)
{
  const s = pres.addSlide({ masterName: "NC_SECTION", sectionTitle: SEC.D });
  s.addText("진행을 위해 필요한 것", { placeholder: "title" });
  const items = [
    ["네이버 광고 계정 열람 권한", "브랜드검색 미노출 원인, ADVoost 전체 시그널 확인"],
    ["자사몰 스크립트 설치 담당자 연결", "올라잇그린·건담닭 전환 스크립트 설치"],
    ["꼬기다 자사몰 상품 카탈로그", "등록 여부 확인 (카탈로그 리타겟 준비)"],
    ["브랜드별 목표 ROAS 또는 마진 기준", "예산 증액 판단의 기준선"],
  ];
  const cw = 5.1, ch = 1.45;
  items.forEach(([h, d], i) => {
    const x = 1.3 + (i % 2) * (cw + 0.3), y = 1.85 + Math.floor(i / 2) * (ch + 0.3);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.12, fill: { color: C.white, transparency: 88 }, line: { color: C.white, width: 0.75, transparency: 60 } });
    badge(s, x + 0.3, y + 0.3, i + 1, C.brand, 0.45, 15);
    T(s, h, { x: x + 0.95, y: y + 0.28, w: cw - 1.15, h: 0.5, fontSize: 17, bold: true, color: C.white, valign: "middle" });
    T(s, d, { x: x + 0.95, y: y + 0.8, w: cw - 1.15, h: 0.45, fontSize: 14, color: C.white });
  });
  T(s, "NICECORE 퍼포먼스마케팅팀", { x: 1.3, y: 5.4, w: 6, h: 0.35, fontSize: 16, bold: true, color: C.white });
  T(s, "모든 전환 지표는 총전환이 아닌 구매완료 전환수·구매완료 전환매출액 기준", { x: 1.3, y: 5.75, w: 9, h: 0.3, fontSize: 12, color: C.white });
  s.addNotes(
    "바로 시작하기 위해 네 가지가 필요합니다.\n" +
    "첫째, 네이버 광고 계정 열람 권한입니다. 브랜드검색이 왜 꺼져 있었는지, ADVoost 시그널 전체가 어떻게 들어가 있는지 확인하겠습니다.\n" +
    "둘째, 올라잇그린과 건담닭 자사몰 스크립트를 설치해 주실 담당자 연결입니다.\n" +
    "셋째, 꼬기다 자사몰 상품 카탈로그가 등록돼 있는지 여부입니다. 카탈로그 리타겟 광고에 필요합니다.\n" +
    "넷째, 브랜드별 목표 ROAS나 마진 기준입니다. 예산을 어디까지 늘릴지 판단하는 기준이 됩니다.\n" +
    "말씀드린 모든 숫자는 구매완료 전환 기준입니다. 감사합니다."
  );
}

// =====================================================================
// 부록
// =====================================================================
pres.addSection({ title: SEC.X });
{
  const s = content(SEC.X, "데이터 기준", "부록");
  const items = [
    ["전환 지표", "모든 전환 지표는 구매완료 전환수·구매완료 전환매출액 기준. 총전환은 쓰지 않음"],
    ["재계산", "ROAS·CPA·CPC·클릭률·CPM은 비용·노출·클릭·구매·매출 원본으로 다시 계산"],
    ["분석 기간", "2026.07.01 ~ 09.30. GFA 추이만 5월부터"],
    ["브랜드검색", "계약금 방식이라 보고서 비용 0원. 계약금을 하루 단위로 나눠 분석 기간 해당분만 반영"],
    ["GFA 일별", "62일 조회 제한 때문에 7월과 8~9월을 나눠 받아 합침"],
    ["ADVoost", "애셋 그룹 보고서 기준. 캠페인당 애셋 그룹이 1개라 캠페인 성과와 같고, 소재 단위 성과는 없음"],
    ["5~6월 GFA", "5/1~9 중단을 반영해 5/10~6/30(52일) 기준으로 월 환산"],
    ["증액 효율", "7~9월 월 광고비·매출은 3개월 단순 평균. 늘린 돈의 효율(239%)은 추정치"],
  ];
  const cw = (CW - 0.3) / 2, ch = 1.08;
  items.forEach(([h, d], i) => {
    const x = MX + (i % 2) * (cw + 0.3), y = TOP + Math.floor(i / 2) * (ch + 0.15);
    card(s, x, y, cw, ch);
    T(s, h, { x: x + 0.25, y: y + 0.15, w: cw - 0.5, h: 0.32, fontSize: 14, bold: true, color: C.navy });
    T(s, d, { x: x + 0.25, y: y + 0.48, w: cw - 0.5, h: 0.55, fontSize: 13, color: C.sub });
  });
  s.addNotes("질문이 나오면 참고할 데이터 기준입니다. 모든 전환은 구매완료 기준이고, 브랜드검색 계약금은 하루 단위로 나눠 분석 기간만큼 반영했습니다.");
}

// 마지막 장
{
  const s = pres.addSlide({ masterName: "NC_COVER", sectionTitle: SEC.X });
  T(s, "THANK YOU", { x: 1.3, y: 2.6, w: 10.7, h: 0.9, fontSize: 44, bold: true, color: C.white, align: "center", valign: "middle" });
  T(s, "언제나 최상의 퍼포먼스로 클라이언트의 성장이라는 목표를 함께 하겠습니다.", { x: 1.3, y: 3.5, w: 10.7, h: 0.4, fontSize: 14, color: C.white, align: "center" });
  s.addNotes("감사합니다. 질문 받겠습니다.");
}

(async () => {
  const out = path.join(__dirname, OUT);
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("written", out);
})();
