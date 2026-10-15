// 통영숙수 광고 운영 제안서 생성 스크립트
// 실행: node build.js  →  통영숙수_광고운영제안서.pptx
// 문구·숫자 기준: 통영숙수_PPT_슬라이드원고.md
// 디자인: NICECORE 템플릿(nicecore_proposal_pt12) 배경·로고·SUIT 폰트, 빨간색 대신 어두운 하늘색 강조

const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply_theme.js");

const OUT = "통영숙수_광고운영제안서.pptx";
const A = (f) => path.join(__dirname, "assets", f);

const HEX = {
  text: "191F28", white: "FFFFFF", slateDark: "2E3A4E", gray: "F2F4F6",
  accent: "1568A3", // 어두운 하늘색
  sub: "4E5968", cap: "8B95A1", warn: "F08A00", good: "03B26C", slate: "6F7D97",
  line: "D1D6DB", tint: "E3EFF8", warnTint: "FDF0DD",
};
const THEME = {
  name: "NICECORE Sky",
  headFontFace: "SUIT",
  bodyFontFace: "SUIT",
  colors: {
    dk1: HEX.text, lt1: HEX.white, dk2: HEX.slateDark, lt2: HEX.gray,
    accent1: HEX.accent, accent2: HEX.sub, accent3: HEX.cap,
    accent4: HEX.good, accent5: HEX.warn, accent6: HEX.slate,
    hlink: HEX.accent, folHlink: HEX.slate,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "NICECORE 퍼포먼스마케팅팀";
pres.company = "NICECORE";
pres.title = "통영숙수 광고 운영 제안서";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

const S = pres.SchemeColor;
const C = {
  text: S.text1, white: S.background1, dark: S.text2, gray: S.background2,
  accent: S.accent1, sub: S.accent2, cap: S.accent3, good: S.accent4, warn: S.accent5, slate: S.accent6,
};

const W = 13.333, MX = 0.6, CW = W - MX * 2, TOP = 1.65;

// ---------- 레이아웃 ----------
const darkFooter = [
  { image: { path: A("logo_white.png"), x: 11.1, y: 6.62, w: 1.7, h: 0.25 } },
  { text: { text: "Credible Outcome Reliable Evolution", options: { x: 11.1, y: 6.95, w: 2.0, h: 0.2, fontSize: 8, color: C.white, margin: 0 } } },
  { text: { text: "©2026 NICECORE. ALL RIGHTS RESERVED.", options: { x: 0.25, y: 7.1, w: 3.5, h: 0.2, fontSize: 7, color: C.white, margin: 0 } } },
];
pres.defineSlideMaster({ title: "NC_DARK", background: { path: A("bg.jpg") }, objects: darkFooter });
pres.defineSlideMaster({
  title: "NC_CONTENT",
  background: { color: C.white },
  objects: [
    { image: { path: A("logo_dark.png"), x: 11.43, y: 6.98, w: 1.3, h: 0.19 } },
    { text: { text: "©2026 NICECORE. ALL RIGHTS RESERVED.", options: { x: MX, y: 7.0, w: 4, h: 0.18, fontSize: 8, color: C.cap, margin: 0 } } },
    { placeholder: { options: { name: "kicker", type: "body", x: MX, y: 0.38, w: 8, h: 0.3, fontSize: 12, bold: true, color: C.accent, valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.72, w: CW, h: 0.7, fontSize: 28, bold: true, color: C.text, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 10.9, y: 6.98, w: 0.4, h: 0.2, fontSize: 8, color: C.cap, align: "right" },
});

// ---------- 도우미 ----------
const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, fontSize: 14, color: C.text, valign: "top", ...o });
const card = (s, x, y, w, h, fill = C.gray, o = {}) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill, ...(o.transparency ? { transparency: o.transparency } : {}) }, line: o.line || { type: "none" }, ...(o.name ? { objectName: o.name } : {}) });
const badge = (s, x, y, n, fill = C.accent, d = 0.42, size = 14) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  T(s, String(n), { x, y, w: d, h: d, fontSize: size, bold: true, color: C.white, align: "center", valign: "middle" });
};
const bullets = (items, o = {}) => items.map((t, k) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: k < items.length - 1, ...o } }));
const head = (cells, size = 14) => cells.map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: HEX.slateDark }, fontSize: size } }));
const cell = (size) => (t, o = {}) => ({ text: t, options: { fontSize: size, color: C.text, ...o } });
const TB = (o = {}) => ({ border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: [0.05, 0.12, 0.05, 0.12], ...o });

const SEC = { I: "Ⅰ 현황", II: "Ⅱ 전략", III: "Ⅲ 운영안", IV: "Ⅳ 예산·일정" };
const content = (sec, title, kicker = sec) => {
  const s = pres.addSlide({ masterName: "NC_CONTENT", sectionTitle: sec });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
};

// =====================================================================
// 1. 표지
// =====================================================================
pres.addSection({ title: "표지" });
{
  const s = pres.addSlide({ masterName: "NC_DARK", sectionTitle: "표지" });
  T(s, "NICECORE 광고 운영 제안", { x: 1.3, y: 1.25, w: 8, h: 0.35, fontSize: 16, bold: true, color: C.white });
  T(s, "통영숙수 광고 운영 제안서", { x: 1.3, y: 1.7, w: 10, h: 0.85, fontSize: 40, bold: true, color: C.white, valign: "middle" });
  T(s, "전복죽 집중으로 BEP부터, 셰프·리타겟·재구매로 450%까지", { x: 1.3, y: 2.7, w: 10, h: 0.45, fontSize: 20, color: C.white });
  T(s, "2026년 10월 19일  ·  NICECORE 퍼포먼스마케팅팀 → 브레이브컴퍼니", { x: 1.3, y: 3.35, w: 10, h: 0.35, fontSize: 14, color: C.white });
  s.addNotes("오늘은 통영숙수 7월부터 10월 중순까지의 광고·판매 데이터를 바탕으로, 4분기 운영안과 내년 상반기까지의 방향을 말씀드리겠습니다.");
}

// 목차
{
  const s = pres.addSlide({ masterName: "NC_DARK", sectionTitle: "표지" });
  T(s, "CONTENTS", { x: 1.16, y: 1.13, w: 5, h: 0.5, fontSize: 32, bold: true, color: C.white, valign: "middle" });
  T(s, "목차", { x: 1.16, y: 1.68, w: 2, h: 0.36, fontSize: 19, color: C.white });
  s.addShape(pres.shapes.LINE, { x: 1.16, y: 2.16, w: 0.46, h: 0, line: { color: "8CC8EE", width: 2 } });
  const items = [
    ["01", "현황", "목표와 현재 · 스마트스토어 · 판매 비중 · 핵심 문제"],
    ["02", "전략", "포지셔닝 · 상품별 역할 · 두 대상"],
    ["03", "운영안", "착지와 측정 · 매체별 캠페인 · 메시지 · 소재 · 객단가"],
    ["04", "예산·일정", "목표까지의 경로 · 4분기 예산 · 착수 후 8주 · 로드맵"],
  ];
  items.forEach(([n, t, d], i) => {
    const y = 1.28 + i * 1.42;
    T(s, n, { x: 7.79, y, w: 0.65, h: 0.42, fontSize: 26, bold: true, color: C.white, valign: "middle" });
    T(s, t, { x: 8.5, y, w: 4, h: 0.42, fontSize: 22, bold: true, color: C.white, valign: "middle" });
    T(s, d, { x: 8.5, y: y + 0.47, w: 4.1, h: 0.3, fontSize: 11, color: C.white });
    s.addShape(pres.shapes.LINE, { x: 7.79, y: y + 0.82, w: 4.75, h: 0, line: { color: "C4CED8", width: 1 } });
  });
  s.addNotes("순서는 네 부분입니다. 현황은 짧게 보고, 전략과 운영안, 예산과 일정을 중심으로 말씀드리겠습니다.");
}

// =====================================================================
// 2. 제안 요약
// =====================================================================
pres.addSection({ title: "요약" });
{
  const s = pres.addSlide({ masterName: "NC_CONTENT", sectionTitle: "요약" });
  s.addText("제안 요약", { placeholder: "kicker" });
  // 원고 제목이 길어 이 장만 제목 상자를 직접 그림 (같은 위치, 26pt)
  T(s, "전복죽에 집중해 BEP부터 회복하고, 450%까지 단계적으로 올립니다", { x: MX, y: 0.72, w: CW, h: 0.7, fontSize: 26, bold: true, valign: "middle" });
  const cards = [
    ["전복죽 집중", "8월 이후 매출 95%. 4분기 예산 대부분을 전복죽에, 쭈꾸미·생선구이는 소액 테스트"],
    ["전환 품질 회복", "클릭만 만드는 노출을 걷어내고, 셰프 검색·쇼핑검색·리타겟에 집중"],
    ["객단가 올리기", "무료배송 5만 원 = 전복죽 2팩. \"2팩부터 무료배송\""],
    ["다시 부르기", "카탈로그·카카오 리타겟·카카오톡 채널로 재구매"],
  ];
  const cw = (CW - 3 * 0.25) / 4;
  cards.forEach(([h, d], i) => {
    const x = MX + i * (cw + 0.25);
    card(s, x, TOP, cw, 3.15);
    badge(s, x + 0.3, TOP + 0.3, i + 1);
    T(s, h, { x: x + 0.3, y: TOP + 0.95, w: cw - 0.55, h: 0.45, fontSize: 18, bold: true });
    T(s, d, { x: x + 0.3, y: TOP + 1.5, w: cw - 0.55, h: 1.5, fontSize: 14, color: C.sub });
  });
  card(s, MX, 5.05, CW, 1.45, C.accent);
  T(s, [
    { text: "예산  ", options: { bold: true } },
    { text: "월 300만 원으로 시작 → BEP 근접 시 12월 400만 원", options: { breakLine: true } },
    { text: "경로  ", options: { bold: true } },
    { text: "12월 BEP 근접 → 2027년 1분기 BEP 이상 → 2분기 450%" },
  ], { x: MX + 0.4, y: 5.05, w: CW - 0.8, h: 1.45, fontSize: 18, color: C.white, valign: "middle", paraSpaceAfter: 6 });
  s.addNotes("결론부터 말씀드리면, 지금은 예산을 늘릴 때가 아니라 효율을 바로잡을 때입니다. 네 가지를 바꾸면 단계적으로 목표에 갈 수 있습니다.");
}

// =====================================================================
// Ⅰ 현황
// =====================================================================
pres.addSection({ title: SEC.I });

// 3. 목표와 현재
{
  const s = content(SEC.I, "현재 ROAS는 BEP의 절반 수준입니다", "Ⅰ 현황 · 목표와 현재");
  const labels = ["3분기 (7~9월)", "10월 1~14일"];
  // 콤보 차트(막대+선)는 PowerPoint 복구 오류를 내므로, 막대 차트 + 도형 기준선으로 그림
  const ch = { x: MX - 0.1, y: TOP, w: 6.1, h: 4.6 };
  const pa = { y: 0.06, h: 0.82 }; // 플롯 영역(차트 높이 비율), 값 축 0~500
  s.addChart(pres.charts.BAR, [{ name: "ROAS", labels, values: [152, 82] }], {
    ...ch, barDir: "col", barGapWidthPct: 90, chartColors: [HEX.accent],
    layout: { x: 0.04, y: pa.y, w: 0.92, h: pa.h },
    showValue: true, dataLabelFontSize: 13, dataLabelFontBold: true, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.text, dataLabelFormatCode: '0"%"', dataLabelPosition: "outEnd",
    catAxisLabelColor: HEX.text, catAxisLabelFontSize: 14, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 500, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  });
  const vy = (v) => ch.y + ch.h * (pa.y + pa.h * (1 - v / 500));
  [[300, "BEP 300%", HEX.warn, C.warn], [450, "목표 450%", HEX.accent, C.accent]].forEach(([v, l, hex, col]) => {
    s.addShape(pres.shapes.LINE, { x: ch.x + 0.25, y: vy(v), w: ch.w - 0.5, h: 0, line: { color: hex, width: 2, dashType: "dash" } });
    T(s, l, { x: ch.x + ch.w - 1.6, y: vy(v) - 0.32, w: 1.35, h: 0.28, fontSize: 12, bold: true, color: col, align: "right" });
  });
  const rx = 6.8, rw = W - MX - rx;
  const c = cell(12);
  s.addTable([
    head(["구분", "광고비", "구매", "구매 매출", "ROAS"], 12),
    [c("3분기"), c("약 306만 원", { align: "right" }), c("122건", { align: "right" }), c("약 466만 원", { align: "right" }), c("152%", { align: "right", bold: true })],
    [c("10월 1~14일"), c("약 168만 원", { align: "right" }), c("33건", { align: "right" }), c("약 137만 원", { align: "right" }), c("82%", { align: "right", bold: true, color: C.warn })],
  ], TB({ x: rx, y: TOP + 0.1, w: rw, colW: [1.35, 1.3, 0.7, 1.3, rw - 4.65], rowH: 0.45, margin: [0.04, 0.08, 0.04, 0.08] }));
  card(s, rx, 3.45, rw, 1.55, HEX.warnTint);
  T(s, [
    { text: "10월은 2주 만에", options: { breakLine: true } },
    { text: "3분기 석 달 광고비의 절반 이상을 썼고,", options: { bold: true, breakLine: true } },
    { text: "효율은 절반으로 떨어졌습니다", options: { bold: true, color: C.warn } },
  ], { x: rx + 0.3, y: 3.45, w: rw - 0.6, h: 1.55, fontSize: 16, valign: "middle" });
  T(s, "통영숙수 캠페인만, 구매완료 기준. 목표·BEP는 광고주 제공", { x: MX, y: 6.4, w: 8, h: 0.3, fontSize: 11, color: C.cap });
  s.addNotes("10월은 2주 만에 3분기 석 달 광고비의 절반 이상을 썼는데, 효율은 절반으로 떨어졌습니다.");
}

// 4. 스마트스토어
{
  const s = content(SEC.I, "방문은 세 배 늘었지만, 결제는 1.4배에 그쳤습니다", "Ⅰ 현황 · 스마트스토어");
  const mini = (x, title, vals, fmt, mult) => {
    s.addChart(pres.charts.BAR, [{ name: title, labels: ["9월", "10월 1~14일"], values: vals }], {
      x, y: TOP, w: 3.2, h: 3.5, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.slate, HEX.accent],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.text, dataLabelFormatCode: fmt,
      catAxisLabelColor: HEX.text, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valAxisMinVal: 0,
      valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
      showTitle: true, title, titleFontSize: 14, titleFontFace: "+mn-lt", titleColor: HEX.text,
    });
    T(s, mult, { x, y: TOP + 3.55, w: 3.2, h: 0.35, fontSize: 15, bold: true, color: C.accent, align: "center" });
  };
  mini(MX - 0.1, "하루 평균 방문", [89, 258], '"약 "0"건"', "약 2.9배");
  mini(MX + 3.3, "하루 평균 결제", [4.3, 5.9], '"약 "0.0"건"', "약 1.4배");
  const rx = 7.5, rw = W - MX - rx;
  card(s, rx, TOP, rw, 3.9);
  T(s, "구매 전환율", { x: rx + 0.4, y: TOP + 0.35, w: rw - 0.8, h: 0.4, fontSize: 16, color: C.sub });
  T(s, [
    { text: "4.9%", options: { color: C.text } },
    { text: "  →  ", options: { color: C.cap } },
    { text: "2.3%", options: { color: C.warn } },
  ], { x: rx + 0.4, y: TOP + 0.85, w: rw - 0.6, h: 1.0, fontSize: 40, bold: true, valign: "middle" });
  T(s, "9월 → 10월 1~14일", { x: rx + 0.4, y: TOP + 2.05, w: rw - 0.8, h: 0.35, fontSize: 14, color: C.sub });
  T(s, "광고로 늘어난 방문 대부분이 사지 않는 사람이었습니다", { x: rx + 0.4, y: TOP + 2.75, w: rw - 0.8, h: 0.8, fontSize: 15, bold: true });
  card(s, MX, 5.75, CW, 0.65, HEX.tint);
  T(s, [
    { text: "방문 1위 경로  ", options: { bold: true, color: C.accent } },
    { text: "9월 네이버 통합검색 38.4%  →  10월 디스플레이 광고 74.9%" },
  ], { x: MX + 0.3, y: 5.75, w: CW - 0.6, h: 0.65, fontSize: 15, valign: "middle" });
  T(s, "스마트스토어 통계, 광고주 제공", { x: MX, y: 6.5, w: 6, h: 0.25, fontSize: 11, color: C.cap });
  s.addNotes("광고로 늘어난 방문 대부분이 사지 않는 사람이었다는 뜻입니다. 트래픽 목적 캠페인, 스마트채널 지면, 60세 이상 노출이 클릭만 만든 결과입니다.");
}

// 5. 판매 비중
{
  const s = content(SEC.I, "지금 통영숙수의 매출은 사실상 전복죽 하나입니다", "Ⅰ 현황 · 판매 비중");
  const labels = ["7월~8월 초", "8월~현재"];
  s.addChart(pres.charts.BAR, [
    { name: "전복죽", labels, values: [25, 95] },
    { name: "물회", labels, values: [70, 0] },
    { name: "기타·그 외", labels, values: [5, 5] },
  ], {
    x: MX - 0.1, y: TOP, w: CW + 0.2, h: 3.3, barDir: "bar", barGrouping: "percentStacked", barGapWidthPct: 45,
    catAxisOrientation: "maxMin", chartColors: [HEX.accent, HEX.slate, "C4CED8"],
    showValue: true, dataLabelPosition: "ctr", dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.white, dataLabelFormatCode: '0"%";;',
    catAxisLabelColor: HEX.text, catAxisLabelFontSize: 15, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendFontSize: 13, legendFontFace: "+mn-lt", legendColor: HEX.sub,
  });
  card(s, MX, 5.15, CW, 0.85, HEX.warnTint);
  T(s, [
    { text: "10월 광고비의 40% 이상이 쭈꾸미", options: { bold: true } },
    { text: "에 쓰였고, 쭈꾸미 ROAS는 " },
    { text: "35%", options: { bold: true, color: C.warn } },
    { text: "입니다" },
  ], { x: MX + 0.35, y: 5.15, w: CW - 0.7, h: 0.85, fontSize: 18, valign: "middle" });
  T(s, "스마트스토어·자사몰 동일 비율, 광고주 제공", { x: MX, y: 6.2, w: 6, h: 0.25, fontSize: 11, color: C.cap });
  s.addNotes("팔리는 상품과 광고비가 가는 상품이 어긋나 있습니다. 반대로 지난여름 물회는 매출의 70%를 낸 검증된 상품이라 내년 여름의 핵심 카드입니다.");
}

// 6. 핵심 문제
{
  const s = content(SEC.I, "효율을 깎는 원인 다섯 가지", "Ⅰ 현황 · 핵심 문제");
  const items = [
    ["사려는 사람보다 클릭하는 사람에게 돈을 씁니다", "상품·목적·연령으로 쪼갰지만 소재는 모두 같고, 구매는 ADVoost와 쇼핑검색에만 몰림"],
    ["셰프라는 자산을 쓰지 않습니다", "셰프 얼굴은 판매 종료된 물회 소재에만, 셰프명 검색광고 없음. 셰프명 검색은 거의 비용 없이 구매로 이어졌음"],
    ["팔리는 상품과 광고 비중이 맞지 않습니다", "매출 95%가 전복죽인데 광고비 40% 이상이 쭈꾸미"],
    ["브랜드는 프리미엄, 상품명은 가성비입니다", "광고에 그대로 노출되는 상품명에 '실속', 'BEST 42%', 전 상품 상시 할인"],
    ["다시 부르는 장치가 없습니다", "담고 안 산 고객이 많지만 카탈로그 리타겟·재구매 채널 없음"],
  ];
  items.forEach(([h, d], i) => {
    const y = TOP + i * 0.97;
    card(s, MX, y, CW, 0.85);
    badge(s, MX + 0.22, y + 0.2, i + 1);
    T(s, h, { x: MX + 0.85, y, w: 4.9, h: 0.85, fontSize: 16, bold: true, valign: "middle" });
    T(s, d, { x: MX + 5.9, y, w: CW - 6.15, h: 0.85, fontSize: 14, color: C.sub, valign: "middle" });
  });
  s.addNotes("다섯 가지 모두 예산을 늘리지 않고도 고칠 수 있는 문제입니다.");
}

// =====================================================================
// Ⅱ 전략
// =====================================================================
pres.addSection({ title: SEC.II });

// 7. 포지셔닝
{
  const s = content(SEC.II, "셰프 얼굴이 아니라, 셰프의 '고향'으로 차별화합니다", "Ⅱ 전략 · 포지셔닝");
  card(s, MX, TOP, CW, 1.45, C.accent);
  T(s, "\"통영에서 나고 자란 장호준 셰프가, 고향의 맛을 기준으로 직접 개발한 한 끼\"", { x: MX + 0.5, y: TOP, w: CW - 1.0, h: 1.45, fontSize: 21, bold: true, color: C.white, valign: "middle" });
  T(s, "통영숙수만 가진 것", { x: MX, y: 3.35, w: 5, h: 0.35, fontSize: 16, bold: true });
  const cards = [
    ["통영 출신 셰프라는 진정성", ""],
    ["백화점 팝업으로 검증된 격", ""],
    ["상품별로 분명한 원물 강점", "완도산 전복 3마리분, 48시간 숙성"],
  ];
  const cw = (CW - 2 * 0.25) / 3;
  cards.forEach(([h, d], i) => {
    const x = MX + i * (cw + 0.25);
    card(s, x, 3.8, cw, 1.5);
    badge(s, x + 0.3, 3.8 + 0.3, i + 1);
    T(s, h, { x: x + 0.9, y: 3.8 + 0.3, w: cw - 1.1, h: 0.42, fontSize: 16, bold: true, valign: "middle" });
    if (d) T(s, d, { x: x + 0.9, y: 3.8 + 0.8, w: cw - 1.1, h: 0.55, fontSize: 14, color: C.sub });
  });
  T(s, "가격은 대기업과, 리뷰 수는 맛집 브랜드와, 셰프 이름은 다른 셰프 브랜드와 겹칩니다. 겹치지 않는 건 위의 세 가지입니다.", { x: MX, y: 5.55, w: CW, h: 0.7, fontSize: 15, color: C.sub, valign: "middle" });
  s.addNotes("셰프 협업 간편식이 쏟아지면서 셰프 이름만으로는 구분되지 않습니다. 통영 출신이라는 이야기가 핵심입니다.");
}

// 8. 상품별 역할
{
  const s = content(SEC.II, "4분기는 전복죽, 내년은 시즌 상품으로 이어갑니다", "Ⅱ 전략 · 상품별 역할");
  const c = cell(14);
  const hl = { fill: { color: HEX.tint } };
  s.addTable([
    head(["상품", "역할", "핵심 사실", "이길 포인트"]),
    [c("게우 통전복죽", { bold: true, ...hl }), c("4분기 주력 (매출 95%)", { bold: true, color: C.accent, ...hl }), c("완도산 전복 3마리분과 내장, 밥솥 취사 13분, 1팩 2인분", hl), c("내장으로 낸 초록빛, 식당 전복죽의 진한 맛", hl)],
    [c("생선구이 (고등어·삼치)", { bold: true }), c("소액 테스트"), c("1팩 280g, 가시 99.9% 제거, 48시간 숙성"), c("속살까지 밴 간, 셰프 비법")],
    [c("쭈꾸미볶음", { bold: true }), c("소액 테스트"), c("370g, 쭈꾸미 75%, 맵기 3단계, 당류 100g당 1g 미만"), c("셰프 양념, 함량, 맵기 조절")],
    [c("새조개 샤브샤브", { bold: true }), c("2027년 1분기 시즌"), c("내년 상반기 재운영"), c("시즌 한정과 셰프")],
    [c("참돔 물회", { bold: true }), c("2027년 2분기 핵심"), c("7월~8월 초 매출 70%"), c("검증된 상품, 셰프 얼굴 소재")],
  ], TB({ x: MX, y: TOP, w: CW, colW: [2.5, 2.45, 4.1, 3.08], rowH: [0.45, 0.75, 0.7, 0.7, 0.6, 0.6] }));
  card(s, MX, 5.75, CW, 0.6);
  T(s, "콰삭오징어는 전복죽 주문에 하나 더 담는 안주 상품으로 활용", { x: MX + 0.3, y: 5.75, w: CW - 0.6, h: 0.6, fontSize: 14, color: C.sub, valign: "middle" });
  s.addNotes("생선구이·쭈꾸미는 판매 비중이 작아 4분기에는 반응만 확인하고, 좋으면 다음 분기에 키웁니다.");
}

// 9. 두 대상
{
  const s = content(SEC.II, "같은 전복죽도 대상에 따라 다르게 말합니다", "Ⅱ 전략 · 두 대상");
  const c = cell(15);
  const lab = (t) => c(t, { bold: true, color: C.sub, fill: { color: HEX.gray } });
  s.addTable([
    head(["구분", "30대 주부", "40~60대"], 16),
    [lab("찾는 이유"), c("아이에게 먹일 음식"), c("부드럽고 담백한 음식, 경제적 여유")],
    [lab("비교 상대"), c("대기업·아이 전용 순살 생선, 마트 파우치 죽"), c("백화점 수산·죽 코너, 선물용 수산 세트")],
    [lab("이길 논리"), c("\"매일은 아니어도, 아이 먹을 것만큼은 셰프 기준으로\"", { bold: true, color: C.accent }), c("\"백화점 팝업에서 선보인 셰프의 한 끼를 집에서\"", { bold: true, color: C.accent })],
    [lab("앞세울 상품"), c("전복죽(아이 아침), 생선구이(국내산 삼치)"), c("전복죽, 생선구이, 쭈꾸미(당류), 선물")],
  ], TB({ x: MX, y: TOP, w: CW, colW: [1.9, 5.1, 5.13], rowH: [0.5, 0.8, 0.8, 0.95, 0.8], margin: [0.06, 0.18, 0.06, 0.18] }));
  T(s, "20대 이하는 구매가 거의 없어 광고 대상에서 제외", { x: MX, y: 5.85, w: CW, h: 0.3, fontSize: 12, color: C.cap });
  s.addNotes("지금은 연령 그룹을 나눠 놓고 같은 소재를 보여주고 있습니다. 나누려면 메시지도 나눠야 합니다.");
}

// =====================================================================
// Ⅲ 운영안
// =====================================================================
pres.addSection({ title: SEC.III });

// 10. 착지와 측정
{
  const s = content(SEC.III, "신규는 네이버에서 스토어로, 다시 부르기는 카카오에서 자사몰로", "Ⅲ 운영안 · 착지와 측정");
  const rows = [
    ["신규", "네이버", "검색광고·GFA", "스마트스토어", "결제 문턱이 가장 낮아 구매 데이터가 빨리 쌓임. 전환 자동 연동, 별점·찜 노출", C.accent],
    ["다시 부르기", "카카오", "디스플레이·비즈보드", "자사몰", "스마트스토어엔 카카오 픽셀 설치 불가. 자사몰에서만 구매 측정", C.slate],
    ["재구매", "카카오톡 채널", "재구매 메시지", "자사몰", "회원 데이터, 수수료 절감, 정기배송", C.dark],
  ];
  rows.forEach(([tag, media, sub, land, why, col], i) => {
    const y = TOP + 0.1 + i * 1.6;
    T(s, tag, { x: MX, y, w: 3.2, h: 0.3, fontSize: 13, bold: true, color: col });
    card(s, MX, y + 0.35, 3.2, 1.05, col);
    T(s, [{ text: media, options: { bold: true, fontSize: 18, breakLine: true } }, { text: sub, options: { fontSize: 13 } }], { x: MX + 0.25, y: y + 0.35, w: 2.8, h: 1.05, color: C.white, valign: "middle" });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 3.95, y: y + 0.7, w: 0.6, h: 0.35, fill: { color: HEX.line }, line: { type: "none" } });
    card(s, 4.7, y + 0.35, 2.5, 1.05, C.gray, { line: { color: col, width: 1.5 } });
    T(s, land, { x: 4.7, y: y + 0.35, w: 2.5, h: 1.05, fontSize: 18, bold: true, align: "center", valign: "middle" });
    T(s, why, { x: 7.5, y: y + 0.35, w: W - MX - 7.5, h: 1.05, fontSize: 14, color: C.sub, valign: "middle" });
  });
  s.addNotes("구매 데이터가 부족한 지금은 전환이 가장 잘 일어나는 곳에서 먼저 쌓고, 그 고객을 자사몰 단골로 옮기는 순서입니다.");
}

// 11. 네이버
{
  const s = content(SEC.III, "네이버는 사려는 사람에게만 돈을 씁니다", "Ⅲ 운영안 · 매체별 캠페인 설계 — 네이버");
  const c = cell(13);
  const b = (t) => c(t, { bold: true });
  s.addTable([
    head(["캠페인", "대상·키워드", "운영 원칙"], 14),
    [b("파워링크 셰프·브랜드"), c("장호준, 장호준 셰프, 장호준 전복죽, 통영숙수"), c("상시 1순위. 확장소재에 셰프 얼굴, \"2팩부터 무료배송\"")],
    [b("파워링크 전복죽"), c("게우 전복죽, 통전복죽, 아이 아침 전복죽, 전복죽 밀키트"), c("입찰 상한, 2주마다 구매 없는 키워드 정리")],
    [b("쇼핑검색"), c("전복죽 중심 + 생선구이 소액. 상품명·셰프명 결합 키워드"), c("'쭈꾸미' 단독 일반어 중단, 콘텐츠 매체 끄기, 제외키워드 등록")],
    [b("GFA ADVoost"), c("25세 이상, 전복죽 비중을 높인 상품 세트"), c("상품·연령으로 쪼개지 않음, 트래픽 목적·스마트채널 없음")],
    [b("GFA 카탈로그 리타겟"), c("상세 조회·장바구니 고객"), c("담고 안 산 고객 회수, 무료배송 문구 함께")],
    [b("GFA 생선구이·쭈꾸미 테스트"), c("① 30대 여성 + 육아·요리 관심  ② 40~60대"), c("소액으로 대상별 반응 확인")],
  ], TB({ x: MX, y: TOP, w: CW, colW: [3.0, 4.6, 4.53], rowH: [0.45, 0.62, 0.62, 0.62, 0.62, 0.55, 0.62] }));
  card(s, MX, 6.0, CW, 0.55, HEX.tint);
  T(s, "트래픽 목적 캠페인은 운영하지 않습니다. 구매가 목적인데 클릭 잘하는 사람에게 노출되도록 학습되기 때문입니다", { x: MX + 0.3, y: 6.0, w: CW - 0.6, h: 0.55, fontSize: 13, valign: "middle" });
  s.addNotes("트래픽 목적 캠페인은 운영하지 않습니다. 구매가 목적인데 클릭 잘하는 사람에게 노출되도록 학습되기 때문입니다.");
}

// 12. 카카오
{
  const s = content(SEC.III, "카카오는 다시 부르는 데 씁니다", "Ⅲ 운영안 · 매체별 캠페인 설계 — 카카오");
  const c = cell(15);
  s.addTable([
    head(["캠페인", "대상", "운영 원칙"], 15),
    [c("카카오 디스플레이 리타겟", { bold: true }), c("자사몰 방문·장바구니 고객, 채널 친구"), c("카카오 픽셀 설치 후 시작")],
    [c("카카오톡 채널", { bold: true }), c("구매 고객"), c("패키지 동봉 QR·가입 쿠폰으로 친구 모집, 전복죽 재구매·시즌 상품 안내")],
    [c("카카오 비즈보드", { bold: true }), c("40~60대"), { text: [{ text: "4분기 초 보류. ", options: { bold: true, color: C.warn } }, { text: "BEP 근접 후 또는 11~12월 브랜디드 콘텐츠 시점에 시작" }], options: { fontSize: 15, color: C.text } }],
  ], TB({ x: MX, y: TOP, w: CW, colW: [3.3, 3.6, 5.23], rowH: [0.5, 0.95, 0.95, 0.95], margin: [0.06, 0.18, 0.06, 0.18] }));
  const steps = [["회수", "디스플레이 리타겟"], ["재구매", "카카오톡 채널"], ["인지", "비즈보드는 효율이 나온 뒤에"]];
  const sw = (CW - 2 * 0.5) / 3;
  steps.forEach(([h, d], i) => {
    const x = MX + i * (sw + 0.5);
    card(s, x, 5.3, sw, 0.95, i < 2 ? HEX.tint : C.gray);
    T(s, [{ text: h + "  ", options: { bold: true, color: i < 2 ? C.accent : C.sub } }, { text: d, options: { color: C.sub } }], { x: x + 0.3, y: 5.3, w: sw - 0.6, h: 0.95, fontSize: 15, valign: "middle" });
    if (i < 2) T(s, "›", { x: x + sw, y: 5.3, w: 0.5, h: 0.95, fontSize: 28, bold: true, color: C.cap, align: "center", valign: "middle" });
  });
  s.addNotes("인지 광고는 효율이 나온 뒤에 붙입니다. 지금은 회수와 재구매가 먼저입니다.");
}

// 13. 메시지와 카피
{
  const s = content(SEC.III, "상품마다, 대상마다 다른 한 줄", "Ⅲ 운영안 · 메시지와 카피");
  const c = cell(13);
  const b = (t) => c(t, { bold: true });
  T(s, "메시지", { x: MX, y: TOP - 0.05, w: 4, h: 0.3, fontSize: 14, bold: true, color: C.accent });
  s.addTable([
    head(["상품", "30대", "40~60대", "선물"], 13),
    [b("전복죽"), c("아이 아침은 전복죽 한 팩으로"), c("내장을 아낌없이 넣은 초록빛 전복죽"), c("부모님 댁에 보내는 셰프의 전복죽")],
    [b("생선구이"), c("아이 생선 반찬, 가시 걱정 없이"), c("48시간 숙성, 속살까지 밴 셰프의 간"), c("셰프 브랜드 생선 선물 구성")],
    [b("쭈꾸미"), c("맵기 3단계, 온 가족이 함께"), c("설탕은 덜고 유자와 매실로 잡은 감칠맛"), c("—", { color: C.cap })],
  ], TB({ x: MX, y: TOP + 0.3, w: CW, colW: [1.5, 3.4, 3.9, 3.33], rowH: 0.42, margin: [0.04, 0.12, 0.04, 0.12] }));
  T(s, "파워링크 카피 (제목 15자 / 설명 45자)", { x: MX, y: 3.95, w: 6, h: 0.3, fontSize: 14, bold: true, color: C.accent });
  s.addTable([
    head(["키워드 묶음", "제목", "설명"], 13),
    [b("셰프·브랜드"), c("장호준 셰프의 통영숙수", { bold: true, color: C.accent }), c("통영에서 나고 자란 장호준 셰프가 직접 개발한 전복죽, 생선구이, 쭈꾸미")],
    [b("전복죽"), c("게우 넣은 통전복죽", { bold: true, color: C.accent }), c("완도산 전복 3마리분에 내장까지 넣은 초록빛 전복죽. 밥솥 취사 한 번이면 완성")],
    [b("생선구이"), c("48시간 숙성 생선구이", { bold: true, color: C.accent }), c("속살까지 간이 밴 고등어·삼치. 통영 출신 장호준 셰프의 비법 생선구이")],
  ], TB({ x: MX, y: 4.3, w: CW, colW: [1.5, 2.9, 7.73], rowH: 0.42, margin: [0.04, 0.12, 0.04, 0.12] }));
  s.addNotes("쭈꾸미의 '가을'은 메뉴 제안으로만 씁니다. 원산지 표현은 쓰지 않습니다.");
}

// 14-1. 신규 소재 기획: 원칙
const TITLE14 = "소재는 새로 짭니다. 셰프 얼굴, 근거 하나, 무료배송";
{
  const s = content(SEC.III, TITLE14, "Ⅲ 운영안 · 신규 소재 기획 ① 공통 원칙");
  const cards = [
    ["셰프 얼굴 필수", "통영 출신 셰프가 첫눈에 보이게"],
    ["근거 배지 하나", "완도산 전복 3마리분 / 48시간 숙성 / 가시 99.9% 제거 / 당류 100g당 1g 미만 중 하나만"],
    ["CTA 통일", "\"2팩부터 무료배송\""],
    ["쓰지 않는 것", "'실속', 할인율 강조, 대상과 맞지 않는 메시지, 제철·국내산(쭈꾸미·고등어), 효능 표현"],
  ];
  const cw = (CW - 0.3) / 2, ch = 1.75;
  cards.forEach(([h, d], i) => {
    const x = MX + (i % 2) * (cw + 0.3), y = TOP + Math.floor(i / 2) * (ch + 0.25);
    const no = i === 3;
    card(s, x, y, cw, ch, no ? HEX.warnTint : C.gray);
    badge(s, x + 0.3, y + 0.3, no ? "×" : i + 1, no ? C.warn : C.accent);
    T(s, h, { x: x + 0.9, y: y + 0.28, w: cw - 1.1, h: 0.46, fontSize: 18, bold: true, valign: "middle" });
    T(s, d, { x: x + 0.9, y: y + 0.85, w: cw - 1.2, h: 0.8, fontSize: 15, color: C.sub });
  });
  card(s, MX, 5.75, CW, 0.75, HEX.tint);
  T(s, [
    { text: "사이즈  ", options: { bold: true, color: C.accent } },
    { text: "GFA 피드 1200×628, 1:1(1200×1200) / 카카오 디스플레이 동일 / 카카오 비즈보드 1029×258 (J만)" },
  ], { x: MX + 0.3, y: 5.75, w: CW - 0.6, h: 0.75, fontSize: 15, valign: "middle" });
  s.addNotes("소재는 전복죽 5종을 중심으로, 생선구이·쭈꾸미는 대상별 2종씩 테스트용으로 만듭니다. 모든 소재에 셰프 얼굴, 근거 배지 하나, \"2팩부터 무료배송\"을 넣습니다.");
}

// 14-2. 신규 소재 기획: 라인업
{
  const s = content(SEC.III, TITLE14, "Ⅲ 운영안 · 신규 소재 기획 ② 소재 라인업");
  const c = cell(11);
  const rows = [
    ["A", "전복죽", "40~60대", "통영 출신 셰프가 끓인 초록빛 전복죽", "완도산 전복 3마리분, 내장까지", "셰프 상반신 + 초록빛 전복죽 클로즈업", "완도산 전복 3마리분", "2주차"],
    ["B", "전복죽", "30대", "아이 아침, 셰프의 전복죽 한 그릇", "밥솥에 넣고 취사 한 번, 1팩 2인분", "아침 식탁, 아이와 나눠 먹는 장면 + 셰프", "취사 13분", "2주차"],
    ["C", "전복죽", "리타겟 (공통)", "2팩이면 4인분, 배송비도 무료", "담아두신 전복죽, 지금 2팩으로", "전복죽 2팩 + 셰프", "5만 원 이상 무료배송", "2주차"],
    ["D", "생선구이", "30대 (테스트)", "아이 생선 반찬, 가시 걱정 없이", "가시 99.9% 제거, 48시간 숙성", "아이 식탁 + 가시 없는 살 클로즈업", "국내산 삼치", "3주차"],
    ["E", "생선구이", "40~60대 (테스트)", "48시간 숙성, 속살까지 밴 셰프의 간", "통영 출신 장호준 셰프의 생선구이", "셰프 + 정갈한 상차림", "48시간 숙성", "3주차"],
    ["F", "쭈꾸미", "30대 (테스트)", "맵기는 우리 집에 맞게, 3단계", "쭈꾸미 75%, 셰프 양념", "가족 저녁 식탁 + 셰프", "맵기 3단계", "3주차"],
    ["G", "쭈꾸미", "40~60대 (테스트)", "설탕은 덜고, 유자와 매실로 잡은 감칠맛", "장호준 셰프의 쭈꾸미볶음", "셰프 + 쭈꾸미볶음", "당류 100g당 1g 미만", "3주차"],
    ["H", "전복죽", "40~60대 (11월)", "찬 바람 불 때, 진한 전복죽 한 그릇", "내장까지 넣어 초록빛이 납니다", "김 오르는 전복죽 + 셰프", "완도산 전복 3마리분", "5~6주차"],
    ["I", "전복죽", "선물 (12월)", "부모님 댁에 보내는 셰프의 전복죽", "백화점 팝업에서 선보인 통영숙수", "선물 포장 + 셰프", "배송 마감일", "7~8주차"],
    ["J", "브랜드", "비즈보드 (12월~)", "통영에서 나고 자란 셰프의 한 끼", "장호준 셰프의 통영숙수", "셰프 + 대표 상품 3종", "—", "7~8주차"],
  ];
  const isMain = (r) => r[1] === "전복죽";
  s.addTable([
    head(["코드", "상품", "대상·용도", "헤드라인", "보조 문구", "비주얼", "배지", "투입"], 11),
    ...rows.map((r) => r.map((t, j) => c(t, {
      bold: j === 0 || j === 3, align: j === 0 ? "center" : "left",
      color: j === 0 ? C.white : j === 3 ? C.accent : C.text,
      fill: j === 0 ? { color: isMain(r) ? HEX.accent : HEX.slate } : undefined,
    }))),
  ], TB({ x: MX, y: TOP, w: CW, colW: [0.5, 0.78, 1.45, 2.75, 2.3, 2.25, 1.4, 0.7], rowH: 0.43, margin: [0.03, 0.07, 0.03, 0.07] }));
  s.addNotes("소재는 전복죽 5종을 중심으로, 생선구이·쭈꾸미는 대상별 2종씩 테스트용으로 만듭니다. 2주 단위로 성과를 보고 하위 소재를 교체합니다.");
}

// 15. 소재 시안 (소재 자리는 비워 둠 — 수동 삽입)
{
  const s = content(SEC.III, "핵심 소재 네 장은 이렇게 구성합니다", "Ⅲ 운영안 · 소재 시안");
  const slots = [
    ["A", "셰프 × 초록빛 전복죽", "40~60대"],
    ["B", "아이 아침", "30대"],
    ["C", "2팩 무료배송", "리타겟"],
    ["D", "아이 생선 반찬", "30대 테스트"],
  ];
  const fw = 3.85, fh = fw * 628 / 1200, colW = (CW - 0.4) / 2;
  slots.forEach(([code, name, who], i) => {
    const x = MX + (i % 2) * (colW + 0.4), y = TOP + 0.05 + Math.floor(i / 2) * (fh + 0.3);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: fw, h: fh, fill: { color: HEX.gray }, line: { color: HEX.cap, width: 1, dashType: "dash" }, objectName: `소재_${code}` });
    const tx = x + fw + 0.25;
    T(s, code, { x: tx, y: y + 0.05, w: 1.5, h: 0.6, fontSize: 32, bold: true, color: C.accent });
    T(s, name, { x: tx, y: y + 0.75, w: colW - fw - 0.25, h: 0.7, fontSize: 16, bold: true });
    T(s, who, { x: tx, y: y + 1.45, w: colW - fw - 0.25, h: 0.3, fontSize: 13, color: C.sub });
  });
  T(s, "실제 촬영·제작은 운영 확정 후 진행. 셰프 IP(얼굴·이름·영상) 사용 가능 확인 완료", { x: MX, y: 6.35, w: CW, h: 0.3, fontSize: 12, color: C.cap });
  s.addNotes(
    "지난여름 반응이 가장 좋았던 물회 광고는 셰프 얼굴, 구체적인 근거, 여름 한정이라는 이유를 한 장에 담았습니다. 신규 소재는 이 틀을 그대로 가져갑니다.\n\n" +
    "[구성 참고]\n" +
    "A 셰프 × 초록빛 전복죽(40~60대): 왼쪽 절반 셰프 상반신 / 오른쪽 전복죽 클로즈업 / 상단 헤드라인 / 우하단 \"완도산 전복 3마리분\" 원형 배지 / 하단 띠 \"2팩부터 무료배송\"\n" +
    "B 아이 아침(30대): 배경 아침 식탁 / 중앙 전복죽 / 좌하단 작은 셰프 얼굴 원형 / 상단 헤드라인 / \"취사 13분\" 배지 / 하단 띠 CTA\n" +
    "C 2팩 무료배송(리타겟): 전복죽 2팩 나란히 / 큰 문구 \"2팩이면 4인분, 배송비도 무료\" / 셰프 얼굴 작게 / 하단 띠 CTA\n" +
    "D 아이 생선 반찬(30대 테스트): 아이 식탁 + 생선 살 클로즈업 / 헤드라인 / \"국내산 삼치\" 배지 / 하단 띠 CTA"
  );
}

// 16. 객단가와 상품명
{
  const s = content(SEC.III, "상품명이 곧 광고 문구, 무료배송이 곧 객단가", "Ⅲ 운영안 · 객단가와 상품명");
  const lw = 5.5;
  card(s, MX, TOP, lw, 4.15);
  T(s, "2팩부터 무료배송", { x: MX + 0.35, y: TOP + 0.25, w: lw - 0.7, h: 0.4, fontSize: 18, bold: true, color: C.accent });
  T(s, [
    { text: "무료배송 5만 원 = 전복죽 2팩(51,800원)", options: { bold: true, breakLine: true } },
    { text: "1팩이 2인분이라 \"2팩 = 4인분\"", options: { color: C.sub } },
  ], { x: MX + 0.35, y: TOP + 0.75, w: lw - 0.7, h: 0.7, fontSize: 15 });
  T(s, "쓰는 곳", { x: MX + 0.35, y: TOP + 1.6, w: 2, h: 0.3, fontSize: 13, bold: true, color: C.sub });
  T(s, bullets([
    "파워링크 확장소재 \"전복죽 2팩부터 무료배송\"",
    "배너 \"2팩이면 4인분, 배송비도 무료\"",
    "리타겟 \"담아두신 전복죽, 2팩이면 무료배송\"",
    "묶음 \"전복죽 1팩 + 고등어구이 세트(스토어 기준 50,400원)\"",
  ], { paraSpaceAfter: 5 }), { x: MX + 0.35, y: TOP + 1.95, w: lw - 0.65, h: 2.1, fontSize: 14 });
  const rx = MX + lw + 0.35, rw = W - MX - rx;
  T(s, "상품명 정리", { x: rx, y: TOP, w: 4, h: 0.35, fontSize: 18, bold: true, color: C.accent });
  const c = cell(13);
  s.addTable([
    head(["지금", "제안"], 13),
    [c("[BEST] 통영숙수 실속 2종 패키지 (고등어+삼치)", { color: C.sub }), c("장호준의 48시간 숙성 생선구이 2종 (고등어·삼치)", { bold: true })],
    [c("통영숙수 비법소스 고등어구이 세트", { color: C.sub }), c("장호준의 48시간 숙성 고등어구이", { bold: true })],
    [c("장호준의 名品 게우 통전복죽", { color: C.sub }), c("장호준의 게우 통전복죽 (완도산 전복 3마리분)", { bold: true })],
    [c("장호준의 인생쭈꾸미", { color: C.sub }), c("장호준의 쭈꾸미볶음 370g (맵기 3단계)", { bold: true })],
  ], TB({ x: rx, y: TOP + 0.5, w: rw, colW: [rw / 2, rw / 2], rowH: [0.4, 0.82, 0.62, 0.82, 0.62] }));
  card(s, MX, 6.0, CW, 0.6, HEX.tint);
  T(s, "상시 할인 대신 11월, 연말, 설처럼 기간을 정한 행사로 운영하면 프리미엄 이미지를 지킬 수 있습니다", { x: MX + 0.3, y: 6.0, w: CW - 0.6, h: 0.6, fontSize: 14, valign: "middle" });
  s.addNotes("ADVoost와 쇼핑검색은 상품명을 그대로 광고 문구로 씁니다. '인생쭈꾸미'는 다른 업체도 쓰는 이름이라 검색에서 혼동이 생깁니다.");
}

// =====================================================================
// Ⅳ 예산·일정
// =====================================================================
pres.addSection({ title: SEC.IV });

// 17. 목표까지의 경로 (계단형)
{
  const s = content(SEC.IV, "82%에서 450%까지, 네 단계로 올립니다", "Ⅳ 예산·일정 · 목표까지의 경로");
  const base = 4.15, scale = 2.05 / 450;
  const steps = [
    { name: "현재", when: "10월 1~14일", lever: "", v: 82, label: "82%", col: HEX.cap },
    { name: "1. 구조 정리", when: "10월 말~11월", lever: "저효율 노출 제거, 쭈꾸미 일반어 중단, 셰프 파워링크, 카탈로그. 스토어 전환율 9월 수준(4.9%) 회복", v: 200, label: "200% 이상", col: "8CC8EE" },
    { name: "2. 객단가·리타겟", when: "11~12월", lever: "2팩부터 무료배송, 리타겟 확대, 브랜디드 콘텐츠 연계", v: 288, label: "12월 BEP 근접", col: "4E9BD1" },
    { name: "3. 시즌·재구매", when: "2027년 1분기", lever: "새조개 샤브샤브, 설 선물(진행 시), 카카오톡 채널 재구매", v: 330, label: "BEP 300% 이상", col: HEX.accent },
    { name: "4. 확장", when: "2027년 2분기", lever: "지난여름 매출 70% 물회 재출시, 쌓인 구매 데이터", v: 450, label: "450%", col: "0E4A77" },
  ];
  const widths = [1.55, 2.55, 2.55, 2.55, 2.55], gap = (CW - widths.reduce((a, b) => a + b, 0)) / 4;
  let x = MX;
  steps.forEach((st, i) => {
    const w = widths[i], h = st.v * scale;
    s.addShape(pres.shapes.RECTANGLE, { x, y: base - h, w, h, fill: { color: st.col }, line: { type: "none" } });
    if (i === 0) T(s, st.label, { x, y: base - h - 0.4, w, h: 0.35, fontSize: 14, bold: true, align: "center", valign: "bottom", color: C.sub });
    else T(s, st.label, { x, y: base - 0.45, w, h: 0.38, fontSize: i === 4 ? 18 : 15, bold: true, align: "center", valign: "middle", color: i === 1 ? C.text : C.white });
    T(s, st.name, { x, y: base + 0.12, w, h: 0.3, fontSize: 14, bold: true });
    T(s, st.when, { x, y: base + 0.42, w, h: 0.28, fontSize: 12, bold: true, color: C.accent });
    if (st.lever) T(s, st.lever, { x, y: base + 0.72, w, h: 1.2, fontSize: 12, color: C.sub });
    x += w + gap;
  });
  const by = base - 300 * scale;
  s.addShape(pres.shapes.LINE, { x: MX, y: by, w: CW, h: 0, line: { color: HEX.warn, width: 1.75, dashType: "dash" } });
  T(s, "BEP 300%", { x: MX, y: by - 0.3, w: 1.5, h: 0.26, fontSize: 12, bold: true, color: C.warn });
  card(s, MX, 6.15, CW, 0.5, C.accent);
  T(s, [{ text: "증액 원칙  ", options: { bold: true } }, { text: "BEP 300% 이상을 2주 연속 유지하면 증액" }], { x: MX + 0.3, y: 6.15, w: CW - 0.6, h: 0.5, fontSize: 15, color: C.white, valign: "middle" });
  s.addNotes("바로 450%로 갈 수는 없습니다. 효율을 깎는 요인을 하나씩 걷어내며 올라가는 경로를 약속드리는 겁니다.\n(막대 높이는 단계별 목표를 나타낸 개념도입니다. 목표·BEP는 광고주 제공)");
}

// 18. 4분기 예산안
{
  const s = content(SEC.IV, "월 300만 원으로 시작해, BEP에 가까워지면 400만 원", "Ⅳ 예산·일정 · 4분기 예산안");
  const c = cell(12);
  const r = (m, n, a, b2, roas, o = {}) => [c(m, { color: C.sub }), c(n, { bold: true }), c(a, { align: "right", ...o }), c(b2, { align: "right" }), c(roas, { align: "right" })];
  const rows = [
    ["네이버", "파워링크 셰프·브랜드", "24만 (8%)", "24만 (6%)", "300~500%"],
    ["네이버", "파워링크 전복죽", "21만 (7%)", "24만 (6%)", "150~220%"],
    ["네이버", "쇼핑검색", "60만 (20%)", "72만 (18%)", "200~250%"],
    ["네이버", "GFA ADVoost", "105만 (35%)", "132만 (33%)", "200~260%"],
    ["네이버", "GFA 카탈로그 리타겟", "45만 (15%)", "60만 (15%)", "250~350%"],
    ["네이버", "생선구이·쭈꾸미 테스트", "21만 (7%)", "24만 (6%)", "80~150%"],
    ["카카오", "디스플레이 리타겟", "15만 (5%)", "20만 (5%)", "150~250%"],
    ["카카오", "카카오톡 채널 메시지", "9만 (3%)", "12만 (3%)", "200~300%"],
    ["카카오", "비즈보드", "보류", "32만 (8%)", "50~100%"],
  ];
  const tot = { fill: { color: HEX.tint }, bold: true };
  s.addTable([
    head(["매체", "캠페인", "300만 원\n(10월 말~11월)", "400만 원\n(12월)", "예상 ROAS\n(추정)"], 12),
    ...rows.map((x) => r(...x, x[2] === "보류" ? { color: C.warn, bold: true } : {})),
    [c("", tot), c("합계", tot), c("300만 원", { align: "right", ...tot }), c("400만 원", { align: "right", ...tot }), c("201~281% /\n189~265%", { align: "right", ...tot, color: C.accent })],
  ], TB({ x: MX, y: TOP, w: 7.9, colW: [0.8, 2.5, 1.6, 1.45, 1.55], rowH: [0.5, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33, 0.33, 0.5], margin: [0.03, 0.1, 0.03, 0.1] }));
  const bars = [["GFA ADVoost", 105], ["쇼핑검색", 60], ["GFA 카탈로그 리타겟", 45], ["파워링크 셰프·브랜드", 24], ["파워링크 전복죽", 21], ["생선구이·쭈꾸미 테스트", 21], ["카카오 디스플레이 리타겟", 15], ["카카오톡 채널 메시지", 9]];
  s.addChart(pres.charts.BAR, [{ name: "300만 원 배분", labels: bars.map((b) => b[0]), values: bars.map((b) => b[1]) }], {
    x: 8.65, y: TOP - 0.05, w: 4.2, h: 4.05, barDir: "bar", barGapWidthPct: 40, chartColors: [HEX.accent], catAxisOrientation: "maxMin",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.text, dataLabelFormatCode: '0"만"',
    catAxisLabelColor: HEX.sub, catAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false,
    valAxisHidden: true, valAxisMaxVal: 130, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "300만 원 배분 (10월 말~11월)", titleFontSize: 12, titleFontFace: "+mn-lt", titleColor: HEX.text,
  });
  card(s, MX, 5.8, CW, 0.62, HEX.warnTint);
  T(s, "4분기 예상치는 BEP보다 낮습니다. 4분기는 구조를 바로잡고 데이터를 쌓는 기간이며, 12월 BEP 근접을 목표로 운영합니다. 스토어 전체 결제 건수·전환율도 함께 보고합니다.", { x: MX + 0.3, y: 5.8, w: CW - 0.6, h: 0.62, fontSize: 13, valign: "middle" });
  T(s, "예상 ROAS는 3분기·10월 실적에 구조 개선 효과를 반영한 추정치(광고 매체 집계, 구매완료 기준). 예산 범위는 광고주 제공", { x: MX, y: 6.5, w: CW - 1.5, h: 0.25, fontSize: 10, color: C.cap });
  s.addNotes(
    "4분기 예상치는 BEP보다 낮습니다. 4분기는 구조를 바로잡고 데이터를 쌓는 기간이며, 12월 BEP 근접을 목표로 운영합니다.\n\n" +
    "매체별 근거:\n" +
    "- 셰프 파워링크: 3분기 셰프명 검색은 비용 거의 없이 구매. 검색량이 적어 예산 소진에 한계\n" +
    "- 쇼핑검색: 10월 전복죽 193%, 메인 키워드 228%\n" +
    "- ADVoost: 10월 ADVoost 전복죽 187%\n" +
    "- 카탈로그: 10월 GFA 장바구니 매출이 구매 매출의 약 3배\n" +
    "- 카카오: 이력이 없어 보수적으로 추정"
  );
}

// 19. 착수 후 8주
{
  const s = content(SEC.IV, "첫 주부터 새는 곳을 막습니다", "Ⅳ 예산·일정 · 착수 후 8주");
  const c = cell(13);
  const wk = (t, col) => ({ text: t, options: { bold: true, fontSize: 14, color: C.white, fill: { color: col }, align: "center" } });
  s.addTable([
    head(["주차", "할 일", "확인할 것"], 14),
    [wk("1주", HEX.accent), c("계정 정리(그룹·이름, 트래픽 재개 안 함, 스마트채널 제외), 쭈꾸미 일반어 중단, 제외키워드, 콘텐츠 매체 끄기, 카카오 픽셀 설치"), c("전환 집계, 스토어 전환율", { color: C.sub })],
    [wk("2주", HEX.accent), c("셰프 파워링크, ADVoost 통합(전복죽 비중 확대), 카탈로그 리타겟, \"2팩부터 무료배송\""), c("셰프명 검색 구매, 객단가", { color: C.sub })],
    [wk("3주", HEX.accent), c("셰프 얼굴 전복죽 신규 소재, 생선구이·쭈꾸미 테스트, 카카오 디스플레이 리타겟"), c("대상별 클릭률·장바구니", { color: C.sub })],
    [wk("4주", HEX.slate), c("1차 결과 보고, 하위 키워드·소재 교체"), c("ROAS 200% 이상, 전환율 회복", { color: C.sub })],
    [wk("5~6주", HEX.slate), c("이긴 메시지 확산, 브랜디드 콘텐츠 일정 맞춰 준비, 채널 친구 모집"), c("리타겟 매출 비중, 채널 친구 수", { color: C.sub })],
    [wk("7~8주", HEX.slateDark), c("BEP 근접 시 400만 원 증액, 비즈보드 시작, 연말·설 준비"), c("12월 BEP 근접 여부", { color: C.sub })],
  ], TB({ x: MX, y: TOP, w: CW, colW: [1.2, 7.4, 3.53], rowH: [0.45, 0.8, 0.62, 0.62, 0.55, 0.62, 0.62] }));
  card(s, MX, 6.0, CW, 0.55, HEX.tint);
  T(s, "1주차 작업만으로도 10월에 클릭만 만들던 노출이 대부분 정리됩니다", { x: MX + 0.3, y: 6.0, w: CW - 0.6, h: 0.55, fontSize: 14, valign: "middle" });
  s.addNotes("1주차 작업만으로도 10월에 클릭만 만들던 노출이 대부분 정리됩니다.");
}

// 20. 로드맵과 캘린더
{
  const s = content(SEC.IV, "4분기에 쌓고, 1분기에 지키고, 2분기에 키웁니다", "Ⅳ 예산·일정 · 로드맵과 캘린더");
  const qs = [
    ["2026년 4분기", "12월 BEP 근접", "전복죽", "구조 정리, 셰프 파워링크, 카탈로그, 2팩 무료배송, 브랜디드 콘텐츠", "8CC8EE"],
    ["2027년 1분기", "BEP 300% 이상", "새조개 샤브샤브, 설 선물(진행 시), 전복죽", "시즌 상품 2주 전 소재 준비, 채널 재구매, 테스트 상품 확대", HEX.accent],
    ["2027년 2분기", "450%", "참돔 물회 재출시", "물회 소재 틀로 사전 예고, BEP 이상 시 단계 증액", "0E4A77"],
  ];
  const gap = 0.35, cw = (CW - 2 * gap) / 3, ch = 2.15;
  qs.forEach(([q, goal, prod, task, col], i) => {
    const x = MX + i * (cw + gap);
    card(s, x, TOP, cw, ch);
    T(s, q, { x: x + 0.25, y: TOP + 0.18, w: cw - 0.5, h: 0.35, fontSize: 16, bold: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: TOP + 0.6, w: 1.75, h: 0.34, rectRadius: 0.08, fill: { color: col }, line: { type: "none" } });
    T(s, goal, { x: x + 0.25, y: TOP + 0.6, w: 1.75, h: 0.34, fontSize: 12, bold: true, color: i === 0 ? C.text : C.white, align: "center", valign: "middle" });
    T(s, [{ text: "앞세울 상품  ", options: { bold: true, color: C.sub } }, { text: prod }], { x: x + 0.25, y: TOP + 1.05, w: cw - 0.5, h: 0.45, fontSize: 12 });
    T(s, task, { x: x + 0.25, y: TOP + 1.5, w: cw - 0.5, h: 0.6, fontSize: 12, color: C.sub });
    if (i < 2) T(s, "›", { x: x + cw, y: TOP + 0.7, w: gap, h: 0.6, fontSize: 26, bold: true, color: C.cap, align: "center", valign: "middle" });
  });
  const months = [
    ["10월", "전복죽 (구조 정리)", "내장을 아낌없이 넣은 초록빛 전복죽"],
    ["11월", "전복죽", "전복죽 2팩부터 무료배송"],
    ["12월", "전복죽, 연말", "부모님 댁에 보내는 셰프의 전복죽"],
    ["1월", "새조개 예고, 설 선물(진행 시)", "올해 설 선물은 통영 출신 셰프의 한 끼로"],
    ["2월", "새조개 샤브샤브", "지금만 만나는 셰프의 새조개 샤브샤브"],
    ["3월", "새조개, 전복죽", "시즌 마감 안내"],
    ["4월", "전복죽, 테스트 상품", "테스트 결과에 따라 결정"],
    ["5월", "물회 사전 예고", "올여름도 셰프의 물회가 돌아옵니다"],
    ["6월", "참돔 물회 재출시", "통영산 참돔 물회, 여름 한정"],
  ];
  const c = cell(10);
  const lab = (t) => c(t, { bold: true, color: C.sub, fill: { color: HEX.gray } });
  const mw = (CW - 1.0) / 9;
  s.addTable([
    [{ text: "월", options: { bold: true, color: C.white, fill: { color: HEX.slateDark }, fontSize: 11 } }, ...months.map((m, i) => ({ text: m[0], options: { bold: true, color: C.white, fill: { color: i < 3 ? HEX.accent : HEX.slateDark }, fontSize: 11, align: "center" } }))],
    [lab("앞세울 상품"), ...months.map((m) => c(m[1], { bold: true }))],
    [lab("메시지 예시"), ...months.map((m) => c(m[2], { color: C.sub }))],
  ], TB({ x: MX, y: 4.05, w: CW, colW: [1.0, ...Array(9).fill(mw)], rowH: [0.32, 0.62, 0.95], margin: [0.04, 0.06, 0.04, 0.06], valign: "top" }));
  T(s, "새조개 샤브샤브·물회 출시 일정과 설 선물 세트 진행 여부는 확정되는 대로 반영", { x: MX, y: 6.15, w: CW, h: 0.25, fontSize: 11, color: C.cap });
  s.addNotes("상품이 바뀌어도 캠페인은 유지하고 소재만 바꿉니다. 시즌 상품은 최소 4주 운영합니다.");
}

// 21. 성과 관리와 다음 단계 (어두운 배경)
{
  const s = pres.addSlide({ masterName: "NC_DARK", sectionTitle: SEC.IV });
  T(s, "4분기를 '테스트하고 끄는 분기'가 아니라 '쌓는 분기'로", { x: 0.9, y: 0.6, w: 11.5, h: 0.7, fontSize: 30, bold: true, color: C.white, valign: "middle" });
  const glass = (x, y, w, h) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.white, transparency: 85 }, line: { color: C.white, width: 0.75, transparency: 55 } });
  const lw = 5.2;
  glass(0.9, 1.65, lw, 3.85);
  T(s, "리포트", { x: 1.2, y: 1.85, w: 3, h: 0.4, fontSize: 18, bold: true, color: C.white });
  T(s, [
    { text: "주간", options: { bold: true, breakLine: true } },
    { text: "매체별 광고비·구매·ROAS, 장바구니 회수, 키워드·소재 교체", options: { breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "월간", options: { bold: true, breakLine: true } },
    { text: "상품별 성과, 대상별 테스트 결과, 셰프·브랜드명 검색량, 채널 친구 수" },
  ], { x: 1.2, y: 2.4, w: lw - 0.6, h: 2.9, fontSize: 14, color: C.white });
  const rx = 0.9 + lw + 0.35, rw = 12.43 - rx;
  glass(rx, 1.65, rw, 3.85);
  T(s, "요청드릴 것", { x: rx + 0.3, y: 1.85, w: 3, h: 0.4, fontSize: 18, bold: true, color: C.white });
  const asks = ["11~12월 브랜디드 콘텐츠 일정과 적용 프로모션", "새조개 샤브샤브 출시 일정, 설 선물 세트 확정 시점", "스마트스토어 신규·재구매 고객 비중", "상품명 정리와 상시 할인 운영에 대한 의견"];
  asks.forEach((t, i) => {
    const y = 2.45 + i * 0.72;
    badge(s, rx + 0.3, y, i + 1, C.accent, 0.4, 13);
    T(s, t, { x: rx + 0.85, y, w: rw - 1.1, h: 0.4, fontSize: 14, color: C.white, valign: "middle" });
  });
  T(s, "모든 전환은 총전환이 아닌 구매완료 기준 · 광고 성과는 통영숙수 캠페인만 집계 · 스마트스토어 통계·판매 비중·목표는 광고주 제공 · NICECORE 퍼포먼스마케팅팀", { x: 0.9, y: 5.75, w: 9.8, h: 0.5, fontSize: 11, color: C.white });
  s.addNotes("통영숙수는 아직 아는 사람이 적은 만큼, 지금 쌓는 구매 데이터와 셰프 인지도가 그대로 설과 내년 여름의 출발점이 됩니다.");
}

// 감사합니다
{
  const s = pres.addSlide({ masterName: "NC_DARK", sectionTitle: SEC.IV });
  T(s, [{ text: "감사합니다", options: { bold: true, fontSize: 44, breakLine: true } }, { text: "THANK YOU", options: { fontSize: 28 } }], { x: 0.95, y: 2.0, w: 7.4, h: 1.6, color: C.white });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 4.01, w: 0.68, h: 0, line: { color: C.white, width: 1.5 } });
  T(s, "언제나 최상의 퍼포먼스로\n클라이언트의 성장이라는 목표를 함께 하겠습니다.", { x: 0.95, y: 4.3, w: 6, h: 0.75, fontSize: 16, color: C.white });
  s.addNotes("감사합니다. 질문 받겠습니다.");
}

(async () => {
  const out = path.join(__dirname, OUT);
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("written", out);
})();
