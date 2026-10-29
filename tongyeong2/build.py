"""통영숙수 2차 과제 회신 PPT 생성 스크립트

실행: python3 build.py  →  output/통영숙수_2차과제_NICECORE.pptx
- 문구·숫자: 통영숙수_2차과제_슬라이드원고.md (그대로)
- 규칙: 통영숙수_2차과제_작업지시서.md
- 디자인: 사용자가 준 레퍼런스 PPT(251112)의 구성 방식을 적용
  · 상단 붉은 가는 선 + 섹션 탭(현재 섹션만 붉은 알약) + 우상단 로고
  · 붉은 라벨 칩 → 결론형 제목 → 본문 → 하단 결론 띠
  · 섹션 성격의 장(표지·마무리)은 남색 패널
  · 글꼴: 레퍼런스와 같은 페이퍼로지(Paperlogy)
  레퍼런스의 문구·로고·이미지는 쓰지 않음
"""
import re
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).parent
OUT = ROOT / "output" / "통영숙수_2차과제_NICECORE.pptx"
LOGO = ROOT / "assets" / "logo_nicecore.png"
LOGO_W = ROOT / "assets" / "logo_nicecore_white.png"

# ---------- 색 (작업지시서 6장 + 레퍼런스 메인 빨강) ----------
RED = "ED1438"        # 레퍼런스 메인 빨강 → 강조색
NAVY = "394B63"       # 표 머리글
NAVY_D = "1E3A50"     # 남색 패널 (레퍼런스 섹션 장)
CARD = "F2F4F6"
CARD_HL = "FFEDEA"
WARN = "F08A00"
COMP = "C4CED8"
GBOX, GLINE = "E5E8EB", "8B95A1"
TEXT, SUB, CAP, LINE, WHITE = "191F28", "4E5968", "8B95A1", "D1D6DB", "FFFFFF"

F_L, F_R, F_SB, F_B, F_XB = "페이퍼로지 3 Light", "페이퍼로지 4 Regular", "페이퍼로지 6 SemiBold", "페이퍼로지 7 Bold", "페이퍼로지 8 ExtraBold"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
MX = 0.5
W = 13.333 - 2 * MX

TABS = ["과제 요약", "① 소재", "② 판정 기준", "③ 비용 누수", "④ 상시 점검", "운영·일정"]


def rgb(h):
    return RGBColor.from_string(h)


# ---------------------------------------------------------------- 텍스트
def set_font(run, size, color=TEXT, font=F_R):
    f = run.font
    f.size = Pt(size)
    f.bold = False
    f.color.rgb = rgb(color)
    f.name = font
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    ea = etree.SubElement(rPr, qn("a:ea"))
    latin.addnext(ea)
    ea.set("typeface", font)


def fill(tf, paras, size=12, color=TEXT, font=F_R, align=PP_ALIGN.LEFT, spacing=1.2, after=0, bullet=False):
    """paras: str | [para]; para: str | [(text, {size,color,font})]"""
    if isinstance(paras, str):
        paras = [paras]
    tf.clear()
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        if after:
            p.space_after = Pt(after)
        pPr = p._p.get_or_add_pPr()
        pPr.set("latinLnBrk", "0")   # 영문 단어 중간에서 끊지 않음
        pPr.set("eaLnBrk", "1")
        runs = [(para, {})] if isinstance(para, str) else para
        for text, o in runs:
            r = p.add_run()
            r.text = text
            set_font(r, o.get("size", size), o.get("color", color), o.get("font", font))
        if bullet:
            sz = runs[0][1].get("size", size)
            pPr.set("marL", str(int(Pt(sz * 0.95))))
            pPr.set("indent", str(-int(Pt(sz * 0.95))))
            etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def tb(s, x, y, w, h, paras, size=12, anchor=MSO_ANCHOR.TOP, name=None, **kw):
    b = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = b.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    fill(tf, paras, size=size, **kw)
    if name:
        b.name = name
    return b


def _nostyle(sh):
    st = sh._element.find(qn("p:style"))
    if st is not None:
        sh._element.remove(st)


def box(s, x, y, w, h, fill_=CARD, line=None, lw=0.75, radius=0.08, shape=MSO_SHAPE.ROUNDED_RECTANGLE, dash=None, name=None):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    _nostyle(sh)
    if fill_ is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill_)
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
        if dash:
            sh.line.dash_style = dash
    else:
        sh.line.fill.background()
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = min(radius / min(w, h), 0.5)
    if name:
        sh.name = name
    return sh


def stext(sh, paras, size=12, anchor=MSO_ANCHOR.MIDDLE, ml=0.15, mt=0.05, **kw):
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Inches(ml)
    tf.margin_top = tf.margin_bottom = Inches(mt)
    tf.vertical_anchor = anchor
    fill(tf, paras, size=size, **kw)


def pill(s, x, y, w, h, text, fill_=RED, color=WHITE, size=11, font=F_B):
    sh = box(s, x, y, w, h, fill_=fill_, radius=h / 2)
    stext(sh, [[(text, {})]], size=size, color=color, font=font, align=PP_ALIGN.CENTER, ml=0.04, mt=0)
    return sh


def badge(s, x, y, n, d=0.4, fill_=RED, size=13):
    sh = box(s, x, y, d, d, fill_=fill_, shape=MSO_SHAPE.OVAL)
    stext(sh, [[(str(n), {})]], size=size, color=WHITE, font=F_B, align=PP_ALIGN.CENTER, ml=0, mt=0)
    return sh


def hline(s, x, y, w, color=RED, width=1.0):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    _nostyle(c)
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(width)
    return c


def gray_box(s, x, y, w, h, ratio, name):
    """소재 자리: 회색 박스 + 점선 + 안내 문구 + 비율 (그룹으로 묶지 않음)"""
    b = box(s, x, y, w, h, fill_=GBOX, line=GLINE, lw=1, radius=0.04, dash=MSO_LINE_DASH_STYLE.DASH, name=name)
    stext(b, [[("소재 이미지", {"font": F_SB})], [("(디자인팀 제작)", {})], [(ratio, {"size": 10, "color": CAP})]],
          size=11, color=SUB, align=PP_ALIGN.CENTER, ml=0.05)
    return b


# ---------------------------------------------------------------- 표
NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def _borders(cell, color=LINE):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, qn(tag), w=str(int(Pt(0.75))))
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    f = tcPr.find(qn("a:solidFill"))
    if f is not None:
        tcPr.remove(f)
        tcPr.append(f)


def table(s, x, y, col_w, rows, row_h=0.4, size=11, head_size=11, first_col=True, aligns=None, fills=None, colors=None):
    nr, nc = len(rows), len(rows[0])
    gf = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(sum(col_w)), Inches(0.3 * nr))
    t = gf.table
    tblPr = t._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = NO_STYLE
    for i, w in enumerate(col_w):
        t.columns[i].width = Inches(w)
    for r in range(nr):
        t.rows[r].height = Inches(row_h[r] if isinstance(row_h, (list, tuple)) else row_h)
        for c in range(nc):
            cell = t.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            head = r == 0
            fc = NAVY if head else (CARD if (first_col and c == 0) else WHITE)
            if fills:
                fc = fills.get((r, c), fills.get(r, fc))
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(fc)
            v = rows[r][c]
            runs = v if isinstance(v, list) else [(str(v), {})]
            col = WHITE if head else (colors or {}).get((r, c), TEXT)
            fill(cell.text_frame, [runs], size=head_size if head else size, color=col,
                 font=F_SB if head or (first_col and c == 0) else F_R,
                 align=(aligns or {}).get(c, PP_ALIGN.CENTER if head else PP_ALIGN.LEFT), spacing=1.15)
            _borders(cell)
    return gf


# ---------------------------------------------------------------- 장 틀
def frame(tab, label, title):
    s = prs.slides.add_slide(BLANK)
    hline(s, MX, 0.15, W, RED, 1.25)
    # 섹션 탭: 현재 섹션만 붉은 알약
    x = MX
    for t in TABS:
        w = 0.32 + len(t) * 0.13
        if t == tab:
            pill(s, x, 0.27, w, 0.28, t, size=10)
        else:
            tb(s, x, 0.27, w, 0.28, [[(t, {})]], size=10, color=CAP, font=F_SB, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += w + 0.08
    s.shapes.add_picture(str(LOGO), Inches(13.333 - MX - 1.5), Inches(0.28), Inches(1.5), Inches(0.22))
    # 라벨 칩 + 결론형 제목 (모든 장 같은 위치·크기)
    pill(s, MX, 0.78, 0.3 + len(label) * 0.14, 0.3, label, size=11)
    tb(s, MX, 1.15, W, 0.55, [[(title, {})]], size=24, font=F_SB, anchor=MSO_ANCHOR.MIDDLE)
    # 하단: 가운데 로고 없이 쪽 번호만 (레퍼런스의 회사 로고 자리는 비워 둠)
    tb(s, 13.333 - MX - 0.6, 7.08, 0.6, 0.25, [[(str(len(prs.slides._sldIdLst)), {})]], size=9, color=CAP, align=PP_ALIGN.RIGHT)
    return s


def band(s, text, y=6.45):
    b = box(s, MX, y, W, 0.5, fill_=RED, radius=0.06)
    stext(b, [[(text, {})]], size=13, color=WHITE, font=F_SB, ml=0.25)
    return b


def navy_slide():
    s = prs.slides.add_slide(BLANK)
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = rgb(NAVY_D)
    hline(s, MX, 0.15, W, RED, 1.25)
    s.shapes.add_picture(str(LOGO_W), Inches(13.333 - MX - 1.5), Inches(0.28), Inches(1.5), Inches(0.22))
    return s


# =====================================================================
# 1. 표지 (사용자가 새로 만들 예정 — 남색 패널로 문구만 배치)
s = navy_slide()
pill(s, 1.1, 2.0, 2.6, 0.36, "NICECORE 2차 과제 회신", size=12)
tb(s, 1.1, 2.55, 11, 0.9, [[("통영숙수 네이버 광고 운영 과제", {})]], size=40, font=F_B, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
tb(s, 1.1, 3.5, 11, 0.45, [[("소재와 판정 기준, 비용 누수 진단, 상시 점검 체계", {})]], size=18, color=COMP, anchor=MSO_ANCHOR.MIDDLE)
hline(s, 1.1, 4.25, 4.2, RED, 1.5)
tb(s, 1.1, 4.45, 11, 0.4, [[("2026년 11월 3일 · NICECORE 퍼포먼스마케팅팀", {})]], size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

# 2. 과제 요약
s = frame("과제 요약", "과제 요약", "네 가지 과제에 대한 답을 먼저 정리했습니다")
cards = [("소재 6종과 가설", "상품별로 3종씩 준비했습니다. 소재마다 바꾸는 요소를 하나로 정해, 무엇이 성과를 냈는지 분명히 가려냅니다."),
         ("유지·교체 기준", "집행 14일 또는 클릭 300회가 지나면 판정합니다. BEP 300%를 기준으로 확대, 유지, 개선, 교체 중 하나를 고릅니다."),
         ("비용 누수 Top 5", "계정 전체와 통영숙수를 함께 살폈습니다. 월 약 516만 원을 줄일 수 있고, 이익 기준으로 보면 약 390만 원입니다."),
         ("상시 점검 체계", "매일, 매주, 매월 보는 항목과 8가지 이상 신호를 정했습니다. 꼬기다 브랜드검색 종료도 3일 치 데이터로 확인했습니다.")]
cw = (W - 3 * 0.25) / 4
for i, (t, d) in enumerate(cards):
    x = MX + i * (cw + 0.25)
    box(s, x, 1.95, cw, 3.05, fill_=CARD)
    badge(s, x + 0.25, 2.15, i + 1)
    tb(s, x + 0.25, 2.7, cw - 0.5, 0.4, [[(t, {})]], size=16, font=F_B)
    tb(s, x + 0.25, 3.2, cw - 0.5, 1.7, [[(d, {})]], size=12, color=SUB, spacing=1.35)
b = box(s, MX, 5.2, W, 0.95, fill_=CARD_HL)
stext(b, [[("김현수 팀장(14년차)과 김범준 매니저(1년차)가 함께 맡고, 사내 디자인팀이 월 40개 이상 소재를 만듭니다.", {"font": F_SB})]], size=14, ml=0.3)
band(s, "모든 성과는 구매완료 기준이며, 2026년 10월 1일~28일 계정 데이터를 썼습니다")


# 3·4. 소재 장 (공통 배치)
jeonbok = [
    {"code": "A", "who": "40~60대 · 셰프형", "head": "통영 사람이 끓인 초록빛 전복죽", "sub": "완도산 통전복 한 마리, 내장까지 그대로",
     "badge": "통전복 한 마리", "cta": "2팩부터 무료배송",
     "hyp": "40~60대는 셰프를 믿을 수 있을 때 구매합니다. 상황형(L)보다 구매 전환율이 높을 것으로 봅니다.", "kpi": "구매 전환율, ROAS"},
    {"code": "L", "who": "40~60대 · 상황형", "head": "입맛 없는 날, 식당 전복죽 그대로", "sub": "완도산 통전복 한 마리, 내장까지 그대로",
     "badge": "통전복 한 마리", "cta": "2팩부터 무료배송",
     "hyp": "고객의 상황을 짚으면 관심 있는 사람만 클릭합니다. 클릭률은 A보다 낮아도 구매 전환율은 높을 것으로 봅니다.", "kpi": "클릭률, 구매 전환율"},
    {"code": "B", "who": "30대 · 아이 아침", "head": "아침 안 먹는 아이도 한 그릇 비우는 전복죽", "sub": "밥솥에 넣고 취사 한 번, 1팩이면 2인분",
     "badge": "취사 13분", "cta": "2팩부터 무료배송",
     "hyp": "\"1팩 2인분\"과 무료배송 조건을 함께 보여주면 2팩을 담는 고객이 늘어납니다. 객단가가 5만 원을 넘을 것으로 봅니다.", "kpi": "객단가, 구매 전환율"},
]
jjukkumi = [
    {"code": "F", "who": "30대 · 가족 저녁", "head": "매운 거 못 먹는 아이 있는 집도, 쭈꾸미 한 팬", "sub": "맵기 조절 소스 포함, 순한 맛부터 3단계",
     "badge": "맵기 3단계", "cta": "5만 원 이상 무료배송",
     "hyp": "아이가 있는 집은 맵기 때문에 구매를 망설입니다. 맵기 조절을 앞세우면 30대 구매 전환율이 오를 것으로 봅니다.", "kpi": "구매 전환율"},
    {"code": "G", "who": "40~60대 · 당류", "head": "설탕은 덜고, 유자와 매실로", "sub": "장호준 셰프의 쭈꾸미볶음 · 쭈꾸미 75%",
     "badge": "당류 100g당 1g 미만", "cta": "5만 원 이상 무료배송",
     "hyp": "40~60대는 단맛 부담을 덜어 주는 구체적인 숫자에 반응합니다. 맛을 강조한 소재보다 구매 전환율이 높을 것으로 봅니다.", "kpi": "구매 전환율, ROAS"},
    {"code": "P", "who": "30~40대 · 퇴근 후 안주", "head": "퇴근 후 맥주 한 잔엔, 볶기만 하면 되는 쭈꾸미", "sub": "콰삭오징어 하나 더 담으면 안주 완성",
     "badge": "쭈꾸미 75%", "cta": "5만 원 이상 무료배송",
     "hyp": "안주를 찾는 새 고객층이 있는지 확인하고, 콰삭오징어를 함께 담게 해 객단가를 높입니다.", "kpi": "새 고객층 구매 전환율, 객단가"},
]


def creative(label, title, items, band_text):
    s = frame("① 소재", label, title)
    cw = (W - 2 * 0.3) / 3
    for i, it in enumerate(items):
        x = MX + i * (cw + 0.3)
        pill(s, x, 1.9, 0.42, 0.3, it["code"], size=12)
        tb(s, x + 0.52, 1.9, cw - 0.52, 0.3, [[(it["who"], {})]], size=12, font=F_SB, color=SUB, anchor=MSO_ANCHOR.MIDDLE)
        gray_box(s, x, 2.32, 1.55, 1.55, "1:1 / 1200×628", f"소재_{it['code']}")
        rx, rw = x + 1.7, cw - 1.7
        tb(s, rx, 2.3, rw, 0.8, [[(it["head"], {})]], size=13, font=F_B, spacing=1.15)
        tb(s, rx, 3.1, rw, 0.42, [[(it["sub"], {})]], size=10, color=SUB, spacing=1.15)
        pill(s, rx, 3.55, rw, 0.25, it["badge"], fill_=CARD_HL, color=RED, size=10)
        tb(s, x, 3.95, cw, 0.25, [[("CTA  ", {"font": F_SB, "color": RED}), (it["cta"], {"font": F_SB})]], size=10)
        box(s, x, 4.28, cw, 2.02, fill_=CARD)
        tb(s, x + 0.2, 4.38, cw - 0.4, 0.26, [[("가설", {})]], size=11, font=F_B, color=RED)
        tb(s, x + 0.2, 4.66, cw - 0.4, 1.15, [[(it["hyp"], {})]], size=11, spacing=1.3)
        tb(s, x + 0.2, 5.85, cw - 0.4, 0.38, [[("볼 지표  ", {"font": F_SB, "color": RED}), (it["kpi"], {"font": F_SB})]], size=11, anchor=MSO_ANCHOR.MIDDLE)
    band(s, band_text)
    return s


creative("① 소재 · 게우 통전복죽", "셰프로 설득할지, 상황으로 부를지 비교합니다", jeonbok,
         "A와 L은 같은 40~60대 그룹에서 같은 예산으로, B는 30대 그룹에서 기존 소재와 비교합니다")
s = creative("① 소재 · 인생쭈꾸미", "쭈꾸미는 가격보다 먹는 장면으로 팝니다", jjukkumi,
             "원물이 외국산이라 제철·국내산 표현은 쓰지 않고, 상품명은 다른 업체와 겹치지 않게 \"장호준 쭈꾸미\"로 씁니다")
s.notes_slide.notes_text_frame.text = "쭈꾸미 CTA는 \"2팩부터\"가 아니라 \"5만 원 이상 무료배송\"입니다(쭈꾸미 2팩은 5만 원 미만)."

# 5. 테스트 운영
s = frame("① 소재", "① 소재 · 테스트 운영", "두 소재씩 맞붙이고, 14일 뒤 성과가 좋은 쪽을 남깁니다")
table(s, MX, 1.95, [1.75, 1.2, 3.75, 3.0, 2.63], [
    ["상품", "그룹", "비교", "바꾸는 요소", "볼 지표"],
    ["게우 통전복죽", "40~60대", "A 셰프형 vs L 상황형", "메시지 (셰프 신뢰 vs 고객 상황)", "구매 전환율, ROAS"],
    ["게우 통전복죽", "30대", "B vs 기존 \"아이 아침은 전복죽 한 팩으로\"", "헤드라인, 2인분 정보", "객단가, 구매 전환율"],
    ["인생쭈꾸미", "30대", "F vs 기존 \"맵기 3단계\" 소재", "대상 상황 (가족·아이)", "구매 전환율"],
    ["인생쭈꾸미", "40~60대", "G 단독, 2주 뒤 셰프형 소재와 비교", "숫자 근거 (당류)", "구매 전환율, ROAS"],
    ["인생쭈꾸미", "30~40대", "P 단독 (새 고객층)", "새 대상, 함께 구매", "새 고객층 구매 전환율, 객단가"],
], row_h=[0.42, 0.55, 0.55, 0.55, 0.55, 0.55], size=12, head_size=12)
b = box(s, MX, 5.4, W, 0.85, fill_=CARD)
stext(b, [[("비교하는 두 소재는 그룹, 예산, 기간을 똑같이 맞춥니다. 판정 전에는 소재와 입찰을 바꾸지 않고, 한 상품에서 동시에 하는 비교는 두 건까지만 둡니다.", {})]],
      size=12, ml=0.3, spacing=1.3)
band(s, "이긴 소재의 헤드라인은 검색광고 확장소재, 상품명, 카카오 메시지에도 그대로 씁니다")

# 6. 판정 기준
s = frame("② 판정 기준", "② 소재 유지·교체 판정 기준", "소재는 여섯 가지 기준으로 남기거나 바꿉니다")
b = box(s, MX, 1.92, W, 0.82, fill_=CARD_HL)
stext(b, [[("판정 시점  ", {"font": F_B, "color": RED}), ("소재마다 집행 14일 또는 클릭 300회 중 먼저 오는 때. 그 전에는 소재를 바꾸지 않습니다.", {})],
          [("기준선(10월 통영숙수 실적)  ", {"font": F_B, "color": RED}), ("피드 클릭률 1.1~1.8%, 스마트채널 클릭률 0.1%, ADVoost 전복죽 클릭 대비 구매 5.3%", {})]],
      size=12, ml=0.3, after=3)
judge = [["판정", "조건", "조치"],
         ["확대", "ROAS 300%(BEP) 이상", "예산을 20% 늘리고 2주 뒤 다시 판정"],
         ["유지", "ROAS 200~300%, 클릭률 기준선 이상", "그대로 두고 보조 문구만 A/B 테스트"],
         ["개선", "클릭률은 괜찮은데 구매 전환율이 낮음", "이미지는 두고 카피와 상세페이지 첫 화면 점검"],
         ["교체", "클릭률이 기준선의 70% 미만(피드 0.8% 미만), 또는 ROAS 150% 미만이면서 구매 3건 이하", "새 소재로 교체"],
         ["피로도 교체", "최근 7일 클릭률이 첫 7일보다 30% 이상 떨어짐", "메시지는 두고 이미지만 새로 제작"],
         ["승자 판정", "비교 소재가 모두 구매 10건 이상이고, 구매 전환율 차이가 20% 이상", "이긴 쪽을 기본안으로 두고 진 쪽은 교체"]]
table(s, MX, 2.88, [1.7, 5.9, 4.73], judge, row_h=[0.4] + [0.48] * 6, size=12, head_size=12,
      colors={(4, 0): WARN, (5, 0): WARN, (1, 0): RED})
band(s, "10월 통영숙수 수동 전환 캠페인(17일, ROAS 29%, 구매 10건)은 이 기준으로 보면 교체 대상입니다")

# 7. 비용 누수 Top 5
s = frame("③ 비용 누수", "③ 비용 누수 Top 5", "계정 전체에서 월 약 516만 원이 새고 있습니다")
R = PP_ALIGN.RIGHT
leak = [["순위", "범위", "누수 지점", "근거 (10/1~28)", "현재 상태", "월 절감 비용", "순효과"],
        ["1", "통영숙수", "GFA 수동 전환 캠페인", "176.5만 원 집행, ROAS 25% (전복죽 29%, 쭈꾸미 12%)", "전복죽 집행 중 (하루 약 8.1만 원)", "약 242만 원", "약 208만 원"],
        ["2", "계정 전체", "GFA 저효율 연령", "꼬기다 계열 24세 이하 ROAS 151%, 60세 이상 168% (핵심 연령 290%) / 통영숙수 60세 이상 85%, 연령 미상 35%", "집행 중", "약 209만 원", "약 122만 원"],
        ["3", "통영숙수", "쭈꾸미 쇼핑검색 일반어", "30.7만 원, ROAS 20% (\"쭈꾸미\" 단독 21만 원에 구매 1건)", "집행 중", "약 33만 원", "약 31만 원"],
        ["4", "계정 전체", "제외키워드 0개", "구매 없는 비브랜드 검색어(기장끝집전복죽, 호박죽 등)와 \"죽\" 단독(ROAS 94%)", "집행 중", "약 17만 원", "약 14만 원"],
        ["5", "계정 전체", "전환 측정 누락 파워링크", "건담닭·올라잇그린 34.9만 원, 구매 0건. 올라잇그린 클릭당 447원(꼬기다 약 100원)", "3분기부터 계속", "약 16만 원", "약 16만 원"],
        ["", "", [("합계", {"font": F_B})], "", "", [("약 516만 원", {"font": F_B, "color": RED})], [("약 390만 원", {"font": F_B, "color": RED})]]]
table(s, MX, 1.92, [0.55, 1.0, 1.85, 4.45, 1.85, 1.33, 1.3], leak, row_h=[0.4, 0.55, 0.78, 0.55, 0.55, 0.55, 0.4], size=11, head_size=11,
      aligns={0: PP_ALIGN.CENTER, 5: R, 6: R}, fills={6: CARD_HL})
tb(s, MX, 5.82, W, 0.58, [[("월 환산은 28일 실적에 30/28을 곱했고, 집행 중인 캠페인은 최근 7일 집행 속도로 계산했습니다. 순효과는 줄인 광고비에서 함께 줄어드는 매출의 3분의 1(BEP 300% 기준 이익)을 뺀 값입니다. 5번은 전환이 측정되면 나머지 약 37만 원의 효율도 판단할 수 있습니다.", {})]],
   size=10, color=SUB, spacing=1.25)
band(s, "장바구니 매출이 구매 매출의 2.1배(GFA 6,464만 원, 3,025만 원)라, 리타겟으로 되찾을 몫이 큽니다")

# 8. 개선 방법
s = frame("③ 비용 누수", "③ 비용 누수 · 개선 방법", "새는 곳은 막고, 아낀 예산은 효율이 확인된 곳으로 옮깁니다")
fix = [["누수 지점", "개선 방법", "아낀 예산을 옮길 곳"],
       ["1. 수동 전환 캠페인", "수동 전환 캠페인을 멈추고, 상세페이지를 보거나 장바구니에 담은 고객에게 다시 보여주는 카탈로그 리타겟으로 바꿉니다", "ADVoost 전복죽(10월 ROAS 155%)과 카탈로그 리타겟"],
       ["2. 저효율 연령", "24세 이하는 제외합니다. 60세 이상과 연령 미상은 입찰을 낮추고 2주 뒤 다시 판단합니다", "핵심 연령(25~59세) 그룹"],
       ["3. 쭈꾸미 일반어", "\"쭈꾸미\" 단독 키워드는 멈추고, \"장호준 쭈꾸미\"처럼 셰프 이름이 붙은 키워드로 바꿉니다", "쭈꾸미 GFA 신규 소재 테스트"],
       ["4. 제외키워드", "관련 없는 검색어와 타사 상품 검색어, \"죽\" 단독 키워드를 제외합니다. 제외할 후보 173개는 이미 뽑아 두었습니다(예: 기장끝집전복죽, 금채식탁전복죽, 호박죽, 수제호박죽, 하남쭈꾸미500G5팩). 전체 목록은 계약 후 첫 주에 등록하고, 이후 2주마다 검색어 보고서를 보고 추가합니다", "전복죽 상품 키워드"],
       ["5. 측정 누락", "건담닭·올라잇그린 자사몰에 전환 스크립트를 연결합니다. 연결 전까지는 올라잇그린 입찰을 클릭당 150원 안팎으로 낮춥니다", "측정이 연결된 뒤 효율을 보고 결정"]]
table(s, MX, 1.92, [2.1, 7.0, 3.23], fix, row_h=[0.4, 0.62, 0.55, 0.55, 1.0, 0.62], size=12, head_size=12)
band(s, "첫 달은 새는 돈을 막아 효율이 확인된 곳으로 옮기는 데 집중합니다")
s.notes_slide.notes_text_frame.text = "4번은 개수와 예시 5개까지만. 전체 목록은 계약 전에 공개하지 않습니다."

# 9. 점검 주기
s = frame("④ 상시 점검", "④ 상시 점검 방식", "매일, 매주, 매월 볼 것을 정해 두고 점검합니다")
cols = [("매일", "오전 10시까지", ["캠페인 켜짐·꺼짐, 학습 제한 상태", "하루 광고비와 구매를 최근 7일 평균과 비교", "예산이 일찍 소진되는지", "브랜드검색이 노출되고 있는지", "클릭은 나는데 구매가 0건인 캠페인 (측정 누락 의심)"]),
        ("매주", "월요일", ["소재 판정 (② 기준)", "검색어 보고서를 보고 제외키워드 추가", "연령·지면별 효율과 입찰 조정", "키워드 노출 순위와 클릭당 비용", "주간 리포트 발송"]),
        ("매월", "첫째 주", ["캠페인 구조 점검 (지나치게 쪼갠 캠페인, 겹치는 타겟)", "상품별 성과를 스토어 판매와 맞춰 보기", "전환 측정 점검 (스크립트·픽셀·전환 설정)", "다음 달 소재·예산 계획", "월간 리뷰 미팅"])]
cw = (W - 2 * 0.3) / 3
for i, (t, when, items) in enumerate(cols):
    x = MX + i * (cw + 0.3)
    box(s, x, 1.92, cw, 4.3, fill_=CARD)
    head = box(s, x, 1.92, cw, 0.62, fill_=NAVY, radius=0.08)
    stext(head, [[(t, {"font": F_B, "size": 17}), (f"   {when}", {"color": COMP})]], size=12, color=WHITE, ml=0.25)
    tb(s, x + 0.25, 2.75, cw - 0.5, 3.35, [[(it, {})] for it in items], size=13, bullet=True, after=8, spacing=1.25)
band(s, "매니저가 매일 점검하고, 팀장이 매주 직접 검토합니다. 이상 신호는 당일 바로 알려드립니다")

# 10. 이상 신호 기준
s = frame("④ 상시 점검", "④ 이상 신호 기준", "8가지 이상 신호 중 하나라도 보이면 바로 움직입니다")
sig = [["지표", "이상 신호", "먼저 할 일", "공유"],
       ["하루 광고비", "최근 7일 평균보다 50% 이상 많거나 적음", "예산·입찰·경쟁 상황 확인", "당일"],
       ["전환 측정", "클릭 100회 이상인데 구매 0건이 3일 이어짐", "전환 추적·스크립트 점검", "당일"],
       ["브랜드검색", "노출 0건, 또는 계약 종료 7일 전", "재계약 확인. 연장하지 않으면 파워링크 브랜드 키워드에 입찰 상한 설정", "즉시"],
       ["캠페인 상태", "학습 제한, 운영 불가, 예산 조기 소진", "예산·타겟 조정", "당일"],
       ["클릭당 비용", "최근 7일 평균보다 30% 이상 오름", "입찰과 경쟁 광고 확인", "주간"],
       ["캠페인 효율", "최근 2주 ROAS 150% 미만", "예산 절반으로 축소, 소재·키워드 교체", "주간"],
       ["소재", "클릭률이 기준선의 70% 미만", "② 기준에 따라 교체", "주간"],
       ["스토어 전환율", "3% 미만 (9월 4.9%, 10월 초 2.3%)", "유입 품질과 상세페이지 점검", "주간"]]
share_col = {}
for r in range(1, len(sig)):
    share_col[(r, 3)] = RED if sig[r][3] in ("당일", "즉시") else SUB
table(s, MX, 1.92, [1.9, 4.35, 4.93, 1.15], sig, row_h=[0.4] + [0.47] * 8, size=12, head_size=12,
      aligns={3: PP_ALIGN.CENTER}, colors=share_col)
band(s, "기준값은 이번 계정 데이터로 정했고, 운영하면서 2주마다 다시 맞춥니다")

# 11. 실제 사례
s = frame("④ 상시 점검", "④ 실제 사례", "꼬기다 브랜드검색이 끝난 뒤의 변화를 3일 만에 확인했습니다")
days = ["10/19", "10/20", "10/21", "10/22", "10/23", "10/24", "10/25", "10/26", "10/27", "10/28"]
spend = [61446, 50330, 39906, 30794, 30815, 23755, 27695, 116504, 108038, 89916]
buys = [464, 281, 339, 310, 244, 227, 280, 331, 217, 120]
cw = (W - 0.3) / 2
for i, (title, vals, fmt) in enumerate([("꼬기다 파워링크 하루 광고비", spend, "#,##0"), ("꼬기다 브랜드 관련 구매 (파워링크 + 브랜드검색)", buys, "#,##0")]):
    x = MX + i * (cw + 0.3)
    box(s, x, 1.9, cw, 2.55, fill_=WHITE, line=LINE)
    cd = CategoryChartData()
    cd.categories = days
    cd.add_series(title, vals)
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x + 0.1), Inches(1.95), Inches(cw - 0.2), Inches(2.45), cd).chart
    ch.font.name = F_R
    ch.font.size = Pt(9)
    ch.font.color.rgb = rgb(SUB)
    ch.has_title = True
    fill(ch.chart_title.text_frame, [[(title + (" (원)" if i == 0 else " (건)"), {})]], size=11, font=F_SB)
    ch.has_legend = False
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.value_axis.minimum_scale = 0
    ch.category_axis.format.line.color.rgb = rgb(LINE)
    pl = ch.plots[0]
    pl.gap_width = 45
    pl.has_data_labels = True
    pl.data_labels.number_format = fmt
    pl.data_labels.number_format_is_linked = False
    pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    pl.data_labels.font.size = Pt(8)
    for k in range(len(days)):
        pt = pl.series[0].points[k]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(RED if k >= 7 else COMP)
b = box(s, MX, 4.55, W, 0.4, fill_=CARD_HL)
stext(b, [[("10월 26일부터 브랜드검색 노출 0건 (계약 10월 25일 종료)", {"font": F_B, "color": RED})]], size=12, ml=0.25)
table(s, MX, 5.05, [3.1, 5.05, 4.18], [
    ["어떤 기준으로 잡았나", "숫자", "어떻게 대응하나"],
    ["브랜드검색 노출 0건 (매일 점검)", "10월 26일부터 노출 0건", "재계약 여부를 바로 확인"],
    ["하루 광고비 ±50% (매일 점검)", "꼬기다 파워링크 하루 3.8만 원 → 10.5만 원 (2.8배), 브랜드 관련 구매 하루 306건 → 223건 (27% 감소)",
     "연장하지 않으면 브랜드 키워드에 입찰 상한을 걸고, 브랜드 구매 추이를 매일 보고"],
], row_h=[0.32, 0.36, 0.62], size=10, head_size=10)
band(s, "정해 둔 기준이 있어서 변화가 생긴 지 사흘 만에 원인과 영향을 숫자로 잡았습니다")

# 12. 운영 조건
s = frame("운영·일정", "운영 조건", "팀장이 매주 직접 들여다보는 2인 전담 체제입니다")
blocks = [("2인 전담", [[("김현수 팀장 · 14년차  ", {"font": F_B}), ("전략, 주간 점검과 소재 판정, 리포트, 월간 리뷰", {})],
                       [("김범준 매니저 · 1년차  ", {"font": F_B}), ("매일 점검, 세팅, 소재 반영, 검색어·제외키워드 관리", {})]]),
          ("운영 계정 수", [[("담당 매니저 기준 약 150개", {"font": F_B})],
                          [("대부분 월 광고비가 크지 않은 검색광고 위주 계정입니다. 브레이브컴퍼니는 팀장이 직접 관리하는 계정으로 지정해, 매일 이상 신호를 점검하고 매주 팀장이 검토합니다.", {})]]),
          ("소재 제작", [[("사내 디자인팀, 월 40개 이상 (GFA·카카오 규격 포함)", {"font": F_B})],
                       [("통영숙수에 필요한 양: 상품별 3종을 2주마다 바꾸면 월 약 12개", {})]]),
          ("소통", [[("주간 리포트(매주 월요일), 이상 신호 당일 공유, 월간 리뷰 미팅", {"font": F_B})],
                  [("운영 계정이 많아도 ④의 이상 신호 기준으로 빠짐없이 관리합니다", {})]])]
cw, chh = (W - 0.3) / 2, 2.05
for i, (t, paras) in enumerate(blocks):
    x, y = MX + (i % 2) * (cw + 0.3), 1.92 + (i // 2) * (chh + 0.25)
    box(s, x, y, cw, chh, fill_=CARD)
    badge(s, x + 0.25, y + 0.22, i + 1)
    tb(s, x + 0.8, y + 0.22, cw - 1.0, 0.4, [[(t, {})]], size=16, font=F_B, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, x + 0.8, y + 0.75, cw - 1.05, chh - 0.85, paras, size=12, spacing=1.3, after=5)
band(s, "담당자 감에 맡기지 않고, 정해 둔 기준과 점검 체계로 계정을 지킵니다")

# 13. 계약 후 첫 2주
s = frame("운영·일정", "실행 일정", "계약 다음 날부터 바로 시작할 수 있게 준비해 두었습니다")
table(s, MX, 1.95, [2.0, 8.83, 1.5], [
    ["시점", "할 일", "관련 과제"],
    ["1주차 1~2일", "계정 점검, 제외키워드 173개 등록, \"쭈꾸미\" 단독 키워드 정리", "③"],
    ["1주차 3~5일", "수동 전환 캠페인 정리와 카탈로그 리타겟 세팅, 24세 이하 제외와 고연령 입찰 조정, 건담닭·올라잇그린 스크립트 연결 요청", "③"],
    ["2주차", "디자인팀 신규 소재 6종 투입, 상품별 A/B 테스트 시작", "①②"],
    ["2주차 말", "첫 주간 리포트: 누수 정리 결과와 절감액, 테스트 진행 현황", "④"],
    ["4주차", "첫 소재 판정, 다음 달 예산 배분안 공유", "②④"],
], row_h=[0.42, 0.62, 0.78, 0.62, 0.62, 0.62], size=13, head_size=13, aligns={2: PP_ALIGN.CENTER},
   colors={(r, 2): RED for r in range(1, 6)})
band(s, "준비는 끝났습니다. 계약이 확정되면 첫 주부터 새는 비용을 막겠습니다")

# 14. 마무리
s = navy_slide()
tb(s, 1.1, 2.3, 11, 1.0, [[("감사합니다", {})]], size=44, font=F_B, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
tb(s, 1.1, 3.3, 11, 0.5, [[("THANK YOU", {})]], size=22, font=F_SB, color=COMP, anchor=MSO_ANCHOR.MIDDLE)
hline(s, 1.1, 4.1, 4.2, RED, 1.5)
tb(s, 1.1, 4.3, 11, 0.5, [[("자료로 전하기 어려운 부분은 30분 화상 미팅으로 김현수 팀장이 직접 설명드리겠습니다.", {})]], size=15, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)

OUT.parent.mkdir(exist_ok=True)
prs.save(str(OUT))
print("saved", OUT, len(prs.slides._sldIdLst))
