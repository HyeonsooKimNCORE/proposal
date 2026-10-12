"""data/ 원본 CSV로 제안서 핵심 숫자를 다시 계산한다.  실행: python3 verify.py"""
import collections
import csv
import datetime as dt
from pathlib import Path

D = Path(__file__).parent / "data"


def num(x):
    return float(str(x).replace(",", "") or 0)


def read(name, skip=0):
    with open(D / name, encoding="utf-8-sig") as f:
        lines = f.readlines()[skip:]
    return list(csv.DictReader(lines))


def date(s):
    return dt.date(int(s[:4]), int(s[5:7]), int(s[8:10]))


# ---------- SA 캠페인 일별: 계정 ROAS, 브랜드검색 노출, 파워링크 대체 효과
sa = read("SA_캠페인일별_2026Q3.csv", 1)
cost = sum(num(r["총비용"]) for r in sa)
rev = sum(num(r["구매완료 전환매출액(원)"]) for r in sa)
days = [dt.date(2026, 7, 1) + dt.timedelta(i) for i in range(92)]
bs = {date(r["일별"]): r for r in sa if r["캠페인"] == "꼬기다 브랜드 검색광고"}
exposed = {d for d in days if d in bs and num(bs[d]["노출수"]) > 0}
contract = [d for d in days if dt.date(2026, 7, 28) <= d]
daily_fee = 11_550_000 / 90
fee = daily_fee * len(contract)
bs_rev = sum(num(bs[d]["구매완료 전환매출액(원)"]) for d in contract if d in bs)
bs_clicks = sum(num(bs[d]["클릭수"]) for d in contract if d in bs)
pl = {date(r["일별"]): r for r in sa if r["캠페인"] == "꼬기다_파워링크"}


def pl_avg(ds):
    ds = [d for d in ds if d in pl]
    return (sum(num(pl[d]["클릭수"]) for d in ds) / len(ds), sum(num(pl[d]["총비용"]) for d in ds) / len(ds))


miss_camps = ("건담닭_파워링크", "올라잇그린_파워링크")
print("[SA]")
print(f"  광고비 {cost:,.0f}원  ROAS {rev / cost:.0%}  계약금 포함 {rev / (cost + fee):.0%}")
print(f"  계약 {len(contract)}일 중 노출 {len(exposed & set(contract))}일, 미노출 계약금 {daily_fee * len(set(contract) - exposed):,.0f}원")
print(f"  브랜드검색 실질 ROAS {bs_rev / fee:.0%}, 클릭당 {fee / bs_clicks:,.0f}원")
print("  꼬기다 파워링크 하루 클릭·비용  노출일 %.0f회·%.0f원 / 미노출일 %.0f회·%.0f원"
      % (*pl_avg(exposed), *pl_avg(set(days) - exposed)))
print(f"  측정 누락 비중 {sum(num(r['총비용']) for r in sa if r['캠페인'] in miss_camps) / cost:.0%}")

# ---------- 키워드 보고서: 브랜드 키워드 그룹 비교
kw = read("SA_키워드보고서_2026Q3.csv", 1)
print("[파워링크 브랜드 키워드 그룹]")
for camp in ("꼬기다_파워링크", "건담닭_파워링크", "올라잇그린_파워링크"):
    rs = [r for r in kw if r["캠페인"] == camp and r["광고그룹"] == "01. 브랜드 키워드"]
    i, c, co = (sum(num(r[k]) for r in rs) for k in ("노출수", "클릭수", "총비용"))
    print(f"  {camp}: 클릭률 {c / i:.1%}, 클릭당 {co / c:,.0f}원")
rs = [r for r in kw if r["캠페인"] in miss_camps]
print(f"  측정 누락: {sum(num(r['총비용']) for r in rs):,.0f}원, 클릭 {sum(num(r['클릭수']) for r in rs):,.0f}, 구매 {sum(num(r['구매완료 전환수']) for r in rs):.0f}")

# ---------- 검색어 보고서: 유입 검색어, 구매 0건 검색어
st = read("SA_검색어보고서_2026Q3.csv", 1)
agg = collections.defaultdict(lambda: [0.0, 0.0])
for r in st:
    if r["캠페인"] in miss_camps:
        continue
    a = agg[(r["캠페인"], r["검색어"])]
    a[0] += num(r["총비용"])
    a[1] += num(r["구매완료 전환수"])
zero = [v for v in agg.values() if v[1] == 0 and v[0] > 0]
print("[검색어]")
print(f"  서로 다른 검색어 {len({r['검색어'] for r in st}):,}개, 구매 0건 {len(zero)}개 {sum(v[0] for v in zero):,.0f}원")

# ---------- GFA 일별: 월별 추이와 증액 효율
gfa = read("GFA_일별_2026-07.csv") + read("GFA_일별_2026-08_09.csv")
m = collections.defaultdict(lambda: [0.0] * 4)
for r in gfa:
    a = m[r["기간"][5:7]]
    for i, k in enumerate(("총비용", "노출수", "구매완료 수", "구매완료 전환매출액")):
        a[i] += num(r[k])
print("[GFA 월별]  5~6월(52일) 8,847,295원 / ROAS 364.22% 기준")
for mo, (c, i, p, v) in sorted(m.items()):
    print(f"  {mo}월 광고비 {c:,.0f}  CPM {c / i * 1000:,.0f}  ROAS {v / c:.0%}  구매 {p:.0f}")
base = 8_847_295 / 52 * 30
avg_cost = sum(a[0] for a in m.values()) / 3
avg_rev = sum(a[3] for a in m.values()) / 3
print(f"  월 광고비 {base / 1e4:,.0f}만 → {avg_cost / 1e4:,.0f}만 원 ({avg_cost / base:.1f}배), "
      f"늘린 돈의 효율 {(avg_rev - base * 3.6422) / (avg_cost - base):.0%}")

# ---------- ADVoost 애셋 그룹 (총비용 열 앞에 캠페인 ID가 있어 열 이름 대신 위치로 읽음)
with open(D / "GFA_ADVoost_애셋그룹_2026Q3.csv", encoding="utf-8-sig") as f:
    adv = list(csv.reader(f))[1:]
print("[ADVoost]")
for r in adv:
    print(f"  {r[2]}: {num(r[4]):,.0f}원, ROAS {num(r[9]) / num(r[4]):.0%}")
tc, tr = sum(num(r[4]) for r in adv), sum(num(r[9]) for r in adv)
print(f"  합계 {tc:,.0f}원, ROAS {tr / tc:.0%}")

# ---------- GFA 연령·성별
ag = read("GFA_연령성별_2026Q3.csv")
total = sum(num(r["총비용"]) for r in ag)
pick = lambda keys: [r for r in ag if r["연령 및 성별"].strip() in keys]
for label, keys in (("30대 남성", ("30세 ~ 34세 남성", "35세 ~ 39세 남성")), ("45~49세 여성", ("45세 ~ 49세 여성",)), ("연령 알 수 없음", ("알 수 없음",))):
    rs = pick(keys)
    c, v = sum(num(r["총비용"]) for r in rs), sum(num(r["구매완료 전환매출액"]) for r in rs)
    print(f"[연령·성별] {label}: 광고비 비중 {c / total:.0%}, ROAS {v / c:.0%}")
