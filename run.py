# ============================================================
# 사방넷 매출 자동 가공 스크립트 v2
# - 쇼핑몰별 파일 + 일별 파일 동시 처리
# - output/index.html 고정 출력 (GitHub Pages 자동 반영)
# 사용법: input 폴더에 사방넷 다운로드 파일 저장 후 실행
# ============================================================

import pandas as pd
import os, glob, json, re
from datetime import datetime

# ── 경로 설정 ──────────────────────────────────────────────
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR    = os.path.join(BASE_DIR, "input")
OUTPUT_DIR   = os.path.join(BASE_DIR, "output")
HISTORY_FILE = os.path.join(BASE_DIR, "history.json")
OUTPUT_FILE  = os.path.join(OUTPUT_DIR, "index.html")   # 항상 index.html 덮어쓰기

os.makedirs(INPUT_DIR,  exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── 공통 유틸 ──────────────────────────────────────────────
def to_int(val):
    try:
        return int(str(val).replace(",", "").replace(" ", "").replace("\xa0", ""))
    except:
        return 0

def fmt(n):
    return f"{n:,}"

def get_latest_file(pattern_keywords):
    """input 폴더에서 키워드를 포함하는 가장 최근 파일 반환"""
    all_files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx")) + \
                glob.glob(os.path.join(INPUT_DIR, "*.xls"))
    matched = [f for f in all_files
               if any(k in os.path.basename(f) for k in pattern_keywords)]
    if not matched:
        return None
    return sorted(matched, key=os.path.getmtime, reverse=True)[0]

# ── 파일 탐색 ──────────────────────────────────────────────
# 쇼핑몰별: "쇼핑몰" 키워드 포함 파일
# 일별:     "일별" 키워드 포함 파일
# 둘 다 없으면 가장 최근 파일 단독 처리

file_shop = get_latest_file(["쇼핑몰"])
file_day  = get_latest_file(["일별"])

# 키워드 미포함 시 전체 최신 파일로 폴백
if not file_shop and not file_day:
    all_files = sorted(
        glob.glob(os.path.join(INPUT_DIR, "*.xlsx")) +
        glob.glob(os.path.join(INPUT_DIR, "*.xls")),
        key=os.path.getmtime, reverse=True
    )
    if not all_files:
        print("❌ input 폴더에 xlsx 또는 xls 파일이 없습니다.")
        print(f"   경로 확인: {INPUT_DIR}")
        input("\nEnter를 눌러 닫기...")
        exit()
    file_shop = all_files[0]

print("=" * 50)
print("  사방넷 매출 대시보드 생성기 v2")
print("=" * 50)
if file_shop: print(f"✅ 쇼핑몰별 파일: {os.path.basename(file_shop)}")
if file_day:  print(f"✅ 일별 파일:     {os.path.basename(file_day)}")

# ── 날짜 추출 ──────────────────────────────────────────────
def extract_date(filename):
    m = re.search(r"(\d{8})", filename)
    if m:
        d = m.group(1)
        return f"{d[:4]}-{d[4:6]}-{d[6:]}"
    return datetime.today().strftime("%Y-%m-%d")

report_date = extract_date(os.path.basename(file_shop or file_day))
print(f"   집계 기준일: {report_date}\n")

# ── 쇼핑몰별 파싱 ──────────────────────────────────────────
shop_data = []
if file_shop:
    df = pd.read_excel(file_shop, header=None)
    for i, row in df.iterrows():
        if i < 3: continue
        val = str(row.iloc[0]).strip()
        if val in ["합 계", "합계", "NaN", "nan", ""]: continue
        if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() != "":
            shop_data.append({
                "쇼핑몰명":     str(row.iloc[1]).strip(),
                "주문수량":     to_int(row.iloc[2]),
                "주문금액":     to_int(row.iloc[3]),
                "취소수량":     to_int(row.iloc[4]),
                "취소금액":     to_int(row.iloc[5]),
                "반품수량":     to_int(row.iloc[6]),
                "반품금액":     to_int(row.iloc[7]),
                "순매출수량":   to_int(row.iloc[8]),
                "순매출금액":   to_int(row.iloc[9]),
                "총이익액":     to_int(row.iloc[10]),
                "판매수수료":   to_int(row.iloc[12]),
                "정산예정금액": to_int(row.iloc[16]),
            })
    print(f"   쇼핑몰 수: {len(shop_data)}개 인식")

# ── 일별 파싱 ──────────────────────────────────────────────
day_data = []
if file_day:
    df = pd.read_excel(file_day, header=None)
    for i, row in df.iterrows():
        if i < 3: continue
        val = str(row.iloc[0]).strip()
        if val in ["합 계", "합계", "NaN", "nan", ""]: continue
        if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() != "":
            day_data.append({
                "일자":         str(row.iloc[1]).strip(),
                "요일":         str(row.iloc[2]).strip(),
                "주문수량":     to_int(row.iloc[3]),
                "주문금액":     to_int(row.iloc[4]),
                "취소수량":     to_int(row.iloc[5]),
                "취소금액":     to_int(row.iloc[6]),
                "반품수량":     to_int(row.iloc[7]),
                "반품금액":     to_int(row.iloc[8]),
                "순매출수량":   to_int(row.iloc[9]),
                "순매출금액":   to_int(row.iloc[10]),
                "총이익액":     to_int(row.iloc[11]),
                "판매수수료":   to_int(row.iloc[13]),
                "정산예정금액": to_int(row.iloc[17]),
            })
    # 날짜 오름차순 정렬
    day_data.sort(key=lambda x: x["일자"])
    print(f"   일별 행 수: {len(day_data)}일 인식")

# ── 히스토리 저장 ──────────────────────────────────────────
history = {}
if os.path.exists(HISTORY_FILE):
    with open(HISTORY_FILE, "r", encoding="utf-8") as f:
        history = json.load(f)
history[report_date] = {"shop": shop_data, "day": day_data}
with open(HISTORY_FILE, "w", encoding="utf-8") as f:
    json.dump(history, f, ensure_ascii=False, indent=2)

# ── 집계 ───────────────────────────────────────────────────
src = shop_data if shop_data else day_data
total_order_qty  = sum(r["주문수량"]     for r in src)
total_order_amt  = sum(r["주문금액"]     for r in src)
total_cancel_amt = sum(r["취소금액"]     for r in src)
total_return_amt = sum(r["반품금액"]     for r in src)
total_net_amt    = sum(r["순매출금액"]   for r in src)
total_settle_amt = sum(r["정산예정금액"] for r in src)
settle_ratio     = f"{total_settle_amt/total_net_amt*100:.1f}%" if total_net_amt else "-"

# ── 쇼핑몰별 테이블 HTML ───────────────────────────────────
shop_table_html = ""
if shop_data:
    for r in shop_data:
        net = r["순매출금액"]; settle = r["정산예정금액"]
        ratio = f"{settle/net*100:.1f}%" if net else "-"
        shop_table_html += f"""
        <tr>
          <td>{r['쇼핑몰명']}</td>
          <td class="num">{fmt(r['주문수량'])}</td>
          <td class="num">{fmt(r['주문금액'])}</td>
          <td class="num red">{fmt(r['취소금액'])}</td>
          <td class="num red">{fmt(r['반품금액'])}</td>
          <td class="num bold">{fmt(r['순매출금액'])}</td>
          <td class="num">{fmt(r['판매수수료'])}</td>
          <td class="num purple">{fmt(r['정산예정금액'])}</td>
          <td class="num">{ratio}</td>
        </tr>"""
    shop_table_html += f"""
        <tr class="total-row">
          <td>합 계</td>
          <td class="num">{fmt(total_order_qty)}</td>
          <td class="num">{fmt(total_order_amt)}</td>
          <td class="num red">{fmt(total_cancel_amt)}</td>
          <td class="num red">{fmt(total_return_amt)}</td>
          <td class="num bold">{fmt(total_net_amt)}</td>
          <td class="num">-</td>
          <td class="num purple">{fmt(total_settle_amt)}</td>
          <td class="num">{settle_ratio}</td>
        </tr>"""

# ── 일별 테이블 HTML ───────────────────────────────────────
day_table_html = ""
if day_data:
    for r in day_data:
        net = r["순매출금액"]; settle = r["정산예정금액"]
        ratio = f"{settle/net*100:.1f}%" if net else "-"
        day_table_html += f"""
        <tr>
          <td>{r['일자']}</td>
          <td class="center">{r['요일']}</td>
          <td class="num">{fmt(r['주문수량'])}</td>
          <td class="num">{fmt(r['주문금액'])}</td>
          <td class="num red">{fmt(r['취소금액'])}</td>
          <td class="num red">{fmt(r['반품금액'])}</td>
          <td class="num bold">{fmt(r['순매출금액'])}</td>
          <td class="num">{fmt(r['판매수수료'])}</td>
          <td class="num purple">{fmt(r['정산예정금액'])}</td>
          <td class="num">{ratio}</td>
        </tr>"""
    day_net  = sum(r["순매출금액"]   for r in day_data)
    day_set  = sum(r["정산예정금액"] for r in day_data)
    day_qty  = sum(r["주문수량"]     for r in day_data)
    day_amt  = sum(r["주문금액"]     for r in day_data)
    day_can  = sum(r["취소금액"]     for r in day_data)
    day_ret  = sum(r["반품금액"]     for r in day_data)
    day_fee  = sum(r["판매수수료"]   for r in day_data)
    day_ratio = f"{day_set/day_net*100:.1f}%" if day_net else "-"
    day_table_html += f"""
        <tr class="total-row">
          <td>합 계</td><td></td>
          <td class="num">{fmt(day_qty)}</td>
          <td class="num">{fmt(day_amt)}</td>
          <td class="num red">{fmt(day_can)}</td>
          <td class="num red">{fmt(day_ret)}</td>
          <td class="num bold">{fmt(day_net)}</td>
          <td class="num">{fmt(day_fee)}</td>
          <td class="num purple">{fmt(day_set)}</td>
          <td class="num">{day_ratio}</td>
        </tr>"""

# ── Chart.js 데이터 ────────────────────────────────────────
shop_labels  = json.dumps([r["쇼핑몰명"]     for r in shop_data], ensure_ascii=False)
shop_net_j   = json.dumps([r["순매출금액"]   for r in shop_data])
shop_set_j   = json.dumps([r["정산예정금액"] for r in shop_data])

day_labels_j = json.dumps([r["일자"]         for r in day_data], ensure_ascii=False)
day_net_j    = json.dumps([r["순매출금액"]   for r in day_data])
day_set_j    = json.dumps([r["정산예정금액"] for r in day_data])

# ── 탭 렌더 조건 ───────────────────────────────────────────
show_shop = "block" if shop_data else "none"
show_day  = "block" if day_data  else "none"
active_shop = "active" if shop_data else ""
active_day  = "active" if (day_data and not shop_data) else ""

# ── HTML 생성 ──────────────────────────────────────────────
filenames_used = ", ".join(filter(None, [
    os.path.basename(file_shop) if file_shop else "",
    os.path.basename(file_day)  if file_day  else "",
]))

html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>매출 현황 대시보드 {report_date}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --bg:#F5F6FA;--surface:#fff;--border:#E4E6EF;
  --t1:#1A1D2E;--t2:#6B7280;
  --blue:#3B82F6;--green:#10B981;--red:#EF4444;--purple:#8B5CF6;--amber:#F59E0B;
  --r:12px;
}}
body{{font-family:'Pretendard',-apple-system,sans-serif;background:var(--bg);color:var(--t1);padding:28px 24px 60px;min-height:100vh}}

/* 헤더 */
.header{{display:flex;align-items:baseline;gap:10px;margin-bottom:24px;flex-wrap:wrap}}
.header h1{{font-size:21px;font-weight:700;letter-spacing:-.4px}}
.badge{{font-size:12px;font-weight:500;color:var(--t2);background:var(--border);padding:3px 10px;border-radius:20px}}
.fname{{font-size:11px;color:var(--t2);margin-left:auto}}

/* KPI */
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:22px}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:16px 18px}}
.kpi .lbl{{font-size:10px;font-weight:600;color:var(--t2);text-transform:uppercase;letter-spacing:.6px;margin-bottom:7px}}
.kpi .val{{font-size:19px;font-weight:700;letter-spacing:-.5px;line-height:1}}
.kpi .sub{{font-size:11px;color:var(--t2);margin-top:4px}}
.kpi.k-blue{{border-top:3px solid var(--blue)}}
.kpi.k-red{{border-top:3px solid var(--red)}}
.kpi.k-green{{border-top:3px solid var(--green)}}
.kpi.k-purple{{border-top:3px solid var(--purple)}}

/* 탭 */
.tabs{{display:flex;gap:4px;margin-bottom:16px}}
.tab-btn{{padding:8px 18px;font-size:13px;font-weight:600;border:1px solid var(--border);border-radius:8px;background:var(--surface);color:var(--t2);cursor:pointer;transition:all .15s}}
.tab-btn.active{{background:var(--t1);color:#fff;border-color:var(--t1)}}

/* 차트 그리드 */
.chart-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
@media(max-width:800px){{.chart-grid{{grid-template-columns:1fr}}}}
.card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px}}
.card-title{{font-size:12px;font-weight:600;color:var(--t2);margin-bottom:14px}}
.chart-wrap{{position:relative;height:240px}}

/* 테이블 */
.tbl-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px;overflow-x:auto;margin-bottom:14px}}
table{{width:100%;border-collapse:collapse;font-size:12.5px}}
thead th{{font-size:10px;font-weight:600;color:var(--t2);text-transform:uppercase;letter-spacing:.5px;padding:0 8px 9px;border-bottom:2px solid var(--border);white-space:nowrap;text-align:left}}
thead th.num{{text-align:right}}
tbody tr{{border-bottom:1px solid var(--border);transition:background .12s}}
tbody tr:last-child{{border-bottom:none}}
tbody tr:hover{{background:#F8F9FF}}
tbody td{{padding:10px 8px;vertical-align:middle}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
td.center{{text-align:center}}
td.red{{color:var(--red)}}
td.purple{{color:var(--purple);font-weight:600}}
td.bold{{font-weight:600}}
tr.total-row{{font-weight:700;border-top:2px solid var(--border)!important}}

.footer{{margin-top:20px;text-align:center;font-size:11px;color:var(--t2)}}
.panel{{display:none}}.panel.active{{display:block}}
</style>
</head>
<body>

<div class="header">
  <h1>📊 매출 현황 대시보드</h1>
  <span class="badge">{report_date}</span>
  <span class="fname">{filenames_used}</span>
</div>

<!-- KPI -->
<div class="kpi-grid">
  <div class="kpi k-blue">
    <div class="lbl">주문금액</div>
    <div class="val">₩{fmt(total_order_amt)}</div>
    <div class="sub">{fmt(total_order_qty)}건</div>
  </div>
  <div class="kpi k-red">
    <div class="lbl">취소·반품</div>
    <div class="val">₩{fmt(total_cancel_amt + total_return_amt)}</div>
    <div class="sub">취소 {fmt(total_cancel_amt)} / 반품 {fmt(total_return_amt)}</div>
  </div>
  <div class="kpi k-green">
    <div class="lbl">순매출금액</div>
    <div class="val">₩{fmt(total_net_amt)}</div>
    <div class="sub">&nbsp;</div>
  </div>
  <div class="kpi k-purple">
    <div class="lbl">정산예정금액</div>
    <div class="val">₩{fmt(total_settle_amt)}</div>
    <div class="sub">순매출 대비 {settle_ratio}</div>
  </div>
</div>

<!-- 탭 버튼 -->
<div class="tabs">
  <button class="tab-btn {active_shop}" onclick="switchTab('shop')" id="btn-shop" style="display:{show_shop}">쇼핑몰별</button>
  <button class="tab-btn {active_day}"  onclick="switchTab('day')"  id="btn-day"  style="display:{show_day}">일별</button>
</div>

<!-- 쇼핑몰별 패널 -->
<div class="panel {"active" if shop_data else ""}" id="panel-shop" style="display:{show_shop}">
  <div class="chart-grid">
    <div class="card"><div class="card-title">쇼핑몰별 순매출금액</div><div class="chart-wrap"><canvas id="cShopNet"></canvas></div></div>
    <div class="card"><div class="card-title">쇼핑몰별 정산예정금액</div><div class="chart-wrap"><canvas id="cShopSettle"></canvas></div></div>
  </div>
  <div class="tbl-card">
    <div class="card-title">쇼핑몰별 상세 내역</div>
    <table>
      <thead><tr>
        <th>쇼핑몰</th><th class="num">주문수량</th><th class="num">주문금액</th>
        <th class="num">취소금액</th><th class="num">반품금액</th>
        <th class="num">순매출금액</th><th class="num">판매수수료</th>
        <th class="num">정산예정금액</th><th class="num">정산율</th>
      </tr></thead>
      <tbody>{shop_table_html}</tbody>
    </table>
  </div>
</div>

<!-- 일별 패널 -->
<div class="panel {"active" if (day_data and not shop_data) else ""}" id="panel-day" style="display:{show_day}">
  <div class="chart-grid">
    <div class="card"><div class="card-title">일별 순매출금액</div><div class="chart-wrap"><canvas id="cDayNet"></canvas></div></div>
    <div class="card"><div class="card-title">일별 정산예정금액</div><div class="chart-wrap"><canvas id="cDaySettle"></canvas></div></div>
  </div>
  <div class="tbl-card">
    <div class="card-title">일별 상세 내역</div>
    <table>
      <thead><tr>
        <th>일자</th><th class="center">요일</th><th class="num">주문수량</th><th class="num">주문금액</th>
        <th class="num">취소금액</th><th class="num">반품금액</th>
        <th class="num">순매출금액</th><th class="num">판매수수료</th>
        <th class="num">정산예정금액</th><th class="num">정산율</th>
      </tr></thead>
      <tbody>{day_table_html}</tbody>
    </table>
  </div>
</div>

<div class="footer">생성: {datetime.now().strftime("%Y-%m-%d %H:%M")} · {filenames_used}</div>

<script>
const COLORS=['#FEE500','#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#06B6D4','#EC4899'];

function makeBar(id, labels, data, label){{
  const el = document.getElementById(id);
  if(!el) return;
  new Chart(el.getContext('2d'),{{
    type:'bar',
    data:{{
      labels,
      datasets:[{{
        label, data,
        backgroundColor: labels.map((_,i)=>COLORS[i%COLORS.length]+'BB'),
        borderColor:     labels.map((_,i)=>COLORS[i%COLORS.length]),
        borderWidth:1.5, borderRadius:5,
      }}]
    }},
    options:{{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{callbacks:{{label:c=>'₩'+c.parsed.y.toLocaleString()}}}}
      }},
      scales:{{
        y:{{
          ticks:{{callback:v=>v>=1e8?(v/1e8).toFixed(1)+'억':v>=1e4?(v/1e4).toFixed(0)+'만':v, font:{{size:10}}}},
          grid:{{color:'#E4E6EF'}}
        }},
        x:{{ticks:{{font:{{size:10}}}},grid:{{display:false}}}}
      }}
    }}
  }});
}}

// 쇼핑몰별 차트
const shopLabels = {shop_labels};
makeBar('cShopNet',    shopLabels, {shop_net_j},  '순매출금액');
makeBar('cShopSettle', shopLabels, {shop_set_j}, '정산예정금액');

// 일별 차트
const dayLabels = {day_labels_j};
makeBar('cDayNet',    dayLabels, {day_net_j},  '순매출금액');
makeBar('cDaySettle', dayLabels, {day_set_j}, '정산예정금액');

// 탭 전환
let activeTab = '{("shop" if shop_data else "day")}';
function switchTab(tab){{
  ['shop','day'].forEach(t=>{{
    const panel = document.getElementById('panel-'+t);
    const btn   = document.getElementById('btn-'+t);
    if(!panel||!btn) return;
    if(t===tab){{panel.style.display='block';btn.classList.add('active');}}
    else        {{panel.style.display='none'; btn.classList.remove('active');}}
  }});
  activeTab = tab;
}}
</script>
</body>
</html>
"""

# ── index.html 저장 (항상 덮어쓰기) ───────────────────────
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\n✅ 대시보드 생성 완료!")
print(f"   저장 위치: {OUTPUT_FILE}")
print(f"   (GitHub에 output/index.html 을 push하면 자동 반영됩니다)")
input("\nEnter를 눌러 닫기...")
