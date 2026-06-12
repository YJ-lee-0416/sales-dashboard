# ============================================================
# 사방넷 매출 자동 가공 스크립트 v5
# ============================================================

import pandas as pd
import os, glob, json, re
from datetime import datetime, timedelta
from collections import defaultdict

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR    = os.path.join(BASE_DIR, "input")
OUTPUT_DIR   = os.path.join(BASE_DIR, "output")
HISTORY_FILE = os.path.join(BASE_DIR, "history.json")
OUTPUT_FILE  = os.path.join(OUTPUT_DIR, "index.html")

os.makedirs(INPUT_DIR,  exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def to_int(val):
    try: return int(str(val).replace(",","").replace(" ","").replace("\xa0",""))
    except: return 0

def fmt(n): return f"{n:,}"

def extract_channel(filename):
    name = os.path.splitext(filename)[0]
    parts = name.split("_")
    return parts[-1].strip() if parts else filename

def extract_date_from_files(files):
    for f in files:
        m = re.search(r"(\d{8})", os.path.basename(f))
        if m:
            d = m.group(1)
            return f"{d[:4]}-{d[4:6]}-{d[6:]}"
    return datetime.today().strftime("%Y-%m-%d")

def compute_trends(rows):
    """채널별 rows에서 일별/주별/월별 추이 데이터 계산"""
    date_agg = defaultdict(lambda: {"순매출금액": 0, "정산예정금액": 0})
    for r in rows:
        date_agg[r["일자"]]["순매출금액"]   += r["순매출금액"]
        date_agg[r["일자"]]["정산예정금액"] += r["정산예정금액"]
    sorted_dates = sorted(date_agg.keys())
    if not sorted_dates:
        return {}, {}, {}

    # 일별 최근 7일
    daily_keys = sorted_dates[-7:]
    daily = {
        "labels": daily_keys,
        "net":    [date_agg[d]["순매출금액"]   for d in daily_keys],
        "settle": [date_agg[d]["정산예정금액"] for d in daily_keys],
    }

    # 주별 최근 30일
    week_agg = defaultdict(lambda: {"순매출금액": 0, "정산예정금액": 0})
    for d in sorted_dates:
        dt = datetime.strptime(d, "%Y-%m-%d")
        week_start = (dt - timedelta(days=dt.weekday())).strftime("%Y-%m-%d")
        week_agg[week_start]["순매출금액"]   += date_agg[d]["순매출금액"]
        week_agg[week_start]["정산예정금액"] += date_agg[d]["정산예정금액"]
    last_date = datetime.strptime(sorted_dates[-1], "%Y-%m-%d")
    cutoff_30 = (last_date - timedelta(days=30)).strftime("%Y-%m-%d")
    weekly_keys = sorted(k for k in week_agg if k >= cutoff_30)
    weekly = {
        "labels": [f"{k[5:7]}/{k[8:]}~" for k in weekly_keys],
        "net":    [week_agg[k]["순매출금액"]   for k in weekly_keys],
        "settle": [week_agg[k]["정산예정금액"] for k in weekly_keys],
    }

    # 월별 최근 12개월
    month_agg = defaultdict(lambda: {"순매출금액": 0, "정산예정금액": 0})
    for d in sorted_dates:
        ym = d[:7]
        month_agg[ym]["순매출금액"]   += date_agg[d]["순매출금액"]
        month_agg[ym]["정산예정금액"] += date_agg[d]["정산예정금액"]
    all_months = sorted(month_agg.keys())[-12:]
    monthly = {
        "labels": [f"{m[2:4]}.{m[5:7]}" for m in all_months],
        "net":    [month_agg[m]["순매출금액"]   for m in all_months],
        "settle": [month_agg[m]["정산예정금액"] for m in all_months],
    }

    return daily, weekly, monthly

# ── 파일 탐색 ──────────────────────────────────────────────
all_files = sorted(
    glob.glob(os.path.join(INPUT_DIR, "*.xlsx")) +
    glob.glob(os.path.join(INPUT_DIR, "*.xls")),
    key=os.path.getmtime
)
if not all_files:
    print("❌ input 폴더에 xlsx 또는 xls 파일이 없습니다.")
    input("\nEnter를 눌러 닫기..."); exit()

shop_files = [f for f in all_files if "쇼핑몰" in os.path.basename(f)]
day_files  = [f for f in all_files if "일별"   in os.path.basename(f)]
if not shop_files and not day_files:
    day_files = all_files

print("=" * 54)
print("  사방넷 매출 대시보드 생성기 v5")
print("=" * 54)
report_date = extract_date_from_files(all_files)
print(f"  집계 기준일: {report_date}\n")

# ── 채널별 일별 파싱 ───────────────────────────────────────
channel_day_data = {}
for fpath in day_files:
    channel = extract_channel(os.path.basename(fpath))
    df = pd.read_excel(fpath, header=None)
    rows = []
    for i, row in df.iterrows():
        if i < 3: continue
        val = str(row.iloc[0]).strip()
        if val in ["합 계","합계","NaN","nan",""]: continue
        if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() != "":
            rows.append({
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
    rows.sort(key=lambda x: x["일자"])
    channel_day_data[channel] = rows
    print(f"  ✅ {channel}: {len(rows)}행 파싱")

# ── 쇼핑몰별 파싱 ──────────────────────────────────────────
shop_data = []
if shop_files:
    fpath = sorted(shop_files, key=os.path.getmtime, reverse=True)[0]
    df = pd.read_excel(fpath, header=None)
    for i, row in df.iterrows():
        if i < 3: continue
        val = str(row.iloc[0]).strip()
        if val in ["합 계","합계","NaN","nan",""]: continue
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

# ── KPI 집계 ───────────────────────────────────────────────
src = shop_data if shop_data else [r for rows in channel_day_data.values() for r in rows]
total_order_qty  = sum(r["주문수량"]     for r in src)
total_order_amt  = sum(r["주문금액"]     for r in src)
total_cancel_amt = sum(r["취소금액"]     for r in src)
total_return_amt = sum(r["반품금액"]     for r in src)
total_net_amt    = sum(r["순매출금액"]   for r in src)
total_settle_amt = sum(r["정산예정금액"] for r in src)
settle_ratio     = f"{total_settle_amt/total_net_amt*100:.1f}%" if total_net_amt else "-"

# ── HTML 생성 헬퍼 ─────────────────────────────────────────
def jd(obj): return json.dumps(obj, ensure_ascii=False)

def make_day_table_html(rows, safe_ch):
    """날짜 필터 포함 테이블 HTML"""
    if not rows: return ""
    min_date = rows[0]["일자"]
    max_date = rows[-1]["일자"]
    tbl = ""
    for r in rows:
        net = r["순매출금액"]; settle = r["정산예정금액"]
        ratio = f"{settle/net*100:.1f}%" if net else "-"
        tbl += f"""<tr class="data-row" data-date="{r['일자']}">
          <td>{r['일자']}</td><td class="center">{r['요일']}</td>
          <td class="num">{fmt(r['주문수량'])}</td><td class="num">{fmt(r['주문금액'])}</td>
          <td class="num red">{fmt(r['취소금액'])}</td><td class="num red">{fmt(r['반품금액'])}</td>
          <td class="num bold">{fmt(r['순매출금액'])}</td><td class="num">{fmt(r['판매수수료'])}</td>
          <td class="num purple">{fmt(r['정산예정금액'])}</td><td class="num">{ratio}</td>
        </tr>"""
    return f"""
    <div class="date-filter">
      <label>조회 기간</label>
      <input type="date" id="from_{safe_ch}" value="{min_date}" min="{min_date}" max="{max_date}">
      <span>~</span>
      <input type="date" id="to_{safe_ch}"   value="{max_date}" min="{min_date}" max="{max_date}">
      <button onclick="filterTable('{safe_ch}')">조회</button>
      <button class="reset-btn" onclick="resetFilter('{safe_ch}','{min_date}','{max_date}')">초기화</button>
      <span class="filter-result" id="result_{safe_ch}"></span>
    </div>
    <table id="tbl_{safe_ch}">
      <thead><tr>
        <th>일자</th><th class="center">요일</th><th class="num">주문수량</th><th class="num">주문금액</th>
        <th class="num">취소금액</th><th class="num">반품금액</th><th class="num">순매출금액</th>
        <th class="num">판매수수료</th><th class="num">정산예정금액</th><th class="num">정산율</th>
      </tr></thead>
      <tbody>{tbl}
        <tr class="total-row" id="total_{safe_ch}">
          <td colspan="2">합 계</td>
          <td class="num" id="tot_qty_{safe_ch}"></td><td class="num" id="tot_amt_{safe_ch}"></td>
          <td class="num red" id="tot_can_{safe_ch}"></td><td class="num red" id="tot_ret_{safe_ch}"></td>
          <td class="num bold" id="tot_net_{safe_ch}"></td><td class="num" id="tot_fee_{safe_ch}"></td>
          <td class="num purple" id="tot_set_{safe_ch}"></td><td class="num" id="tot_rat_{safe_ch}"></td>
        </tr>
      </tbody>
    </table>"""

# ── 채널별 패널 + 차트 JS 생성 ────────────────────────────
channel_order = list(channel_day_data.keys())
all_tab_ids   = (["shop"] if shop_data else []) + [ch.replace(" ","_") for ch in channel_order]
first_tab     = "shop" if shop_data else (channel_order[0].replace(" ","_") if channel_order else "")

ch_panels_html = ""
ch_chart_js    = ""

for i, (ch, rows) in enumerate(channel_day_data.items()):
    safe   = ch.replace(" ", "_")
    active = "active" if (not shop_data and i == 0) else ""

    # ── 채널별 추이 계산 ──────────────────────────────────
    daily, weekly, monthly = compute_trends(rows)

    # ── 정산예정금액 막대 (원본 날짜 기준) ───────────────
    settle_lbl = jd([r["일자"]         for r in rows])
    settle_dat = jd([r["정산예정금액"] for r in rows])

    ch_chart_js += f"""
  // ── {ch} ──
  makeLine('cDaily_{safe}',   {jd(daily.get('labels',[]))},   {jd(daily.get('net',[]))},   '순매출금액');
  makeLine('cWeekly_{safe}',  {jd(weekly.get('labels',[]))},  {jd(weekly.get('net',[]))},  '순매출금액');
  makeBar2('cMonthly_{safe}', {jd(monthly.get('labels',[]))}, {jd(monthly.get('net',[]))}, '순매출금액');
  makeBar2('cSettle_{safe}',  {settle_lbl}, {settle_dat}, '정산예정금액');
  calcTotal('{safe}');
"""

    tbl_html = make_day_table_html(rows, safe)

    ch_panels_html += f"""
<div class="panel {active}" id="panel-{safe}">
  <div class="trend-tabs">
    <button class="trend-btn active" onclick="switchTrend('{safe}','daily')">일별 추이 (7일)</button>
    <button class="trend-btn" onclick="switchTrend('{safe}','weekly')">주별 추이 (30일)</button>
    <button class="trend-btn" onclick="switchTrend('{safe}','monthly')">월별 추이 (12개월)</button>
  </div>
  <div class="chart-grid">
    <div id="trend-daily-{safe}">
      <div class="card"><div class="card-title">{ch} · 일별 매출 추이 (최근 7일)</div><div class="chart-wrap"><canvas id="cDaily_{safe}"></canvas></div></div>
    </div>
    <div id="trend-weekly-{safe}" style="display:none">
      <div class="card"><div class="card-title">{ch} · 주별 매출 추이 (최근 30일)</div><div class="chart-wrap"><canvas id="cWeekly_{safe}"></canvas></div></div>
    </div>
    <div id="trend-monthly-{safe}" style="display:none">
      <div class="card"><div class="card-title">{ch} · 월별 매출 추이 (최근 12개월)</div><div class="chart-wrap"><canvas id="cMonthly_{safe}"></canvas></div></div>
    </div>
    <div class="card"><div class="card-title">{ch} · 정산예정금액</div><div class="chart-wrap"><canvas id="cSettle_{safe}"></canvas></div></div>
  </div>
  <div class="tbl-card">
    <div class="card-title">{ch} · 일별 상세 내역</div>
    {tbl_html}
  </div>
</div>"""

# ── 쇼핑몰별 패널 ──────────────────────────────────────────
shop_panel_html = ""
shop_chart_js   = ""
if shop_data:
    s_lbl = jd([r["쇼핑몰명"]     for r in shop_data])
    s_net = jd([r["순매출금액"]   for r in shop_data])
    s_set = jd([r["정산예정금액"] for r in shop_data])
    s_tbl = ""
    for r in shop_data:
        net = r["순매출금액"]; settle = r["정산예정금액"]
        ratio = f"{settle/net*100:.1f}%" if net else "-"
        s_tbl += f"""<tr>
          <td>{r['쇼핑몰명']}</td>
          <td class="num">{fmt(r['주문수량'])}</td><td class="num">{fmt(r['주문금액'])}</td>
          <td class="num red">{fmt(r['취소금액'])}</td><td class="num red">{fmt(r['반품금액'])}</td>
          <td class="num bold">{fmt(r['순매출금액'])}</td><td class="num">{fmt(r['판매수수료'])}</td>
          <td class="num purple">{fmt(r['정산예정금액'])}</td><td class="num">{ratio}</td>
        </tr>"""
    sn = sum(r["순매출금액"] for r in shop_data)
    ss = sum(r["정산예정금액"] for r in shop_data)
    s_tbl += f"""<tr class="total-row">
      <td>합 계</td>
      <td class="num">{fmt(sum(r['주문수량'] for r in shop_data))}</td>
      <td class="num">{fmt(sum(r['주문금액'] for r in shop_data))}</td>
      <td class="num red">{fmt(sum(r['취소금액'] for r in shop_data))}</td>
      <td class="num red">{fmt(sum(r['반품금액'] for r in shop_data))}</td>
      <td class="num bold">{fmt(sn)}</td><td class="num">-</td>
      <td class="num purple">{fmt(ss)}</td>
      <td class="num">{f"{ss/sn*100:.1f}%" if sn else "-"}</td>
    </tr>"""
    shop_panel_html = f"""
<div class="panel active" id="panel-shop">
  <div class="chart-grid">
    <div class="card"><div class="card-title">쇼핑몰별 순매출금액</div><div class="chart-wrap"><canvas id="cShopNet"></canvas></div></div>
    <div class="card"><div class="card-title">쇼핑몰별 정산예정금액</div><div class="chart-wrap"><canvas id="cShopSettle"></canvas></div></div>
  </div>
  <div class="tbl-card"><div class="card-title">쇼핑몰별 상세 내역</div>
    <table><thead><tr>
      <th>쇼핑몰</th><th class="num">주문수량</th><th class="num">주문금액</th>
      <th class="num">취소금액</th><th class="num">반품금액</th><th class="num">순매출금액</th>
      <th class="num">판매수수료</th><th class="num">정산예정금액</th><th class="num">정산율</th>
    </tr></thead><tbody>{s_tbl}</tbody></table>
  </div>
</div>"""
    shop_chart_js = f"""
  makeBar2('cShopNet',    {s_lbl}, {s_net}, '순매출금액');
  makeBar2('cShopSettle', {s_lbl}, {s_set}, '정산예정금액');"""

# 탭 버튼
tab_btns = ""
if shop_data:
    tab_btns += '<button class="tab-btn active" onclick="switchTab(\'shop\')" id="btn-shop">쇼핑몰별</button>\n'
for i, ch in enumerate(channel_order):
    safe   = ch.replace(" ","_")
    active = "active" if (not shop_data and i == 0) else ""
    tab_btns += f'  <button class="tab-btn {active}" onclick="switchTab(\'{safe}\')" id="btn-{safe}">{ch}</button>\n'

filenames_used = ", ".join(os.path.basename(f) for f in day_files[:3])
if len(day_files) > 3: filenames_used += f" 외 {len(day_files)-3}개"

# ── HTML ───────────────────────────────────────────────────
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
:root{{--bg:#F5F6FA;--surface:#fff;--border:#E4E6EF;--t1:#1A1D2E;--t2:#6B7280;
  --blue:#3B82F6;--green:#10B981;--red:#EF4444;--purple:#8B5CF6;--r:12px}}
body{{font-family:'Pretendard',-apple-system,sans-serif;background:var(--bg);color:var(--t1);padding:28px 24px 60px}}
.header{{display:flex;align-items:baseline;gap:10px;margin-bottom:24px;flex-wrap:wrap}}
.header h1{{font-size:21px;font-weight:700;letter-spacing:-.4px}}
.badge{{font-size:12px;font-weight:500;color:var(--t2);background:var(--border);padding:3px 10px;border-radius:20px}}
.fname{{font-size:11px;color:var(--t2);margin-left:auto}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:22px}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:16px 18px}}
.kpi .lbl{{font-size:10px;font-weight:600;color:var(--t2);text-transform:uppercase;letter-spacing:.6px;margin-bottom:7px}}
.kpi .val{{font-size:19px;font-weight:700;letter-spacing:-.5px;line-height:1}}
.kpi .sub{{font-size:11px;color:var(--t2);margin-top:4px}}
.kpi.k-blue{{border-top:3px solid var(--blue)}}.kpi.k-red{{border-top:3px solid var(--red)}}
.kpi.k-green{{border-top:3px solid var(--green)}}.kpi.k-purple{{border-top:3px solid var(--purple)}}
.tabs{{display:flex;gap:6px;margin-bottom:16px;flex-wrap:wrap}}
.tab-btn{{padding:7px 16px;font-size:13px;font-weight:600;border:1px solid var(--border);border-radius:8px;background:var(--surface);color:var(--t2);cursor:pointer;transition:all .15s}}
.tab-btn.active{{background:var(--t1);color:#fff;border-color:var(--t1)}}
.trend-tabs{{display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap}}
.trend-btn{{padding:5px 14px;font-size:12px;font-weight:600;border:1px solid var(--border);border-radius:6px;background:var(--surface);color:var(--t2);cursor:pointer;transition:all .15s}}
.trend-btn.active{{background:var(--blue);color:#fff;border-color:var(--blue)}}
.chart-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
@media(max-width:800px){{.chart-grid{{grid-template-columns:1fr}}}}
.card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px}}
.card-title{{font-size:12px;font-weight:600;color:var(--t2);margin-bottom:14px}}
.chart-wrap{{position:relative;height:240px}}
.tbl-card{{background:var(--surface);border:1px solid var(--border);border-radius:var(--r);padding:18px 20px;overflow-x:auto;margin-bottom:14px}}
/* 날짜 필터 */
.date-filter{{display:flex;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap}}
.date-filter label{{font-size:12px;font-weight:600;color:var(--t2)}}
.date-filter input[type=date]{{padding:5px 10px;font-size:12px;border:1px solid var(--border);border-radius:6px;font-family:inherit}}
.date-filter button{{padding:5px 14px;font-size:12px;font-weight:600;background:var(--blue);color:#fff;border:none;border-radius:6px;cursor:pointer}}
.date-filter .reset-btn{{background:var(--surface);color:var(--t2);border:1px solid var(--border)}}
.filter-result{{font-size:11px;color:var(--t2);margin-left:4px}}
table{{width:100%;border-collapse:collapse;font-size:12.5px}}
thead th{{font-size:10px;font-weight:600;color:var(--t2);text-transform:uppercase;letter-spacing:.5px;padding:0 8px 9px;border-bottom:2px solid var(--border);white-space:nowrap;text-align:left}}
thead th.num{{text-align:right}}
tbody tr{{border-bottom:1px solid var(--border);transition:background .12s}}
tbody tr:last-child{{border-bottom:none}}
tbody tr:hover{{background:#F8F9FF}}
tbody td{{padding:10px 8px;vertical-align:middle}}
td.num{{text-align:right;font-variant-numeric:tabular-nums}}
td.center{{text-align:center}}td.red{{color:var(--red)}}
td.purple{{color:var(--purple);font-weight:600}}td.bold{{font-weight:600}}
tr.total-row{{font-weight:700;border-top:2px solid var(--border)!important}}
tr.hidden{{display:none}}
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
<div class="kpi-grid">
  <div class="kpi k-blue"><div class="lbl">주문금액</div><div class="val">₩{fmt(total_order_amt)}</div><div class="sub">{fmt(total_order_qty)}건</div></div>
  <div class="kpi k-red"><div class="lbl">취소·반품</div><div class="val">₩{fmt(total_cancel_amt+total_return_amt)}</div><div class="sub">취소 {fmt(total_cancel_amt)} / 반품 {fmt(total_return_amt)}</div></div>
  <div class="kpi k-green"><div class="lbl">순매출금액</div><div class="val">₩{fmt(total_net_amt)}</div><div class="sub">&nbsp;</div></div>
  <div class="kpi k-purple"><div class="lbl">정산예정금액</div><div class="val">₩{fmt(total_settle_amt)}</div><div class="sub">순매출 대비 {settle_ratio}</div></div>
</div>
<div class="tabs">{tab_btns}</div>
{shop_panel_html}
{ch_panels_html}
<div class="footer">생성: {datetime.now().strftime("%Y-%m-%d %H:%M")} · {filenames_used}</div>

<script>
const COLORS=['#3B82F6','#10B981','#F59E0B','#EF4444','#8B5CF6','#06B6D4','#EC4899','#FEE500'];

function makeLine(id, labels, data, label){{
  const el=document.getElementById(id); if(!el)return;
  new Chart(el.getContext('2d'),{{type:'line',
    data:{{labels,datasets:[{{label,data,borderColor:'#3B82F6',
      backgroundColor:'rgba(59,130,246,0.08)',pointBackgroundColor:'#3B82F6',
      pointRadius:4,pointHoverRadius:6,tension:0.3,fill:true,borderWidth:2}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>'₩'+c.parsed.y.toLocaleString()}}}}}},
      scales:{{
        y:{{ticks:{{callback:v=>v>=1e8?(v/1e8).toFixed(1)+'억':v>=1e4?(v/1e4).toFixed(0)+'만':v,font:{{size:10}}}},grid:{{color:'#E4E6EF'}}}},
        x:{{ticks:{{font:{{size:10}}}},grid:{{display:false}}}}
      }}
    }}
  }});
}}

function makeBar2(id, labels, data, label){{
  const el=document.getElementById(id); if(!el)return;
  new Chart(el.getContext('2d'),{{type:'bar',
    data:{{labels,datasets:[{{label,data,
      backgroundColor:labels.map((_,i)=>COLORS[i%COLORS.length]+'BB'),
      borderColor:labels.map((_,i)=>COLORS[i%COLORS.length]),
      borderWidth:1.5,borderRadius:5}}]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>'₩'+c.parsed.y.toLocaleString()}}}}}},
      scales:{{
        y:{{ticks:{{callback:v=>v>=1e8?(v/1e8).toFixed(1)+'억':v>=1e4?(v/1e4).toFixed(0)+'만':v,font:{{size:10}}}},grid:{{color:'#E4E6EF'}}}},
        x:{{ticks:{{font:{{size:10}}}},grid:{{display:false}}}}
      }}
    }}
  }});
}}

// 탭 전환
const ALL_TABS = {jd(all_tab_ids)};
function switchTab(tab){{
  ALL_TABS.forEach(t=>{{
    const p=document.getElementById('panel-'+t);
    const b=document.getElementById('btn-'+t);
    if(p) p.style.display=t===tab?'block':'none';
    if(b) b.classList.toggle('active',t===tab);
  }});
}}

// 추이 탭 전환
function switchTrend(ch, type){{
  ['daily','weekly','monthly'].forEach(t=>{{
    const el=document.getElementById('trend-'+t+'-'+ch);
    if(el) el.style.display=t===type?'':'none';
  }});
  const panel=document.getElementById('panel-'+ch);
  if(panel) panel.querySelectorAll('.trend-btn').forEach((btn,i)=>{{
    btn.classList.toggle('active',['daily','weekly','monthly'][i]===type);
  }});
}}

// 합계 행 계산
function calcTotal(ch){{
  const rows=[...document.querySelectorAll('#tbl_'+ch+' .data-row:not(.hidden)')];
  const cols=['qty','amt','can','ret','net','fee','set'];
  const idx =[2,3,4,5,6,7,8];
  const sums=cols.map(()=>0);
  rows.forEach(tr=>{{
    cols.forEach((_,i)=>{{
      const txt=tr.cells[idx[i]].textContent.replace(/[,₩]/g,'').trim();
      sums[i]+=parseInt(txt)||0;
    }});
  }});
  cols.forEach((c,i)=>{{
    const el=document.getElementById('tot_'+c+'_'+ch);
    if(el) el.textContent=sums[i].toLocaleString();
  }});
  const ratEl=document.getElementById('tot_rat_'+ch);
  if(ratEl) ratEl.textContent=sums[4]>0?(sums[6]/sums[4]*100).toFixed(1)+'%':'-';
  const resEl=document.getElementById('result_'+ch);
  if(resEl) resEl.textContent=rows.length+'일 조회 중';
}}

// 날짜 필터
function filterTable(ch){{
  const from=document.getElementById('from_'+ch).value;
  const to  =document.getElementById('to_'+ch).value;
  document.querySelectorAll('#tbl_'+ch+' .data-row').forEach(tr=>{{
    const d=tr.dataset.date;
    tr.classList.toggle('hidden', d<from||d>to);
  }});
  calcTotal(ch);
}}

function resetFilter(ch, minD, maxD){{
  document.getElementById('from_'+ch).value=minD;
  document.getElementById('to_'+ch).value=maxD;
  document.querySelectorAll('#tbl_'+ch+' .data-row').forEach(tr=>tr.classList.remove('hidden'));
  calcTotal(ch);
}}

window.addEventListener('DOMContentLoaded', function(){{
{shop_chart_js}
{ch_chart_js}
}});
</script>
</body>
</html>
"""

with open(OUTPUT_FILE, "w", encoding="utf-8-sig") as f:
    f.write(html)

history = {}
if os.path.exists(HISTORY_FILE):
    with open(HISTORY_FILE, "r", encoding="utf-8") as f:
        history = json.load(f)
history[report_date] = {ch: rows for ch, rows in channel_day_data.items()}
with open(HISTORY_FILE, "w", encoding="utf-8") as f:
    json.dump(history, f, ensure_ascii=False, indent=2)

print(f"\n✅ 대시보드 생성 완료!")
print(f"   처리 채널: {', '.join(channel_order)}")
print(f"   저장 위치: {OUTPUT_FILE}")
input("\nEnter를 눌러 닫기...")
