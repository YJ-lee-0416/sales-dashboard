# -*- coding: utf-8 -*-
"""
매출 대시보드 v2 - run.py v6 (사방넷 엑셀 → output/index.html)
사용: python run.py   (기본 경로: ./input → ./output/index.html, 템플릿 ./template_v2.html)

역할
  1) 사방넷 '상품별' 다운로드 파일 파싱 + 모델명 매핑 (1순위 품번코드 / 2순위 키워드 / 3순위 원문)
  2) 파싱 결과를 원본 '합 계' 행과 대조 검증 (불일치 시 즉시 중단)
  3) DASHBOARD_DATA(JSON) 생성 → template_v2.html 의 /*__DATA__*/ 자리에 주입 → index.html 출력

주의
  - 일별 파일 파서(parse_day_file)는 로드맵 3-2 인덱스 기준으로 작성되었으나
    실제 일별 샘플 파일이 미제공되어 검증되지 않았습니다.
  - --demo 옵션 시 일자별 데이터는 렌더링 확인용 '가상 데이터'로 채워집니다(meta.demo=True → 화면 경고 배너).
"""
import os, re, sys, json, glob, random, argparse
from datetime import datetime, date, timedelta
from collections import OrderedDict
import warnings
import pandas as pd
warnings.filterwarnings("ignore", message="Workbook contains no default style")

CHANNELS = ["카카오톡스토어", "카카오선물하기", "롯데", "CJ", "오늘의집", "SK", "W컨셉", "CJ온스타일"]
UNASSIGNED = "미지정"

# ── 4-2. 모델명 매핑 (품번코드 → 축약명) ───────────────────────────────
MODEL_MAP = {
    "100071": "CX PRO N_혼합", "100085": "CX PRO N_물걸레", "100084": "CX PRO N_단품",
    "100047": "CX PRO N_기본세트", "100072": "CV6+", "100087": "CV6+ADD", "100018": "CV6+",
    "100070": "에이센스", "100059": "에이센스", "100073": "MT7", "100088": "미니클린",
    "100076": "THC1000", "100032": "THC1000", "100044": "CM6+ADD", "100079": "펫드라이룸",
    "100045": "아쿠아샷", "100080": "HC501", "100089": "HC601", "100074": "고데기",
    "100075": "고데기_블랙", "100066": "CXPRON 배터리", "100067": "CXPRO 먼지봉투",
    "100092": "SC360",
    # ── 구 품번·채널 전용 품번 (상품리스트 260923 미등록, 상품명 기준 수동 지정) ──
    "100083": "에이센스",          # [십일절] 에이센스 BLDC 자동충전 거치대 무선청소기
    "100064": "CX PRO N_물걸레",   # CX PRO 매직타워 N +물걸레키트 세트  ※ 혼합/물걸레 구분 확인 필요
    "100090": "CM6+ADD",           # 무선 욕실청소기 CM6 PLUS ADDITION
    "100055": "MT7",               # 미니 핸디형 무선 청소기 미니멀 투인원 MT7
    "100029": "CV6+ADD",           # CV6 PLUS ADDITION 무선청소기 풀패키지
    "100051": "미니클린",          # 소형 무선청소기 Mini Clean(미니클린) / 블랙
    "100086": "미니클린",          # [십일절] MT8 미니클린 핸디형 미니 무선 청소기
    "100069": "HC501",             # 슈퍼에어릭HC501 BLDC 헤어 드라이어
    "100021": "고데기",            # SECRET 01 대용량 배터리 무선고데기 화이트
    "100016": "CV6+",              # [GS특가] CV6 PLUS 무선청소기 + 2년무상AS
    "100058": "HC501",             # 슈퍼 에어릭 헤어 드라이어 (2608 파일 상품명에 HC501 표기)
    "100068": "펫드라이룸",        # 펫드라이룸 반려동물 털 건조기
    "100077": "CM6+ADD",           # 무선 욕실청소기 CM6 PLUS ADDITION 길이 각도 조절
    "100027": "고데기_블랙",       # 시크릿 미드나잇 블랙 에디션 무선 고데기
    "100078": "아쿠아샷",          # 휴대용 무선 아쿠아샷 구강세정기 TC7
    # 단종·비주력 모델 → '기타'로 합산 (사용자 확정 2026-10-02)
    "100030": "기타",              # NF8 2in1 흡입 물걸레 무선 청소기
    "100081": "기타",              # 2in1 물걸레+진공 무선청소기 NF8 자동충전 거치대
    "100015": "기타",              # 3in1 물걸레 로봇 청소기 CR3
    "100061": "기타",              # 무선청소기 HC02
}
# 2순위: 품번코드 누락 시에만 적용. 구체적 키워드를 먼저 배치(예: '배터리'가 'CX PRO'보다 우선).
# 본체 CX PRO N은 구성(혼합/물걸레/단품) 판별이 불가하므로 키워드 매핑 대상에서 제외 → 원문 표시.
KEYWORD_MAP = [
    (r"먼지봉투", "CXPRO 먼지봉투"),
    (r"미드나잇\s*블랙", "고데기_블랙"),
    (r"고데기|SECRET\s*01", "고데기"),
    (r"CX\s*PRO.*배터리|전용\s*배터리", "CXPRON 배터리"),
    (r"CX\s*PRO.*물걸레", "CX PRO N_물걸레"),    # 품번 누락 '[베스트셀러] CX PRO N +물걸레키트 세트' (사용자 확정)
    (r"스티미|SC\s*360", "SC360"),
    (r"CV6.*ADDITION|CV6\+?ADD", "CV6+ADD"),
    (r"CV6", "CV6+"),
    (r"CM6", "CM6+ADD"),
    (r"에이센스|A\s*sense", "에이센스"),
    (r"MT7|미니멀\s*투인원", "MT7"),
    (r"미니클린|Mini\s*Clean|MT8", "미니클린"),
    (r"THC[-\s]?1000", "THC1000"),
    (r"펫드라이룸", "펫드라이룸"),
    (r"아쿠아샷|AQUA\s*SHOT", "아쿠아샷"),
    (r"HC\s*501", "HC501"),
    (r"HC\s*601", "HC601"),
]


def to_int(v):
    """숫자/콤마 문자열('3,140,999')/NaN/퍼센트 모두 안전 처리."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).replace(",", "").replace("₩", "").strip()
    if s in ("", "-", "nan"):
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


def norm_code(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def map_model(code, name):
    if code and code in MODEL_MAP:
        return MODEL_MAP[code], "code"
    if not code:
        for pat, model in KEYWORD_MAP:
            if re.search(pat, name, re.I):
                return model, "keyword"
    return name, "raw"


def is_total_row(row):
    return re.sub(r"\s", "", str(row.iloc[0])) == "합계"


DATE_TOKEN = r"^\d{4}-\d{4}$|^\d{8}$|^\d{4}$|^\d{6}$"


def extract_channel(filename):
    name = os.path.splitext(os.path.basename(filename))[0]
    name = re.sub(r"(?<!\d)\d{4}-\d{4}(?!\d)", "", name)            # MMDD-MMDD 범위 제거
    parts = [p.strip() for p in re.split(r"[_\-]", name) if p.strip()]
    while parts and re.match(DATE_TOKEN, parts[-1]):
        parts.pop()
    ch = parts[-1] if parts else ""
    # 채널명이 아닌 토큰(예: '다운로드', '다운로드 1')이 남으면 '미지정'
    return UNASSIGNED if (not ch or "다운로드" in ch or "매출현황" in ch) else ch


def extract_file_period(filename):
    """파일명 기간 토큰 → (시작일, 종료일). YYYYMMDD=해당일, YYMM(4자리)=해당 월. 없으면 (None, None)."""
    base = os.path.basename(filename)
    m = re.search(r"(?<!\d)(20\d{6})(?!\d)", base)
    if m:
        d = datetime.strptime(m.group(1), "%Y%m%d").strftime("%Y-%m-%d"); return d, d
    m = re.search(r"(?<!\d)(\d{2})(0[1-9]|1[0-2])(?!\d)", base)
    if m:
        y, mo = 2000 + int(m.group(1)), int(m.group(2))
        last = (date(y + (mo == 12), mo % 12 + 1, 1) - timedelta(days=1)).day
        return f"{y}-{mo:02d}-01", f"{y}-{mo:02d}-{last:02d}"
    return None, None


# ── 상품별 파서 + 합계 대조 ───────────────────────────────────────────
PRODUCT_FIELDS = [("주문수량", 3), ("주문금액", 4), ("취소수량", 5), ("취소금액", 6),
                  ("반품수량", 7), ("반품금액", 8), ("순매출수량", 9), ("순매출금액", 10),
                  ("총이익액", 11), ("판매수수료", 13), ("정산예정금액", 17)]


def parse_product_file(fpath):
    df = pd.read_excel(fpath, header=None)
    rows, total = [], None
    for i, row in df.iterrows():
        if i < 3:
            continue
        if is_total_row(row):
            total = {k: to_int(row.iloc[c]) for k, c in PRODUCT_FIELDS}
            continue
        code = norm_code(row.iloc[1])
        name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not code and not name:
            continue
        model, how = map_model(code, name)
        r = {"품번코드": code, "수집상품명": name, "모델명": model, "매핑": how}
        r.update({k: to_int(row.iloc[c]) for k, c in PRODUCT_FIELDS})
        rows.append(r)
    # 검증: 원본 합계행 vs 파싱 합계
    if total:
        diffs = [(k, total[k], sum(r[k] for r in rows)) for k, _ in PRODUCT_FIELDS
                 if total[k] != sum(r[k] for r in rows)]
        if diffs:
            raise ValueError(f"[검증 실패] {os.path.basename(fpath)} (항목, 원본, 파싱): {diffs}")
    return rows, total


# ── 일별 파서 (로드맵 3-2 인덱스, 샘플 미제공으로 미검증) ─────────────────
DAY_FIELDS = [("주문수량", 3), ("주문금액", 4), ("취소수량", 5), ("취소금액", 6), ("반품수량", 7),
              ("반품금액", 8), ("순매출수량", 9), ("순매출금액", 10), ("총이익액", 11),
              ("판매수수료", 13), ("정산예정금액", 17)]


def parse_day_file(fpath):
    df = pd.read_excel(fpath, header=None)
    out, total = [], None
    for i, row in df.iterrows():
        if i < 3:
            continue
        if is_total_row(row):
            total = {k: to_int(row.iloc[c]) for k, c in DAY_FIELDS}
            continue
        if pd.isna(row.iloc[1]):
            continue
        d = pd.to_datetime(row.iloc[1], errors="coerce")
        if pd.isna(d):
            continue
        r = {"일자": d.strftime("%Y-%m-%d"), "요일": str(row.iloc[2]).strip()}
        r.update({k: to_int(row.iloc[c]) for k, c in DAY_FIELDS})
        out.append(r)
    if total:
        diffs = [(k, total[k], sum(r[k] for r in out)) for k, _ in DAY_FIELDS if total[k] != sum(r[k] for r in out)]
        if diffs:
            raise ValueError(f"[검증 실패] {os.path.basename(fpath)} (항목, 원본, 파싱): {diffs}")
    return out


def merge_day_rows(rows):
    """B방식 누적: 동일 채널 복수 파일 → 일자 기준 병합. 같은 일자가 중복되면 최신 파일 값으로 대체(이중 합산 방지)."""
    by_date = OrderedDict()
    for r in rows:
        by_date[r["일자"]] = r
    return [by_date[k] for k in sorted(by_date)]


# ── 렌더링 확인용 가상 일자별 데이터 ─────────────────────────────────
def demo_daily(base, days=400, seed=7):
    rnd = random.Random(seed)
    scale = {"카카오톡스토어": 9_000_000, "카카오선물하기": 2_500_000, "롯데": 3_000_000, "CJ": 4_000_000,
             "오늘의집": 5_000_000, "SK": 1_500_000, "W컨셉": 900_000, "CJ온스타일": 3_500_000}
    fee = {"카카오톡스토어": .12, "카카오선물하기": .15, "롯데": .25, "CJ": .27, "오늘의집": .18,
           "SK": .20, "W컨셉": .30, "CJ온스타일": .28}
    wd = "월화수목금토일"
    out = {}
    for ch in CHANNELS:
        rows = []
        for k in range(days):
            d = base - timedelta(days=days - 1 - k)
            amt = int(scale[ch] * rnd.uniform(.4, 1.6) * (1.25 if d.weekday() >= 5 else 1) / 100) * 100
            qty = max(1, amt // 180_000)
            cancel = int(amt * rnd.choice([0, 0, 0, .03, .06]) / 100) * 100
            ret = int(amt * rnd.choice([0, 0, 0, 0, .02]) / 100) * 100
            net = amt - cancel - ret
            f = int(net * fee[ch])
            rows.append({"일자": d.strftime("%Y-%m-%d"), "요일": wd[d.weekday()], "주문수량": qty,
                         "주문금액": amt, "취소수량": 1 if cancel else 0, "취소금액": cancel,
                         "반품수량": 1 if ret else 0, "반품금액": ret, "순매출수량": qty,
                         "순매출금액": net, "총이익액": 0, "판매수수료": f, "정산예정금액": net - f})
        out[ch] = rows
    return out


TOTAL = "전체"


def main():
    ap = argparse.ArgumentParser()
    # 기본 경로는 기존 v5 저장소 구조(run.py 위치 기준 input\ → output\index.html)를 따름
    BASE = os.path.dirname(os.path.abspath(__file__))
    ap.add_argument("--input", default=os.path.join(BASE, "input"), help="사방넷 다운로드 파일 폴더")
    ap.add_argument("--template", default=os.path.join(BASE, "template_v2.html"))
    ap.add_argument("--out", default=os.path.join(BASE, "output", "index.html"))
    ap.add_argument("--demo", action="store_true", help="일별 파일이 없을 때 가상 데이터로 채움")
    a = ap.parse_args()

    files = sorted(glob.glob(os.path.join(a.input, "*.xlsx")))
    day_files = [f for f in files if "일별" in os.path.basename(f)]
    prod_files = [f for f in files if "상품별" in os.path.basename(f)]
    warnings, log = {}, []
    warn = lambda ch, msg: warnings.setdefault(ch, []).append(msg)

    daily, fingerprints = {}, {}
    for f in day_files:
        ch = extract_channel(f)
        rows = parse_day_file(f)
        daily.setdefault(ch, []).extend(rows)
        fp = json.dumps([[r[k] for k in ("일자", "주문금액", "정산예정금액")] for r in rows])
        fingerprints.setdefault(fp, []).append(ch)
        log.append({"file": os.path.basename(f), "channel": ch, "days": len(rows), "합계검증": "일치"})
    daily = {ch: merge_day_rows(r) for ch, r in daily.items()}

    # 데이터 점검: 서로 다른 채널 파일의 내용이 완전히 같으면 경고(다운로드 시 채널 선택 오류 의심)
    for chs in fingerprints.values():
        if len(chs) > 1:
            for ch in chs:
                warn(ch, f"{' · '.join(c for c in chs if c != ch)} 파일과 일자별 수치가 완전히 동일함 (다운로드 채널 확인 필요)")
    # 데이터 점검: 정산예정금액이 전 기간 0이면 경고(수수료율 미설정 의심)
    for ch, rows in daily.items():
        amt = sum(r["주문금액"] for r in rows); st = sum(r["정산예정금액"] for r in rows)
        if amt and st == 0:
            warn(ch, "정산예정금액이 전 기간 0원 (판매수수료 = 주문금액 100%) · 사방넷 수수료 설정 확인 필요")

    products = {}
    for f in prod_files:
        rows, total = parse_product_file(f)
        ch = extract_channel(f)
        ch = TOTAL if ch == UNASSIGNED else ch          # 채널 미표기 상품별 파일 = 전 채널 합산으로 간주
        ps, pe = extract_file_period(f)
        products.setdefault(ch, []).append({"date": ps, "s": ps, "e": pe, "file": os.path.basename(f), "rows": rows})
        log.append({"file": os.path.basename(f), "channel": ch, "period": f"{ps}~{pe}", "rows": len(rows),
                    "합계검증": "일치" if total else "합계행 없음",
                    "키워드매핑": [r["수집상품명"] for r in rows if r["매핑"] == "keyword"],
                    "미매핑": [f'{r["품번코드"]} {r["수집상품명"]}' for r in rows if r["매핑"] == "raw"]})

    all_dates = [r["일자"] for rows in daily.values() for r in rows]
    base = max(all_dates) if all_dates else datetime.today().strftime("%Y-%m-%d")
    demo = False
    if a.demo and not all_dates:
        base = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")
        daily = demo_daily(datetime.strptime(base, "%Y-%m-%d").date())
        demo = True

    # 전체 = 채널 합산 (KPI는 개별 채널만 합산하므로 이중 계산 없음)
    agg = OrderedDict()
    for ch, rows in daily.items():
        for r in rows:
            t = agg.setdefault(r["일자"], {"일자": r["일자"], "요일": r["요일"]})
            for k, _ in DAY_FIELDS:
                t[k] = t.get(k, 0) + r[k]
    total_daily = [agg[k] for k in sorted(agg)]

    # 채널 정렬: 주문금액 큰 순 (데이터 기반). 일별 데이터 없는 상품 전용 채널은 뒤에.
    order = sorted(daily, key=lambda c: -sum(r["주문금액"] for r in daily[c]))
    order += [c for c in products if c not in order and c != TOTAL]
    channels = [TOTAL] + order
    daily_out = {TOTAL: total_daily, **daily}

    # 화면 표시 항목만 출력 (취소·반품은 원본에서 제외된 데이터, 순매출=주문금액이므로 미출력. 검증은 전 항목으로 수행 완료)
    KEEP = ("주문수량", "주문금액", "판매수수료", "정산예정금액")
    daily_out = {ch: [{"일자": r["일자"], "요일": r["요일"], **{k: r[k] for k in KEEP}} for r in rows]
                 for ch, rows in daily_out.items()}
    for snaps in products.values():
        for sn in snaps:
            sn["rows"] = [{**{k: r[k] for k in ("품번코드", "수집상품명", "모델명", "매핑")}, **{k: r[k] for k in KEEP}}
                          for r in sn["rows"]]

    data = {
        "meta": {"baseDate": base, "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
                 "files": [os.path.basename(f) for f in files], "demo": demo, "version": "v2",
                 "total": TOTAL, "warnings": warnings},
        "channels": channels, "daily": daily_out, "products": products,
    }
    with open(a.template, encoding="utf-8") as fp:
        html = fp.read()
    html = html.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False))
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    with open(a.out, "w", encoding="utf-8-sig") as fp:
        fp.write(html)
    print(json.dumps(log, ensure_ascii=False, indent=1))
    print(json.dumps(warnings, ensure_ascii=False, indent=1))
    print(f"[완료] {a.out}  기준일={base}  demo={demo}  채널={channels}")


if __name__ == "__main__":
    main()
