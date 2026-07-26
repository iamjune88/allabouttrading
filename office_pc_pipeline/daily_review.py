# -*- coding: utf-8 -*-
"""
KTB 데일리 리뷰 마스터 스크립트
사용법: python daily_review.py [YYYY/MM/DD] [--ktb3 +240] [--ktb10 -140]
날짜 생략 시 오늘 날짜 자동 사용.
오버나잇 포지션 생략 시 0으로 처리.
"""
import argparse, json, os, re, subprocess, sys, time
from datetime import datetime, timezone, timedelta
from pathlib import Path

import openpyxl
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── 경로 설정 ─────────────────────────────────────────────────────────────
THIS_DIR = Path(__file__).parent
REPO_DIR = THIS_DIR.parent / "allabouttrading"
MCP_DIR  = THIS_DIR.parent / "tradingview-mcp"
RESULT_DIR  = THIS_DIR / "결과"
ASSET_DIR   = THIS_DIR.parent / "journal_assets"
ASSET_DIR.mkdir(exist_ok=True)

KST = timezone(timedelta(hours=9))

# ── 인수 파싱 ─────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser()
parser.add_argument("date", nargs="?", default=None,
                    help="날짜 YYYY/MM/DD (생략 시 오늘)")
parser.add_argument("--ktb3",  type=str, default=None,
                    help="KTB3 오버나잇 포지션 예: +240 (생략 시 Excel 누적합에서 자동 계산)")
parser.add_argument("--ktb10", type=str, default=None,
                    help="KTB10 오버나잇 포지션 예: -140 (생략 시 Excel 누적합에서 자동 계산)")
args = parser.parse_args()

today_kst = datetime.now(tz=KST)
if args.date:
    y, m, d = args.date.split("/")
    target = datetime(int(y), int(m), int(d), tzinfo=KST)
else:
    target = today_kst.replace(hour=0, minute=0, second=0, microsecond=0)

DATE_STR  = target.strftime("%Y/%m/%d")   # "2026/07/01"
DATE_SLUG = target.strftime("%Y%m%d")     # "20260701"
DATE_DISP = target.strftime("%Y-%m-%d (%a)").replace(
    "Mon","월").replace("Tue","화").replace("Wed","수").replace(
    "Thu","목").replace("Fri","금").replace("Sat","토").replace("Sun","일")

# 오버나잇 포지션: 명시적 args 우선, 없으면 Excel 누적합에서 결정론적으로 계산.
import ovn_tracker

_ktb3_explicit  = args.ktb3  is not None
_ktb10_explicit = args.ktb10 is not None

_ovn_info = {}
try:
    _ovn_info = ovn_tracker.compute_ovn(DATE_STR)
    _auto_ktb3, _auto_ktb10 = _ovn_info["ovn_ktb3"], _ovn_info["ovn_ktb10"]
    _auto_desc = f"Excel누적/anchor {_ovn_info['anchor_date']}"
except Exception as e:
    _auto_ktb3, _auto_ktb10 = 0, 0
    _auto_desc = f"계산실패({e})"

OVN_KTB3  = int(args.ktb3)  if _ktb3_explicit  else _auto_ktb3
OVN_KTB10 = int(args.ktb10) if _ktb10_explicit else _auto_ktb10

_src = "입력값" if (_ktb3_explicit or _ktb10_explicit) else _auto_desc

# ── 크로스체크: Excel 누적 OVN ↔ 브로커 확정 미결제약정 ──
import ovn_crosscheck
_snapshot = ovn_crosscheck.load_snapshot(DATE_STR)
_xcheck = None
if _snapshot and not (_ktb3_explicit or _ktb10_explicit):
    _xcheck = ovn_crosscheck.crosscheck(_ovn_info, _snapshot)
    for _w in _xcheck["warnings"]:
        print(f"  [경고] {_w}")
    if not _xcheck["ok"]:
        print(f"\n{'!'*60}")
        print("  [중단] Excel 누적 OVN과 브로커 미결제약정이 불일치합니다.")
        for _m in _xcheck["mismatches"]:
            print(f"    - {_m}")
        print("  Excel 체결 증적 누락/중복 또는 anchor 오류 가능성. 확인 후 재실행하세요.")
        print(f"    (수동 확정 시: --ktb3 {_snapshot['entry']['ktb3']:+d} "
              f"--ktb10 {_snapshot['entry']['ktb10']:+d})")
        print(f"{'!'*60}\n")
        sys.exit(1)
    print(f"  [크로스체크 OK] 브로커 미결제약정과 진입 OVN 일치 "
          f"(brokers: {','.join(_snapshot['brokers'])})")
elif not _snapshot and not (_ktb3_explicit or _ktb10_explicit):
    print("  [크로스체크 생략] 해당 날짜 브로커 스냅샷 없음 (main.py 미실행 또는 SS/NH 미결 부재)")

# 타임스탬프 기준 (1781481600 = 2026-06-15 00:00 UTC = 09:00 KST)
EPOCH_0615 = 1781481600
day0_kst   = datetime(2026, 6, 15, tzinfo=KST)
days_since = (target - day0_kst).days
DAY_BASE      = EPOCH_0615 + days_since * 86400
SESSION_START = DAY_BASE - 900
SESSION_END   = DAY_BASE + 6 * 3600 + 45 * 60

print(f"\n{'='*60}")
print(f"  KTB 데일리 리뷰: {DATE_DISP}")
print(f"  오버나잇: KTB3 {OVN_KTB3:+d}  /  KTB10 {OVN_KTB10:+d}  [{_src}]")
print(f"{'='*60}\n")


# ── 1. Excel 체결내역 읽기 ────────────────────────────────────────────────
def load_fills(date_str):
    y, m, d = date_str.split("/")
    fname = RESULT_DIR / f"선물거래_{y}-{m}.xlsx"
    if not fname.exists():
        sys.exit(f"[오류] Excel 파일 없음: {fname}")
    wb  = openpyxl.load_workbook(fname)
    ws  = wb.active
    rows = list(ws.iter_rows(values_only=True))
    target_date = f"{y}-{m}-{d}"
    fills = []
    for r in rows[1:]:
        if str(r[1]) == target_date or (hasattr(r[1], 'strftime') and r[1].strftime('%Y-%m-%d') == target_date):
            fills.append({
                "source": r[0], "date": r[1], "code": r[3],
                "side": r[4],   "qty": int(r[5]), "price": float(r[6]),
                "time": str(r[7]), "pnl": int(r[9] or 0), "fee": int(r[10] or 0),
            })
    return fills

print("[1/6] Excel 체결내역 로드...")
fills = load_fills(DATE_STR)
if not fills:
    sys.exit(f"[오류] {DATE_STR} 체결내역 없음. main.py를 먼저 실행하세요.")

codes = list({f["code"] for f in fills})
print(f"      {len(fills)}건  종목: {codes}")


# ── 2. 체결 집계 (시각+방향+가격 기준) ────────────────────────────────────
def aggregate_fills(fills):
    from collections import defaultdict
    bucket = defaultdict(lambda: {"qty": 0, "pnl": 0})
    for f in fills:
        hh, mm = int(f["time"][:2]), int(f["time"][3:5])
        key = (f["side"], float(f["price"]), hh, mm)
        bucket[key]["qty"] += f["qty"]
        bucket[key]["pnl"] += f["pnl"]
    result = []
    for (side, price, hh, mm), v in sorted(bucket.items(), key=lambda x: (x[0][2], x[0][3])):
        result.append((side, v["qty"], price, hh, mm, v["pnl"]))
    return result

# 종목코드 분류: A67XX = KTB10, A65XX = KTB3
fills_10 = [f for f in fills if f["code"].startswith("A67")]
fills_3  = [f for f in fills if f["code"].startswith("A65")]

agg    = aggregate_fills(fills)     # 전체 (저널 테이블용)
agg_10 = aggregate_fills(fills_10)  # KTB10 마커용
agg_3  = aggregate_fills(fills_3)   # KTB3 마커용

total_pnl = sum(f["pnl"] for f in fills)
total_fee = sum(f["fee"] for f in fills)

def avg_stats(fill_list):
    s = [(q,p) for s,q,p,*_ in fill_list if s=="매도"]
    b = [(q,p) for s,q,p,*_ in fill_list if s=="매수"]
    sq = sum(q for q,p in s); bq = sum(q for q,p in b)
    sa = sum(q*p for q,p in s)/sq if sq else 0
    ba = sum(q*p for q,p in b)/bq if bq else 0
    return sq, bq, sa, ba

sell_q10,buy_q10,sell_avg10,buy_avg10 = avg_stats(agg_10)
sell_q3, buy_q3, sell_avg3, buy_avg3  = avg_stats(agg_3)
total_sell_q = sell_q10 + sell_q3
total_buy_q  = buy_q10  + buy_q3
sell_avg = (sell_q10*sell_avg10 + sell_q3*sell_avg3) / total_sell_q if total_sell_q else 0
buy_avg  = (buy_q10 *buy_avg10  + buy_q3 *buy_avg3)  / total_buy_q  if total_buy_q  else 0

print(f"      KTB10: SELL {sell_q10} @ {sell_avg10:.3f} / BUY {buy_q10} @ {buy_avg10:.3f}")
print(f"      KTB3:  SELL {sell_q3}  @ {sell_avg3:.3f} / BUY {buy_q3}  @ {buy_avg3:.3f}")
print(f"      손익합계: {total_pnl:+,}원  수수료: {total_fee:,}원")


# ── 3. TV CDP 확인 + OHLCV 취득 ──────────────────────────────────────────
def check_cdp():
    import urllib.request
    try:
        urllib.request.urlopen("http://localhost:9222/json/version", timeout=3)
        return True
    except Exception:
        return False

def start_tv():
    print("      TradingView 실행 중...")
    ps = (
        '$env:ELECTRON_RUN_AS_NODE=$null;'
        '$pkg=Get-AppxPackage|Where-Object{$_.Name -like "*TradingView*"}|Select-Object -First 1;'
        '$exe=Join-Path $pkg.InstallLocation "TradingView.exe";'
        'Start-Process $exe -ArgumentList "--remote-debugging-port=9222"'
    )
    subprocess.run(["powershell", "-Command", ps], capture_output=True)
    for _ in range(15):
        time.sleep(2)
        if check_cdp():
            print("      CDP 연결 확인 [OK]")
            return
    sys.exit("[오류] TradingView CDP 연결 실패")

def get_ohlcv(symbol, out_file, price_near=None, price_tol=1.5, max_retries=4):
    """
    price_near: 전일 정산가 기준 절대값 검증 (브로커 스냅샷 당일정산가에서 로드)
    price_tol:  허용 편차 (기본 +-1.5pt)
    chart_ready 직후에도 bar 데이터가 이전 심볼 것일 수 있으므로 가격으로 재검증.
    """
    env = os.environ.copy()
    env.pop("ELECTRON_RUN_AS_NODE", None)
    is_past = DATE_SLUG < today_kst.strftime("%Y%m%d")
    last_raw, last_data = None, {"bars": []}

    for attempt in range(max_retries):
        if attempt > 0:
            print(f"      [{symbol}] 재시도 {attempt}/{max_retries-1} (6초 대기)...")
            time.sleep(6)

        subprocess.run(
            ["node", "src/cli/index.js", "symbol", symbol],
            cwd=str(MCP_DIR), capture_output=True, env=env
        )
        for _ in range(15):  # 최대 30초
            time.sleep(2)
            r = subprocess.run(
                ["node", "src/cli/index.js", "symbol"],
                cwd=str(MCP_DIR), capture_output=True, env=env
            )
            try:
                st = json.loads(r.stdout.decode("utf-8", errors="replace"))
                if st.get("symbol") == symbol and st.get("chart_ready", False):
                    break
            except Exception:
                pass

        time.sleep(3)  # chart_ready 직후에도 bar 데이터 로드 시간 필요

        # 항상 range 명령 전송 — 이전 실행이 다른 날짜 뷰포트를 고정했을 수 있음
        subprocess.run(
            ["node", "src/cli/index.js", "range",
             "--from", str(SESSION_START), "--to", str(SESSION_END)],
            cwd=str(MCP_DIR), capture_output=True, env=env
        )
        time.sleep(6)

        result = subprocess.run(
            ["node", "src/cli/index.js", "ohlcv", "--count", "500"],
            cwd=str(MCP_DIR), capture_output=True, env=env
        )
        raw = result.stdout
        if isinstance(raw, bytes):
            for enc in ["utf-8-sig", "utf-8", "cp949"]:
                try: raw = raw.decode(enc); break
                except: pass

        try:
            data = json.loads(raw)
        except Exception:
            continue

        bars = data.get("bars", [])
        if not bars:
            continue

        last_raw, last_data = raw, data

        # 검증1: 오늘 세션 봉이 반드시 있어야 함
        session_bars = [b for b in bars if b["time"] >= SESSION_START]
        if not session_bars:
            print(f"      [{symbol}] 오늘 세션 봉 없음 -> 재시도")
            continue

        # 검증2: 전일 정산가 +-price_tol 범위 이내여야 정상 데이터
        if price_near is not None:
            closes = [b["close"] for b in session_bars]
            median_c = sorted(closes)[len(closes) // 2]
            if abs(median_c - price_near) > price_tol:
                print(f"      [{symbol}] 가격 검증 실패: {median_c:.3f} (기준 {price_near:.3f}+-{price_tol}) -> 재시도")
                continue

        with open(out_file, "w", encoding="utf-8") as f:
            f.write(raw if isinstance(raw, str) else raw)
        print(f"      {symbol}: {len(bars)}봉 취득")
        return data

    # 모든 재시도 실패시 마지막 데이터로 진행
    print(f"      [{symbol}] 경고: 가격 검증 실패 지속, 마지막 데이터로 진행")
    if last_raw:
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(last_raw if isinstance(last_raw, str) else last_raw)
    return last_data

print("[2/6] TV CDP 확인 + OHLCV 취득...")
if not check_cdp():
    start_tv()

# 전일 정산가를 기준으로 각 심볼 절대 가격 검증 (BM31/BMA1 혼동 방지)
# 브로커 스냅샷의 당일정산가에서 로드 — 없으면 None (검증 스킵)
def _settle_ref(leg):
    if not _snapshot:
        return None
    s = _snapshot.get("settlement", {}).get(leg, {})
    return s.get("today") or s.get("prev")

_ktb3_ref  = _settle_ref("ktb3")
_ktb10_ref = _settle_ref("ktb10")

ohlcv_bm31 = get_ohlcv("BM31!", THIS_DIR / f"_ohlcv_{DATE_SLUG}_BM31.json",
                         price_near=_ktb3_ref)
ohlcv_bma  = get_ohlcv("BMA1!", THIS_DIR / f"_ohlcv_{DATE_SLUG}_BMA1.json",
                         price_near=_ktb10_ref)


# ── 4. 차트 데이터 처리 ───────────────────────────────────────────────────
def sma(series, w):
    out = [None]*len(series)
    for i in range(len(series)):
        if i+1 >= w:
            out[i] = sum(series[i+1-w:i+1])/w
    return out

def rsi(series, period=14):
    out=[None]*len(series); gains=[]; losses=[]; st={}
    for i in range(1, len(series)):
        d=series[i]-series[i-1]; gains.append(max(d,0)); losses.append(max(-d,0))
        if i>=period:
            if i==period: ag=sum(gains[:period])/period; al=sum(losses[:period])/period
            else: ag=(st["ag"]*(period-1)+gains[-1])/period; al=(st["al"]*(period-1)+losses[-1])/period
            st={"ag":ag,"al":al}; rs=ag/al if al else float("inf")
            out[i]=100-100/(1+rs) if al else 100.0
    return out

def process(ohlcv_data):
    bars = sorted(ohlcv_data.get("bars", []), key=lambda b: b["time"])
    filtered = [b for b in bars if b["time"] < SESSION_END+300]
    bars = filtered if filtered else bars  # TV가 다음날 봉 반환 시 필터 후 빈 경우 원본 사용
    if not bars:  # bars 자체가 완전히 빈 경우 - dummy로 크래시 방지
        dummy = {"time": SESSION_START, "open": 0.0, "high": 0.0, "low": 0.0, "close": 0.0, "volume": 0}
        bars = [dummy]
        print(f"      [경고] 봉 데이터 완전 없음 - dummy 봉으로 대체")
    closes = [b["close"] for b in bars]
    s20=sma(closes,20); s60=sma(closes,60); rs=rsi(closes)
    idx=[i for i,b in enumerate(bars) if b["time"]>=SESSION_START]
    sb=[bars[i] for i in idx]
    if not sb:  # 세션 봉 없으면 마지막 100봉으로 fallback (날짜 오류 상황)
        sb = bars[-min(100, len(bars)):]
        idx = list(range(max(0, len(bars)-min(100, len(bars))), len(bars)))
        print(f"      [경고] 세션 봉 없음 - 최근 {len(sb)}봉으로 대체")
    ts=[datetime.fromtimestamp(b["time"],tz=KST) for b in sb]
    vwap=[]; cpv=0.0; cv=0.0
    for b in sb:
        tp=(b["high"]+b["low"]+b["close"])/3; cpv+=tp*b["volume"]; cv+=b["volume"]
        vwap.append(cpv/cv if cv else None)
    hi=max(b["high"] for b in sb); lo=min(b["low"] for b in sb)
    N=36; ed=[lo+(hi-lo)*k/N for k in range(N+1)]; bv=[0.0]*N
    for b in sb:
        blo,bhi,vol=b["low"],b["high"],b["volume"]; sp=max(bhi-blo,0.001)
        for k in range(N):
            ov=max(0.0,min(bhi,ed[k+1])-max(blo,ed[k]))
            if ov: bv[k]+=vol*ov/sp
    bm=[(ed[k]+ed[k+1])/2 for k in range(N)]
    return sb, ts, [s20[i] for i in idx], [s60[i] for i in idx], [rs[i] for i in idx], vwap, bv, bm, hi, lo

print("[3/6] 차트 데이터 처리...")
(bma_bars,  bma_ts,  bma_s20,  bma_s60,  bma_rsi,  bma_vwap,  bma_bv,  bma_bm,  bma_hi,  bma_lo)  = process(ohlcv_bma)
(bm31_bars, bm31_ts, bm31_s20, bm31_s60, bm31_rsi, bm31_vwap, bm31_bv, bm31_bm, bm31_hi, bm31_lo) = process(ohlcv_bm31)

# 전일 종가: 전일 세션 마감(15:45 KST) 이하 마지막 봉 종가 (갱신차금 계산용)
def prev_session_close(ohlcv_data):
    bars = sorted(ohlcv_data["bars"], key=lambda b: b["time"])
    prev_end = DAY_BASE - 86400 + 24300  # 전일 15:45 KST (야간선물 시작 직전)
    pre = [b for b in bars if b["time"] < prev_end]  # 정규장 마지막 봉(15:40)만 포함, 야간봉 제외
    return pre[-1]["close"] if pre else None

prev_close_ktb3  = prev_session_close(ohlcv_bm31)
prev_close_ktb10 = prev_session_close(ohlcv_bma)
today_close_ktb3  = bm31_bars[-1]["close"] if bm31_bars else None
today_close_ktb10 = bma_bars[-1]["close"]  if bma_bars  else None

MULT = 1_000_000  # KTB3/KTB10: 1pt = 100만원/계약

# 갱신차금(오버나잇 MTM): 브로커 확정값(스냅샷) 우선, 없으면 OHLCV 종가 근사.
_renewal = _xcheck["renewal"] if _xcheck else None
if _renewal and (_renewal.get("ktb3") or _renewal.get("ktb10")):
    ovn_pnl_ktb3  = int(_renewal.get("ktb3", 0))
    ovn_pnl_ktb10 = int(_renewal.get("ktb10", 0))
    _ovn_src = _xcheck["renewal_source"]
    print(f"      KTB3 갱신차금 {ovn_pnl_ktb3:+,}원  KTB10 {ovn_pnl_ktb10:+,}원 [{_ovn_src}]")
else:
    ovn_pnl_ktb3  = int(OVN_KTB3  * (today_close_ktb3  - prev_close_ktb3)  * MULT) if (prev_close_ktb3  and today_close_ktb3  and OVN_KTB3  != 0) else 0
    ovn_pnl_ktb10 = int(OVN_KTB10 * (today_close_ktb10 - prev_close_ktb10) * MULT) if (prev_close_ktb10 and today_close_ktb10 and OVN_KTB10 != 0) else 0
    _ovn_src = "OHLCV근사"
    if prev_close_ktb3:
        print(f"      KTB3 전일종가 {prev_close_ktb3:.2f} -> 당일종가 {today_close_ktb3:.2f}  갱신차금 {ovn_pnl_ktb3:+,}원 [OHLCV근사]")
    if prev_close_ktb10:
        print(f"      KTB10 전일종가 {prev_close_ktb10:.2f} -> 당일종가 {today_close_ktb10:.2f}  갱신차금 {ovn_pnl_ktb10:+,}원 [OHLCV근사]")

total_ovn_pnl = ovn_pnl_ktb3 + ovn_pnl_ktb10
total_pnl_all = total_pnl + total_ovn_pnl
print(f"      총손익: 갱신차금 {total_ovn_pnl:+,}원({_ovn_src}) + 건별 {total_pnl:+,}원 = {total_pnl_all:+,}원")


# ── 5. Plotly 차트 생성 ───────────────────────────────────────────────────
def bar_ts_fn(hh, mm):
    return ((DAY_BASE + (hh-9)*3600 + mm*60) // 300) * 300

BG,GRID,TXT="#0d1117","#21262d","#8b949e"
UP,DOWN="#3fb950","#f85149"
IND={"SMA20":"#e3b341","SMA60":"#39c5cf","VWAP":"#bc8cff","RSI":"#d29922"}

X_LEFT  = datetime(target.year, target.month, target.day, 8, 37, tzinfo=KST)
X_RIGHT = datetime(target.year, target.month, target.day, 15, 55, tzinfo=KST)

fig = make_subplots(
    rows=2, cols=4, row_heights=[0.76,0.24],
    column_widths=[0.435,0.065,0.435,0.065],
    shared_xaxes=False, horizontal_spacing=0.008, vertical_spacing=0.03,
    specs=[[{},{},{},{}],[{},None,{},None]],
)

def add_candle(bars, ts, col, name):
    fig.add_trace(go.Candlestick(
        x=ts, open=[b["open"] for b in bars], high=[b["high"] for b in bars],
        low=[b["low"] for b in bars], close=[b["close"] for b in bars],
        increasing_line_color=UP, decreasing_line_color=DOWN,
        name=name, showlegend=False,
    ), row=1, col=col)

def add_ind(ts, s20, s60, vwap, col, legend):
    for ser, nm, co in [(s20,"SMA20",IND["SMA20"]),(s60,"SMA60",IND["SMA60"]),(vwap,"VWAP",IND["VWAP"])]:
        fig.add_trace(go.Scatter(x=ts, y=ser, mode="lines", name=nm,
                                  line=dict(color=co,width=1.3), showlegend=legend), row=1, col=col)

def add_vrvp(bv, bm, col):
    fig.add_trace(go.Bar(x=bv, y=bm, orientation="h",
                          marker_color="#39c5cf", opacity=0.55, showlegend=False), row=1, col=col)

def add_rsi_panel(ts, rsi_vals, col):
    fig.add_trace(go.Scatter(x=ts, y=rsi_vals, mode="lines",
                              line=dict(color=IND["RSI"],width=1.2), showlegend=False), row=2, col=col)
    for lv, co in [(70,DOWN),(50,TXT),(30,UP)]:
        fig.add_hline(y=lv, line=dict(color=co, width=0.7 if lv!=50 else 0.5,
                                       dash="dot" if lv!=50 else "solid"), row=2, col=col)

# KTB10 (cols 1,2)
add_candle(bma_bars, bma_ts, col=1, name="KTB10")
add_ind(bma_ts, bma_s20, bma_s60, bma_vwap, col=1, legend=True)

def add_markers(agg_list, col, show_legend, sell_avg_val, buy_avg_val):
    for side_kor, color, sym, nm in [
        ("매도","#FF4444","triangle-down","SELL"),
        ("매수","#44BB44","triangle-up",  "BUY"),
    ]:
        tlist = [(s,q,p,hh,mm) for s,q,p,hh,mm,_ in agg_list if s==side_kor]
        if not tlist: continue
        xs = [datetime.fromtimestamp(bar_ts_fn(hh,mm), tz=KST) for _,q,p,hh,mm in tlist]
        ys = [p for _,q,p,hh,mm in tlist]
        hv = [f"{nm} {q}@{p:.2f} {hh:02d}:{mm:02d}" for _,q,p,hh,mm in tlist]
        fig.add_trace(go.Scatter(x=xs, y=ys, mode="markers", name=nm,
            marker=dict(symbol=sym, size=9, color=color, line=dict(width=1,color="#0d1117")),
            text=hv, hovertemplate="%{text}<extra></extra>", showlegend=show_legend), row=1, col=col)
    if sell_avg_val:
        fig.add_hline(y=sell_avg_val, line=dict(color="#FF8888",width=1,dash="dash"), row=1, col=col,
                      annotation_text=f"SELL avg {sell_avg_val:.2f}", annotation_position="right",
                      annotation_font_color="#FF8888")
    if buy_avg_val:
        fig.add_hline(y=buy_avg_val, line=dict(color="#88FF88",width=1,dash="dash"), row=1, col=col,
                      annotation_text=f"BUY avg {buy_avg_val:.2f}", annotation_position="right",
                      annotation_font_color="#88FF88")

# KTB10 마커 (col1), KTB3 마커 (col3)
add_markers(agg_10, col=1, show_legend=True,  sell_avg_val=sell_avg10, buy_avg_val=buy_avg10)

add_vrvp(bma_bv, bma_bm, col=2)
add_rsi_panel(bma_ts, bma_rsi, col=1)

# KTB3 (cols 3,4)
add_candle(bm31_bars, bm31_ts, col=3, name="KTB3")
add_ind(bm31_ts, bm31_s20, bm31_s60, bm31_vwap, col=3, legend=False)
add_markers(agg_3,  col=3, show_legend=False, sell_avg_val=sell_avg3,  buy_avg_val=buy_avg3)
fig.add_hline(y=bm31_bars[-1]["close"],
              line=dict(color="#58a6ff",width=1,dash="dot"), row=1, col=3,
              annotation_text=f"KTB3 종가 {bm31_bars[-1]['close']:.2f}  캐리{OVN_KTB3:+d}",
              annotation_position="right", annotation_font_color="#58a6ff")
add_vrvp(bm31_bv, bm31_bm, col=4)
add_rsi_panel(bm31_ts, bm31_rsi, col=3)

pad10=(bma_hi-bma_lo)*0.13; pad3=(bm31_hi-bm31_lo)*0.13
fig.update_yaxes(range=[bma_lo-pad10,bma_hi+pad10], gridcolor=GRID, color=TXT, title_text="KTB10", row=1, col=1)
fig.update_yaxes(range=[bma_lo-pad10,bma_hi+pad10], matches="y", showticklabels=False, row=1, col=2)
fig.update_yaxes(range=[bm31_lo-pad3,bm31_hi+pad3], gridcolor=GRID, color=TXT, title_text="KTB3", side="right", row=1, col=3)
fig.update_yaxes(range=[bm31_lo-pad3,bm31_hi+pad3], matches="y3", showticklabels=False, row=1, col=4)
fig.update_yaxes(range=[0,100], gridcolor=GRID, color=TXT, title_text="RSI(14)", side="left",  row=2, col=1)
fig.update_yaxes(range=[0,100], gridcolor=GRID, color=TXT, title_text="RSI(14)", side="right", row=2, col=3)

for row, col in [(1,1),(1,3),(2,1),(2,3)]:
    fig.update_xaxes(range=[X_LEFT,X_RIGHT], gridcolor=GRID, color=TXT,
                     showticklabels=(row==2), tickformat="%H:%M" if row==2 else None,
                     row=row, col=col)
for col in [2,4]:
    fig.update_xaxes(showticklabels=False, showgrid=False, autorange="reversed", row=1, col=col)
fig.update_xaxes(rangeslider_visible=False)

fig.update_layout(
    template="plotly_dark", paper_bgcolor=BG, plot_bgcolor=BG,
    font=dict(color="#e6edf3", size=11),
    title=dict(text=(
        f"KTB10 (BMA1!) 체결포함  ·  KTB3 (BM31!) 캐리{OVN_KTB3:+d}  |  "
        f"{DATE_DISP} · 5분봉  |  "
        f"SELL {total_sell_q} / BUY {total_buy_q}  avg {sell_avg:.3f}/{buy_avg:.3f}"
    ), font=dict(size=11)),
    height=620, margin=dict(l=55,r=80,t=52,b=35),
    legend=dict(orientation="h",y=1.07,x=0,font=dict(size=10)),
    hovermode="closest",
)

print("[4/6] 차트 저장...")
chart_path = ASSET_DIR / f"ktb_{DATE_SLUG}_interactive.html"
div = fig.to_html(include_plotlyjs=True, full_html=False, default_height="620px")
with open(chart_path, "w", encoding="utf-8") as f:
    f.write(div)
print(f"      저장: {chart_path.name}  ({len(div)//1000:,}KB)")


# ── 6. 저널 HTML 생성 ─────────────────────────────────────────────────────
def fmt_pnl(v):
    if v == 0: return '<span style="color:#8b949e">—</span>'
    c="#3fb950" if v>0 else "#f85149"
    return f'<span style="color:{c}">{v/10000:+.0f}만</span>'

def fmt_side(s):
    if s=="매도": return '<span class="side-sell">▼ SELL</span>'
    return '<span class="side-buy">▲ BUY</span>'

net_ktb10 = sell_q10 - buy_q10          # KTB10 순매도 = 숏 증가
final_pos_ktb10 = OVN_KTB10 + buy_q10 - sell_q10
net_ktb3  = sell_q3 - buy_q3           # KTB3 순매도 = 숏 증가
final_pos_ktb3  = OVN_KTB3  + buy_q3  - sell_q3

rows_html = ""
pos = OVN_KTB10; cpnl = 0
for side, qty, price, hh, mm, pnl in sorted(agg, key=lambda x: (x[3], x[4])):
    pos += (-qty if side=="매도" else qty)
    cpnl += pnl
    pc = "#f85149" if pos < 0 else "#3fb950"
    rows_html += f"""
      <tr>
        <td>{hh:02d}:{mm:02d}</td><td>{fmt_side(side)}</td><td>{qty}</td>
        <td>{price:.2f}</td>
        <td style="color:{pc};font-weight:600">{pos:+d}</td>
        <td>{fmt_pnl(pnl)}</td><td>{fmt_pnl(cpnl)}</td>
      </tr>"""

html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>KTB 저널 — {DATE_DISP}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{background:#0d1117;color:#e6edf3;font-family:'Segoe UI',Malgun Gothic,sans-serif;font-size:13px;padding:18px}}
  h1{{font-size:18px;font-weight:700;margin-bottom:4px}}
  h2{{font-size:11px;font-weight:600;color:#8b949e;letter-spacing:.5px;text-transform:uppercase;margin:18px 0 8px}}
  .sub{{color:#8b949e;font-size:12px;margin-bottom:16px}}
  .cards{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px}}
  .card{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:12px 16px;flex:1;min-width:140px}}
  .card .lbl{{font-size:11px;color:#8b949e;margin-bottom:4px}}
  .card .val{{font-size:20px;font-weight:700}}
  .card .sub2{{font-size:11px;color:#8b949e;margin-top:3px}}
  .long{{color:#3fb950}}.short{{color:#f85149}}.neu{{color:#e6edf3}}
  .chart-wrap{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:10px;margin-bottom:16px;overflow:hidden}}
  table{{width:100%;border-collapse:collapse;font-size:12px}}
  th{{background:#161b22;color:#8b949e;padding:6px 8px;text-align:left;border-bottom:2px solid #21262d;position:sticky;top:0}}
  td{{padding:5px 8px;border-bottom:1px solid #161b22}}
  tr:hover td{{background:#161b22}}
  .side-sell{{color:#f85149;font-weight:700}}.side-buy{{color:#3fb950;font-weight:700}}
  .tw{{max-height:400px;overflow-y:auto;border:1px solid #21262d;border-radius:6px;margin-bottom:16px}}
  .sr{{background:#161b22!important}}.sr td{{font-weight:700;border-top:2px solid #21262d}}
  .cb{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:14px;margin-bottom:12px}}
  .cb label{{font-size:11px;color:#8b949e;display:block;margin-bottom:6px}}
  #ca{{width:100%;min-height:110px;background:#0d1117;color:#e6edf3;border:1px solid #30363d;border-radius:4px;padding:10px;font-size:13px;font-family:inherit;resize:vertical;outline:none}}
  #ca:focus{{border-color:#58a6ff}}
  .btn{{background:#238636;color:#fff;border:none;border-radius:6px;padding:8px 18px;font-size:13px;cursor:pointer}}
  .btn:hover{{background:#2ea043}}
  #ss{{font-size:11px;color:#8b949e;margin-left:8px}}
  .note{{font-size:11px;color:#6e7681;margin-top:6px}}
</style>
</head>
<body>
<h1>KTB 선물 트레이딩 저널</h1>
<p class="sub">{DATE_DISP} &nbsp;|&nbsp; NH선물 {len(fills)}건 ({len(agg)}버킷) &nbsp;|&nbsp; KTB10·KTB3 병행 차트</p>

<h2>포지션 요약</h2>
<div class="cards">
  <div class="card"><div class="lbl">KTB10 오버나잇</div>
    <div class="val {'short' if OVN_KTB10<0 else 'long' if OVN_KTB10>0 else 'neu'}">{OVN_KTB10:+d}</div>
    <div class="sub2">전일이월</div></div>
  <div class="card"><div class="lbl">KTB10 당일체결</div>
    <div class="val neu">S{sell_q10} / B{buy_q10}</div>
    <div class="sub2">{"숏 추가" if buy_q10-sell_q10<0 else "커버/롱" if buy_q10-sell_q10>0 else "체결 없음"} ({buy_q10-sell_q10:+d})</div></div>
  <div class="card"><div class="lbl">KTB10 마감추정</div>
    <div class="val {'short' if final_pos_ktb10<0 else 'long' if final_pos_ktb10>0 else 'neu'}">{final_pos_ktb10:+d}</div>
    <div class="sub2">미결잔고 PDF 확인 필요</div></div>
  <div class="card"><div class="lbl">KTB3 오버나잇</div>
    <div class="val {'short' if OVN_KTB3<0 else 'long' if OVN_KTB3>0 else 'neu'}">{OVN_KTB3:+d}</div>
    <div class="sub2">전일이월</div></div>
  <div class="card"><div class="lbl">KTB3 당일체결</div>
    <div class="val neu">S{sell_q3} / B{buy_q3}</div>
    <div class="sub2">{"숏 추가" if buy_q3-sell_q3<0 else "커버/롱" if buy_q3-sell_q3>0 else "체결 없음"} ({buy_q3-sell_q3:+d})</div></div>
  <div class="card"><div class="lbl">KTB3 마감추정</div>
    <div class="val {'short' if final_pos_ktb3<0 else 'long' if final_pos_ktb3>0 else 'neu'}">{final_pos_ktb3:+d}</div>
    <div class="sub2">미결잔고 PDF 확인 필요</div></div>
  <div class="card"><div class="lbl">갱신차금(오버나잇MTM)</div>
    <div class="val {'long' if total_ovn_pnl>0 else 'short' if total_ovn_pnl<0 else 'neu'}">{total_ovn_pnl/10000:+.0f}만</div>
    <div class="sub2">KTB3 {ovn_pnl_ktb3/10000:+.0f}만 / KTB10 {ovn_pnl_ktb10/10000:+.0f}만</div></div>
  <div class="card"><div class="lbl">건별손익(당일체결)</div>
    <div class="val {'long' if total_pnl>0 else 'short' if total_pnl<0 else 'neu'}">{total_pnl/10000:+.0f}만</div>
    <div class="sub2">수수료 {total_fee/1000:.0f}천원 별도</div></div>
  <div class="card" style="border-color:#58a6ff"><div class="lbl">총손익합계</div>
    <div class="val {'long' if total_pnl_all>0 else 'short' if total_pnl_all<0 else 'neu'}">{total_pnl_all/10000:+.0f}만</div>
    <div class="sub2">수수료 {total_fee/1000:.0f}천원 별도</div></div>
</div>

<h2>차트 (KTB10 체결·KTB3 캐리 — 드래그 확대 / 더블클릭 리셋)</h2>
<div class="chart-wrap">{div}</div>

<h2>체결 내역 ({len(agg)}버킷 / 원본 {len(fills)}건)</h2>
<div class="tw"><table>
  <thead><tr><th>시각</th><th>매매</th><th>수량</th><th>가격</th><th>KTB10잔고</th><th>건별손익</th><th>누계손익</th></tr></thead>
  <tbody>
    <tr class="sr"><td colspan="4">전일이월</td>
      <td style="color:#f85149;font-weight:700">{OVN_KTB10:+d}</td><td>—</td><td>—</td></tr>
{rows_html}
    <tr class="sr"><td colspan="2"><strong>합계</strong></td>
      <td>S:{total_sell_q}/B:{total_buy_q}</td><td>—</td>
      <td style="color:#f85149;font-weight:700">{final_pos_ktb10:+d}</td>
      <td>—</td><td>{fmt_pnl(total_pnl)}</td></tr>
  </tbody>
</table></div>

<h2>내 코멘트</h2>
<div class="cb"><label>오늘 매매 리뷰</label>
  <textarea id="ca" placeholder="진입 이유, 청산 타이밍, 아쉬운 점, 내일 전략..."></textarea></div>
<button class="btn" id="sb">💾 코멘트 포함 저장 (.html)</button>
<span id="ss"></span>
<p class="note">※ Ctrl+S 저장 시 레이아웃 깨짐 — 위 버튼 사용</p>

<script>
document.getElementById('sb').addEventListener('click',function(){{
  var cl=document.documentElement.cloneNode(true);
  ['sb','ss'].forEach(function(id){{var e=cl.querySelector('#'+id);if(e)e.remove();}});
  var ta=cl.querySelector('#ca');
  if(ta){{ta.textContent=document.querySelector('#ca').value;ta.setAttribute('readonly','readonly');}}
  var html='<!DOCTYPE html>\\n'+cl.outerHTML;
  var b=new Blob([html],{{type:'text/html;charset=utf-8'}});
  var u=URL.createObjectURL(b);
  var a=document.createElement('a');
  a.href=u;a.download='journal_{DATE_SLUG}_저장본.html';
  document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(u);
  document.getElementById('ss').textContent='저장 완료 ✓';
}});
</script>
</body></html>"""

print("[5/6] 저널 HTML 저장...")
journal_path = THIS_DIR.parent / f"journal_{DATE_SLUG}.html"
with open(journal_path, "w", encoding="utf-8") as f:
    f.write(html)
print(f"      저장: journal_{DATE_SLUG}.html  ({len(html)//1000:,}KB)")


# ── 7. git 커밋·푸쉬 ──────────────────────────────────────────────────────
print("[6/6] git 커밋·푸쉬...")
import shutil
shutil.copy(journal_path, REPO_DIR / f"journal_{DATE_SLUG}.html")
shutil.copy(chart_path,   REPO_DIR / "journal_assets" / chart_path.name)

repo_scripts = REPO_DIR / "office_pc_pipeline"
shutil.copy(__file__, repo_scripts / "daily_review.py")

subprocess.run(["git", "add",
                f"journal_{DATE_SLUG}.html",
                f"journal_assets/{chart_path.name}",
                "office_pc_pipeline/daily_review.py"],
               cwd=str(REPO_DIR))

msg = (f"daily journal {DATE_SLUG}: "
       f"KTB10 {OVN_KTB10:+d}→{final_pos_ktb10:+d} (S{sell_q10}/B{buy_q10}) | "
       f"KTB3 {OVN_KTB3:+d}→{final_pos_ktb3:+d} (S{sell_q3}/B{buy_q3}) | "
       f"PnL {total_pnl_all/10000:+.0f}만(갱신{total_ovn_pnl/10000:+.0f}+건별{total_pnl/10000:+.0f})\n\n"
       f"Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>")
r = subprocess.run(["git", "commit", "-m", msg], cwd=str(REPO_DIR), capture_output=True, text=True)
print(f"      {r.stdout.strip().splitlines()[0] if r.stdout else r.stderr.strip()}")
r2 = subprocess.run(["git", "push"], cwd=str(REPO_DIR), capture_output=True, text=True)
print(f"      push: {r2.stdout.strip() or r2.stderr.strip()}")

# ── 종료 포지션 ── positions.json 체인 폐기: 다음 날 OVN은 Excel에서 재계산.
end_ktb3  = OVN_KTB3  + buy_q3  - sell_q3
end_ktb10 = OVN_KTB10 + buy_q10 - sell_q10

print(f"\n{'='*60}")
print(f"  완료! journal_{DATE_SLUG}.html 브라우저로 열어서 확인")
print(f"  KTB10: {bma_bars[0]['open']:.2f} → H{bma_hi:.2f}/L{bma_lo:.2f} → {bma_bars[-1]['close']:.2f}")
print(f"  KTB3:  {bm31_bars[0]['open']:.2f} → H{bm31_hi:.2f}/L{bm31_lo:.2f} → {bm31_bars[-1]['close']:.2f}")
print(f"  내일 오버나잇 예상: KTB3 {end_ktb3:+d} / KTB10 {end_ktb10:+d}  (Excel 누적에서 재계산)")
print(f"{'='*60}\n")
