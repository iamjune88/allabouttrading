# -*- coding: utf-8 -*-
"""
2026-07-01 KTB10/KTB3 인터랙티브 차트 (Plotly) — 2x4 대칭 레이아웃
양쪽 동일 지표: SMA20 / SMA60 / VWAP / VRVP / RSI(14)
X축: 동일 시간범위 공유
"""
import json
from datetime import datetime, timezone, timedelta

import plotly.graph_objects as go
from plotly.subplots import make_subplots

KST = timezone(timedelta(hours=9))
DAY_BASE      = 1781481600 + 16 * 86400   # 2026-07-01 09:00 KST
SESSION_START = DAY_BASE - 900             # 08:45 KST
SESSION_END   = DAY_BASE + 6 * 3600 + 45 * 60  # 15:45 KST

# ── OHLCV ────────────────────────────────────────────────────────────────
with open("_ohlcv_0701_BMA1.json",  encoding="utf-8-sig") as f:
    data_bma  = json.load(f)
with open("_ohlcv_0701_BM31.json",  encoding="utf-8-sig") as f:
    data_bm31 = json.load(f)

def prep_bars(data):
    bars = sorted(data["bars"], key=lambda b: b["time"])
    return [b for b in bars if b["time"] < SESSION_END + 300]

all_bma  = prep_bars(data_bma)
all_bm31 = prep_bars(data_bm31)

# ── KTB10 체결내역 ────────────────────────────────────────────────────────
trades = [
    ("SELL", 20, 106.79, 8, 45), ("SELL",  3, 106.73, 8, 54),
    ("SELL", 17, 106.73, 8, 55), ("SELL", 20, 106.69, 8, 59),
    ("SELL",  9, 106.61, 9,  2), ("SELL", 10, 106.56, 9,  5),
    ("SELL", 10, 106.51, 9,  8), ("BUY",  18, 106.59, 9, 22),
    ("BUY",  30, 106.54, 10, 11), ("SELL", 10, 106.44, 10, 16),
    ("SELL", 30, 106.44, 10, 17), ("SELL", 10, 106.41, 10, 19),
    ("BUY",  10, 106.46, 10, 22), ("SELL", 17, 106.36, 10, 46),
    ("BUY",  20, 106.39, 10, 53), ("SELL", 20, 106.31, 11,  6),
    ("BUY",  20, 106.35, 11,  8), ("SELL", 20, 106.27, 11, 31),
    ("SELL", 20, 106.21, 11, 39), ("BUY",  25, 106.21, 12,  6),
    ("BUY",  15, 106.22, 12,  6), ("SELL", 30, 106.07, 13,  8),
    ("BUY",  30, 106.17, 13, 30), ("SELL", 30, 106.14, 14, 12),
    ("BUY",  30, 106.17, 14, 14), ("SELL", 20, 106.19, 14, 29),
    ("SELL", 20, 106.22, 14, 35), ("BUY",  20, 106.24, 14, 39),
    ("BUY",  30, 106.33, 15,  0), ("SELL",  5, 106.31, 15,  1),
    ("SELL", 20, 106.28, 15,  6), ("BUY",  35, 106.17, 15, 23),
    ("SELL", 10, 106.17, 15, 28), ("SELL",  5, 106.15, 15, 30),
    ("SELL", 30, 106.07, 15, 45),
]
buy_ts   = [t for t in trades if t[0] == "BUY"]
sell_ts  = [t for t in trades if t[0] == "SELL"]
total_sell_q = sum(q for _, q, *_ in sell_ts)
total_buy_q  = sum(q for _, q, *_ in buy_ts)
buy_avg  = sum(q*p for _,q,p,*_ in buy_ts)  / sum(q for _,q,*_ in buy_ts)
sell_avg = sum(q*p for _,q,p,*_ in sell_ts) / sum(q for _,q,*_ in sell_ts)


def bar_ts(hh, mm):
    return ((DAY_BASE + (hh - 9) * 3600 + mm * 60) // 300) * 300


def sma(series, window):
    out = [None] * len(series)
    for i in range(len(series)):
        if i + 1 >= window:
            out[i] = sum(series[i + 1 - window: i + 1]) / window
    return out


def rsi(series, period=14):
    out = [None] * len(series)
    gains, losses, state = [], [], {}
    for i in range(1, len(series)):
        d = series[i] - series[i - 1]
        gains.append(max(d, 0));  losses.append(max(-d, 0))
        if i >= period:
            if i == period:
                ag = sum(gains[:period]) / period
                al = sum(losses[:period]) / period
            else:
                ag = (state["ag"] * (period - 1) + gains[-1]) / period
                al = (state["al"] * (period - 1) + losses[-1]) / period
            state = {"ag": ag, "al": al}
            rs = ag / al if al else float("inf")
            out[i] = 100 - 100 / (1 + rs) if al else 100.0
    return out


def process(all_bars):
    closes      = [b["close"] for b in all_bars]
    sma20_full  = sma(closes, 20)
    sma60_full  = sma(closes, 60)
    rsi_full    = rsi(closes, 14)
    idx         = [i for i, b in enumerate(all_bars) if b["time"] >= SESSION_START]
    bars        = [all_bars[i] for i in idx]
    times       = [datetime.fromtimestamp(b["time"], tz=KST) for b in bars]
    sma20       = [sma20_full[i] for i in idx]
    sma60       = [sma60_full[i] for i in idx]
    rsi_t       = [rsi_full[i]   for i in idx]
    # VWAP (세션앵커)
    vwap, cum_pv, cum_v = [], 0.0, 0.0
    for b in bars:
        tp = (b["high"] + b["low"] + b["close"]) / 3
        cum_pv += tp * b["volume"];  cum_v += b["volume"]
        vwap.append(cum_pv / cum_v if cum_v else None)
    # VRVP
    day_hi = max(b["high"] for b in bars)
    day_lo = min(b["low"]  for b in bars)
    N = 36
    edges   = [day_lo + (day_hi - day_lo) * k / N for k in range(N + 1)]
    bvol    = [0.0] * N
    for b in bars:
        lo, hi, vol = b["low"], b["high"], b["volume"]
        span = max(hi - lo, 0.001)
        for k in range(N):
            ov = max(0.0, min(hi, edges[k+1]) - max(lo, edges[k]))
            if ov: bvol[k] += vol * ov / span
    bmid = [(edges[k] + edges[k+1]) / 2 for k in range(N)]
    return bars, times, sma20, sma60, rsi_t, vwap, bvol, bmid, day_hi, day_lo


(bma_bars,  bma_times,  bma_sma20,  bma_sma60,  bma_rsi,
 bma_vwap,  bma_bvol,  bma_bmid,  bma_hi,  bma_lo)  = process(all_bma)
(bm31_bars, bm31_times, bm31_sma20, bm31_sma60, bm31_rsi,
 bm31_vwap, bm31_bvol, bm31_bmid, bm31_hi, bm31_lo) = process(all_bm31)

# ── 공통 x축 범위 ─────────────────────────────────────────────────────────
X_LEFT  = datetime(2026, 7, 1,  8, 37, tzinfo=KST)
X_RIGHT = datetime(2026, 7, 1, 15, 55, tzinfo=KST)

BG, GRID, TXT = "#0d1117", "#21262d", "#8b949e"
UP, DOWN      = "#3fb950", "#f85149"

IND_COLORS = {
    "SMA20": "#e3b341", "SMA60": "#39c5cf",
    "VWAP":  "#bc8cff", "RSI":   "#d29922",
}

# ── 2×4 대칭 레이아웃 ─────────────────────────────────────────────────────
# row1: [KTB10캔들 | KTB10 VRVP | KTB3캔들 | KTB3 VRVP]
# row2: [KTB10 RSI |     —      | KTB3 RSI |     —    ]
fig = make_subplots(
    rows=2, cols=4,
    row_heights=[0.76, 0.24],
    column_widths=[0.435, 0.065, 0.435, 0.065],
    shared_xaxes=False,
    horizontal_spacing=0.008,
    vertical_spacing=0.03,
    specs=[[{}, {}, {}, {}],
           [{}, None, {}, None]],
)


def add_candle(fig, bars, times, row, col, name):
    fig.add_trace(go.Candlestick(
        x=times,
        open=[b["open"] for b in bars], high=[b["high"] for b in bars],
        low=[b["low"]   for b in bars], close=[b["close"] for b in bars],
        increasing_line_color=UP, decreasing_line_color=DOWN,
        name=name, showlegend=False,
    ), row=row, col=col)


def add_indicators(fig, times, sma20, sma60, vwap, row, col, show_legend):
    for series, name, color in [
        (sma20, "SMA20", IND_COLORS["SMA20"]),
        (sma60, "SMA60", IND_COLORS["SMA60"]),
        (vwap,  "VWAP",  IND_COLORS["VWAP"]),
    ]:
        fig.add_trace(go.Scatter(
            x=times, y=series, mode="lines", name=name,
            line=dict(color=color, width=1.3),
            showlegend=show_legend,
        ), row=row, col=col)


def add_vrvp(fig, bvol, bmid, row, col):
    fig.add_trace(go.Bar(
        x=bvol, y=bmid, orientation="h",
        marker_color="#39c5cf", opacity=0.55,
        name="VRVP", showlegend=False,
    ), row=row, col=col)


def add_rsi(fig, times, rsi_vals, row, col):
    fig.add_trace(go.Scatter(
        x=times, y=rsi_vals, mode="lines", name="RSI14",
        line=dict(color=IND_COLORS["RSI"], width=1.2), showlegend=False,
    ), row=row, col=col)
    for level, color in [(70, DOWN), (50, TXT), (30, UP)]:
        dash = "dot" if level != 50 else "solid"
        w    = 0.7  if level != 50 else 0.5
        fig.add_hline(y=level, line=dict(color=color, width=w, dash=dash),
                      row=row, col=col)


# ── KTB10 (left side: cols 1, 2) ─────────────────────────────────────────
add_candle(fig, bma_bars, bma_times, row=1, col=1, name="KTB10")
add_indicators(fig, bma_times, bma_sma20, bma_sma60, bma_vwap,
               row=1, col=1, show_legend=True)

# 체결 마커
for side, tlist, color, sym in [
    ("BUY",  buy_ts,  "#44BB44", "triangle-up"),
    ("SELL", sell_ts, "#FF4444", "triangle-down"),
]:
    xs = [datetime.fromtimestamp(bar_ts(hh, mm), tz=KST) for _,q,p,hh,mm in tlist]
    ys = [p for _,q,p,hh,mm in tlist]
    hover = [f"{side} {int(q)}@{p:.2f}  {hh:02d}:{mm:02d}" for _,q,p,hh,mm in tlist]
    fig.add_trace(go.Scatter(
        x=xs, y=ys, mode="markers", name=f"{side}",
        marker=dict(symbol=sym, size=9, color=color,
                    line=dict(width=1, color="#0d1117")),
        text=hover, hovertemplate="%{text}<extra></extra>", showlegend=True,
    ), row=1, col=1)

# 평균단가선
for avg, color, label in [
    (buy_avg,  "#88FF88", f"BUY avg {buy_avg:.2f}"),
    (sell_avg, "#FF8888", f"SELL avg {sell_avg:.2f}"),
]:
    fig.add_hline(y=avg, line=dict(color=color, width=1, dash="dash"),
                  row=1, col=1,
                  annotation_text=label, annotation_position="right",
                  annotation_font_color=color)

add_vrvp(fig, bma_bvol, bma_bmid, row=1, col=2)
add_rsi(fig,  bma_times, bma_rsi,  row=2, col=1)

# ── KTB3 (right side: cols 3, 4) ─────────────────────────────────────────
add_candle(fig, bm31_bars, bm31_times, row=1, col=3, name="KTB3")
add_indicators(fig, bm31_times, bm31_sma20, bm31_sma60, bm31_vwap,
               row=1, col=3, show_legend=False)

add_vrvp(fig, bm31_bvol, bm31_bmid, row=1, col=4)
add_rsi(fig,  bm31_times, bm31_rsi,  row=2, col=3)

# KTB3 캐리 포지션 참조선
bm31_close = bm31_bars[-1]["close"]
fig.add_hline(y=bm31_close, line=dict(color="#58a6ff", width=1, dash="dot"),
              row=1, col=3,
              annotation_text=f"KTB3 종가 {bm31_close:.2f}  캐리+240",
              annotation_position="right", annotation_font_color="#58a6ff")

# SMA120/200 불가 주석 (양쪽)
for col, xref, yref, y in [
    (1, "x",  "y",  bma_lo  - (bma_hi  - bma_lo)  * 0.06),
    (3, "x3", "y3", bm31_lo - (bm31_hi - bm31_lo) * 0.06),
]:
    fig.add_annotation(
        x=X_RIGHT, y=y,
        text="SMA120/200 사전이력부족(17봉)", showarrow=False,
        font=dict(color=TXT, size=8), xref=xref, yref=yref, xanchor="right",
    )

# ── 축 공통 설정 ─────────────────────────────────────────────────────────
pad10 = (bma_hi  - bma_lo)  * 0.13
pad3  = (bm31_hi - bm31_lo) * 0.13

# KTB10 캔들 y축
fig.update_yaxes(range=[bma_lo-pad10, bma_hi+pad10],
                 gridcolor=GRID, color=TXT, title_text="KTB10",
                 row=1, col=1)
# KTB10 VRVP y축 — 캔들과 동일 범위
fig.update_yaxes(range=[bma_lo-pad10, bma_hi+pad10],
                 matches="y", showticklabels=False, row=1, col=2)

# KTB3 캔들 y축
fig.update_yaxes(range=[bm31_lo-pad3, bm31_hi+pad3],
                 gridcolor=GRID, color=TXT, title_text="KTB3", side="right",
                 row=1, col=3)
# KTB3 VRVP y축 — KTB3 캔들과 동일 범위
fig.update_yaxes(range=[bm31_lo-pad3, bm31_hi+pad3],
                 matches="y3", showticklabels=False, row=1, col=4)

# RSI y축 — KTB10: 좌측, KTB3: 우측 (대칭)
fig.update_yaxes(range=[0, 100], gridcolor=GRID, color=TXT,
                 title_text="RSI(14)", side="left",  row=2, col=1)
fig.update_yaxes(range=[0, 100], gridcolor=GRID, color=TXT,
                 title_text="RSI(14)", side="right", row=2, col=3)

# x축 — 모든 차트 동일 범위, VRVP만 reversed
for row, col in [(1,1),(1,3),(2,1),(2,3)]:
    tick = "%H:%M" if row == 2 else None
    fig.update_xaxes(
        range=[X_LEFT, X_RIGHT],
        gridcolor=GRID, color=TXT,
        showticklabels=(row == 2),
        tickformat=tick,
        row=row, col=col,
    )
for col in [2, 4]:
    fig.update_xaxes(showticklabels=False, showgrid=False,
                     autorange="reversed", row=1, col=col)

# 모든 x축 rangeslider 비활성화 (캔들 기본값이 RSI 패널로 침범하는 것 방지)
fig.update_xaxes(rangeslider_visible=False)

fig.update_layout(
    template="plotly_dark",
    paper_bgcolor=BG, plot_bgcolor=BG,
    font=dict(color="#e6edf3", size=11),
    title=dict(
        text=(
            "KTB10 (BMA1!) · 체결포함 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            "KTB3 (BM31!) · 캐리 +240 &nbsp;&nbsp;|&nbsp;&nbsp;"
            f"2026-07-01 (수) · 5분봉 &nbsp;|&nbsp; "
            f"NH 체결: SELL {total_sell_q} / BUY {total_buy_q}  "
            f"| SELL avg {sell_avg:.3f} / BUY avg {buy_avg:.3f}"
        ),
        font=dict(size=11),
    ),
    height=620,
    margin=dict(l=55, r=80, t=52, b=35),
    legend=dict(orientation="h", y=1.07, x=0, font=dict(size=10)),
    hovermode="closest",
)

out_path = r"C:\Users\infomax\Desktop\claude협업\journal_assets\ktb_20260701_interactive.html"
div = fig.to_html(include_plotlyjs=True, full_html=False, default_height="620px")
with open(out_path, "w", encoding="utf-8") as f:
    f.write(div)
print(f"saved  {len(div):,} chars")
print(f"KTB10  O:{all_bma[len(all_bma)-len(bma_bars)]['close']:.2f}  "
      f"H:{bma_hi:.2f}  L:{bma_lo:.2f}  C:{bma_bars[-1]['close']:.2f}")
print(f"KTB3   O:{all_bm31[len(all_bm31)-len(bm31_bars)]['close']:.2f}  "
      f"H:{bm31_hi:.2f}  L:{bm31_lo:.2f}  C:{bm31_bars[-1]['close']:.2f}")
