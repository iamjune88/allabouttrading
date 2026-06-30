# -*- coding: utf-8 -*-
"""
2026-06-30 KTB10 인터랙티브 차트 (Plotly) — 줌/팬/호버 가능.
정적 PNG 대신 journal HTML에 직접 임베드할 div 조각을 생성한다.
체결 라벨은 상시 텍스트 대신 마커+호버텍스트로 표시(기본 화면을 덜 빽빽하게).
"""
import json
from datetime import datetime, timezone, timedelta

import plotly.graph_objects as go
from plotly.subplots import make_subplots

KST = timezone(timedelta(hours=9))
DAY_BASE = 1781481600 + 15 * 86400  # 2026-06-30 09:00 KST
SESSION_START = DAY_BASE - 900
SESSION_END = DAY_BASE + 6 * 3600 + 45 * 60

with open("_ohlcv_0630.json", encoding="utf-8") as f:
    data = json.load(f)

all_bars = sorted(data["bars"], key=lambda b: b["time"])
all_bars = [b for b in all_bars if b["time"] < SESSION_END]
closes = [b["close"] for b in all_bars]

trades = [
    ("SELL", 10, 106.84, 8, 45), ("SELL", 2, 106.91, 8, 46), ("SELL", 8, 106.89, 8, 49),
    ("SELL", 20, 106.92, 8, 51), ("BUY", 20, 106.89, 8, 55), ("SELL", 20, 106.94, 9, 0),
    ("BUY", 20, 106.93, 9, 0), ("SELL", 30, 106.97, 9, 31), ("BUY", 10, 106.94, 9, 32),
    ("BUY", 10, 106.97, 9, 39), ("BUY", 10, 106.95, 9, 40), ("BUY", 10, 106.96, 9, 41),
    ("SELL", 10, 106.98, 9, 57), ("SELL", 20, 107.09, 10, 12), ("BUY", 10, 107.07, 10, 19),
    ("BUY", 10, 107.06, 10, 21), ("SELL", 30, 107.21, 11, 25), ("BUY", 10, 107.19, 11, 38),
    ("BUY", 10, 107.18, 11, 39), ("BUY", 10, 107.2, 11, 47), ("BUY", 12, 107.27, 14, 50),
    ("BUY", 8, 107.29, 14, 52), ("SELL", 20, 107.29, 14, 53), ("SELL", 50, 107.27, 14, 55),
    ("BUY", 20, 107.25, 15, 0), ("BUY", 20, 107.22, 15, 0), ("BUY", 10, 107.2, 15, 4),
    ("BUY", 50, 107.13, 15, 9), ("SELL", 50, 107.15, 15, 9), ("SELL", 50, 107.16, 15, 19),
    ("BUY", 50, 107.15, 15, 22), ("SELL", 30, 107.19, 15, 28), ("SELL", 30, 107.26, 15, 45),
]


def bar_ts(hh, mm):
    off = (hh - 9) * 3600 + mm * 60
    return ((DAY_BASE + off) // 300) * 300


def sma(series, window):
    out = [None] * len(series)
    for i in range(len(series)):
        if i + 1 >= window:
            out[i] = sum(series[i + 1 - window:i + 1]) / window
    return out


def rsi(series, period=14):
    out = [None] * len(series)
    gains, losses = [], []
    for i in range(1, len(series)):
        d = series[i] - series[i - 1]
        gains.append(max(d, 0))
        losses.append(max(-d, 0))
        if i >= period:
            if i == period:
                avg_g = sum(gains[:period]) / period
                avg_l = sum(losses[:period]) / period
            else:
                avg_g = (out[i - 1]["avg_g"] * (period - 1) + gains[-1]) / period
                avg_l = (out[i - 1]["avg_l"] * (period - 1) + losses[-1]) / period
            rs = avg_g / avg_l if avg_l != 0 else float("inf")
            val = 100 - 100 / (1 + rs) if avg_l != 0 else 100.0
            out[i] = {"val": val, "avg_g": avg_g, "avg_l": avg_l}
    return [o["val"] if o else None for o in out]


sma20_full, sma60_full = sma(closes, 20), sma(closes, 60)
sma120_full, sma200_full = sma(closes, 120), sma(closes, 200)
rsi_full = rsi(closes, 14)

idx_today = [i for i, b in enumerate(all_bars) if b["time"] >= SESSION_START]
bars = [all_bars[i] for i in idx_today]
times = [datetime.fromtimestamp(b["time"], tz=KST) for b in bars]
sma20 = [sma20_full[i] for i in idx_today]
sma60 = [sma60_full[i] for i in idx_today]
sma120 = [sma120_full[i] for i in idx_today]
sma200 = [sma200_full[i] for i in idx_today]
rsi_today = [rsi_full[i] for i in idx_today]

vwap = []
cum_pv, cum_v = 0.0, 0.0
for b in bars:
    tp = (b["high"] + b["low"] + b["close"]) / 3
    cum_pv += tp * b["volume"]
    cum_v += b["volume"]
    vwap.append(cum_pv / cum_v if cum_v else None)

day_high, day_low = max(b["high"] for b in bars), min(b["low"] for b in bars)
N_BINS = 36
bin_edges = [day_low + (day_high - day_low) * i / N_BINS for i in range(N_BINS + 1)]
bin_vol = [0.0] * N_BINS
for b in bars:
    lo, hi, vol = b["low"], b["high"], b["volume"]
    span = max(hi - lo, 0.001)
    for i in range(N_BINS):
        blo, bhi = bin_edges[i], bin_edges[i + 1]
        overlap = max(0.0, min(hi, bhi) - max(lo, blo))
        if overlap > 0:
            bin_vol[i] += vol * (overlap / span)
bin_mid = [(bin_edges[i] + bin_edges[i + 1]) / 2 for i in range(N_BINS)]

BG, GRID, TXT = "#0d1117", "#21262d", "#8b949e"
UP, DOWN = "#3fb950", "#f85149"

fig = make_subplots(
    rows=2, cols=2, row_heights=[0.78, 0.22], column_widths=[0.85, 0.15],
    shared_xaxes=False, horizontal_spacing=0.01, vertical_spacing=0.03,
    specs=[[{}, {}], [{}, None]],
)

fig.add_trace(go.Candlestick(
    x=times, open=[b["open"] for b in bars], high=[b["high"] for b in bars],
    low=[b["low"] for b in bars], close=[b["close"] for b in bars],
    increasing_line_color=UP, decreasing_line_color=DOWN, name="KTB10", showlegend=False,
), row=1, col=1)

for series, name, color in [
    (sma20, "SMA20", "#e3b341"), (sma60, "SMA60", "#39c5cf"),
    (sma120, "SMA120", "#58a6ff"), (sma200, "SMA200", "#f778ba"),
    (vwap, "VWAP", "#bc8cff"),
]:
    fig.add_trace(go.Scatter(x=times, y=series, mode="lines", name=name,
                              line=dict(color=color, width=1.3)), row=1, col=1)

buy_trades = [t for t in trades if t[0] == "BUY"]
sell_trades = [t for t in trades if t[0] == "SELL"]
buy_avg = sum(q * p for _, q, p, *_ in buy_trades) / sum(q for _, q, p, *_ in buy_trades)
sell_avg = sum(q * p for _, q, p, *_ in sell_trades) / sum(q for _, q, p, *_ in sell_trades)

for side, tlist, color, symbol in [("BUY", buy_trades, "#44BB44", "triangle-up"),
                                    ("SELL", sell_trades, "#FF4444", "triangle-down")]:
    xs = [datetime.fromtimestamp(bar_ts(hh, mm), tz=KST) for _, q, p, hh, mm in tlist]
    ys = [p for _, q, p, hh, mm in tlist]
    qs = [q for _, q, p, hh, mm in tlist]
    hover = [f"{side} {int(q)}@{p}<br>{hh:02d}:{mm:02d} KST" for _, q, p, hh, mm in tlist]
    fig.add_trace(go.Scatter(
        x=xs, y=ys, mode="markers", name=f"{side} 체결",
        marker=dict(symbol=symbol, size=10, color=color, line=dict(width=1, color="#0d1117")),
        text=hover, hovertemplate="%{text}<extra></extra>",
    ), row=1, col=1)

x_right = datetime.fromtimestamp(DAY_BASE, tz=KST).replace(hour=17, minute=30)
x_left = times[0] - timedelta(minutes=8)
fig.add_hline(y=buy_avg, line=dict(color="#88FF88", width=1, dash="dash"), row=1, col=1,
              annotation_text=f"BUY avg {buy_avg:.2f}", annotation_position="right",
              annotation_font_color="#88FF88")
fig.add_hline(y=sell_avg, line=dict(color="#FF8888", width=1, dash="dash"), row=1, col=1,
              annotation_text=f"SELL avg {sell_avg:.2f}", annotation_position="right",
              annotation_font_color="#FF8888")

fig.add_trace(go.Bar(
    x=bin_vol, y=bin_mid, orientation="h", marker_color="#39c5cf", opacity=0.55,
    name="VRVP(당일)", showlegend=False,
), row=1, col=2)

fig.add_trace(go.Scatter(x=times, y=rsi_today, mode="lines", name="RSI14",
                          line=dict(color="#d29922", width=1.3), showlegend=False), row=2, col=1)
fig.add_hline(y=70, line=dict(color=DOWN, width=0.7, dash="dot"), row=2, col=1)
fig.add_hline(y=30, line=dict(color=UP, width=0.7, dash="dot"), row=2, col=1)
fig.add_hline(y=50, line=dict(color=TXT, width=0.5), row=2, col=1)

pad = (day_high - day_low) * 0.12
fig.update_xaxes(range=[x_left, x_right], row=1, col=1, gridcolor=GRID, color=TXT,
                  showticklabels=False)
fig.update_xaxes(range=[x_left, x_right], row=2, col=1, gridcolor=GRID, color=TXT,
                  tickformat="%H:%M")
fig.update_xaxes(showticklabels=False, row=1, col=2)
fig.update_yaxes(range=[day_low - pad, day_high + pad], row=1, col=1, gridcolor=GRID, color=TXT,
                  title_text="가격")
fig.update_yaxes(range=[day_low - pad, day_high + pad], row=1, col=2, matches="y", showticklabels=False)
fig.update_xaxes(autorange="reversed", row=1, col=2, showgrid=False)
fig.update_yaxes(range=[0, 100], row=2, col=1, gridcolor=GRID, color=TXT, title_text="RSI(14)")

fig.update_layout(
    template="plotly_dark",
    paper_bgcolor=BG, plot_bgcolor=BG,
    font=dict(color="#e6edf3", size=11),
    title=dict(text="KTB10 Futures · KRX:BMA1!  |  2026-06-30 (화)  |  5분봉 — 드래그로 확대, 더블클릭으로 리셋",
               font=dict(size=14)),
    height=560,
    margin=dict(l=50, r=110, t=45, b=35),
    legend=dict(orientation="h", y=1.06, x=0, font=dict(size=10)),
    xaxis_rangeslider_visible=False,
    hovermode="closest",
)

div = fig.to_html(include_plotlyjs=True, full_html=False, default_height="560px")
out_path = r"C:\Users\infomax\Desktop\claude협업\journal_assets\ktb10_20260630_interactive.html"
with open(out_path, "w", encoding="utf-8") as f:
    f.write(div)
print("saved", out_path, len(div), "chars")
