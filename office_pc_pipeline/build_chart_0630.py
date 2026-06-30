# -*- coding: utf-8 -*-
"""
2026-06-30 KTB10 커스텀 차트 (OHLCV 직접 렌더링 버전).
오늘은 가용 사전 이력이 223봉이라 SMA20/60/120/200 전부 정상 시계열로 계산 가능.
체결은 NH(분단위)+SS(초단위)를 합쳐 (방향,분,가격) 단위로 합산해서 라벨로 사용(60건→33버킷).
"""
import json
import matplotlib
matplotlib.use("Agg")
matplotlib.rcParams["font.family"] = "Malgun Gothic"
matplotlib.rcParams["axes.unicode_minus"] = False
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Rectangle
from matplotlib.gridspec import GridSpec
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))

DAY_BASE = 1781481600 + 15 * 86400  # 2026-06-30 09:00 KST
SESSION_START = DAY_BASE - 900
SESSION_END = DAY_BASE + 6 * 3600 + 45 * 60

with open("_ohlcv_0630.json", encoding="utf-8") as f:
    data = json.load(f)

all_bars = sorted(data["bars"], key=lambda b: b["time"])
all_bars = [b for b in all_bars if b["time"] < SESSION_END]
closes = [b["close"] for b in all_bars]

# (side, qty, price, hh, mm) — NH(분)+SS(초→분) 합산 버킷, 시간순
trades = [
    ("SELL", 10, 106.84, 8, 45),
    ("SELL", 2, 106.91, 8, 46),
    ("SELL", 8, 106.89, 8, 49),
    ("SELL", 20, 106.92, 8, 51),
    ("BUY", 20, 106.89, 8, 55),
    ("SELL", 20, 106.94, 9, 0),
    ("BUY", 20, 106.93, 9, 0),
    ("SELL", 30, 106.97, 9, 31),
    ("BUY", 10, 106.94, 9, 32),
    ("BUY", 10, 106.97, 9, 39),
    ("BUY", 10, 106.95, 9, 40),
    ("BUY", 10, 106.96, 9, 41),
    ("SELL", 10, 106.98, 9, 57),
    ("SELL", 20, 107.09, 10, 12),
    ("BUY", 10, 107.07, 10, 19),
    ("BUY", 10, 107.06, 10, 21),
    ("SELL", 30, 107.21, 11, 25),
    ("BUY", 10, 107.19, 11, 38),
    ("BUY", 10, 107.18, 11, 39),
    ("BUY", 10, 107.2, 11, 47),
    ("BUY", 12, 107.27, 14, 50),
    ("BUY", 8, 107.29, 14, 52),
    ("SELL", 20, 107.29, 14, 53),
    ("SELL", 50, 107.27, 14, 55),
    ("BUY", 20, 107.25, 15, 0),
    ("BUY", 20, 107.22, 15, 0),
    ("BUY", 10, 107.2, 15, 4),
    ("BUY", 50, 107.13, 15, 9),
    ("SELL", 50, 107.15, 15, 9),
    ("SELL", 50, 107.16, 15, 19),
    ("BUY", 50, 107.15, 15, 22),
    ("SELL", 30, 107.19, 15, 28),
    ("SELL", 30, 107.26, 15, 45),
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


sma20_full = sma(closes, 20)
sma60_full = sma(closes, 60)
sma120_full = sma(closes, 120)
sma200_full = sma(closes, 200)


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

day_high = max(b["high"] for b in bars)
day_low = min(b["low"] for b in bars)
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

BG, PANEL, GRID, TXT, TXT2 = "#0d1117", "#161b22", "#21262d", "#8b949e", "#e6edf3"
UP, DOWN = "#3fb950", "#f85149"
SELL_C, BUY_C = "#FF4444", "#44BB44"

fig = plt.figure(figsize=(13, 6.2), dpi=130)
fig.patch.set_facecolor(BG)
gs = GridSpec(2, 2, width_ratios=[7, 1.3], height_ratios=[4, 1], hspace=0.05, wspace=0.02, figure=fig)
ax = fig.add_subplot(gs[0, 0])
ax_vrvp = fig.add_subplot(gs[0, 1], sharey=ax)
ax_rsi = fig.add_subplot(gs[1, 0], sharex=ax)
ax_dummy = fig.add_subplot(gs[1, 1])
ax_dummy.axis("off")
for a in (ax, ax_vrvp, ax_rsi):
    a.set_facecolor(BG)

width = timedelta(minutes=3.2)
for t, b in zip(times, bars):
    color = UP if b["close"] >= b["open"] else DOWN
    ax.plot([t, t], [b["low"], b["high"]], color=color, linewidth=0.8, zorder=2)
    rect = Rectangle(
        (mdates.date2num(t) - width.total_seconds() / 86400 / 2, min(b["open"], b["close"])),
        width.total_seconds() / 86400,
        max(abs(b["close"] - b["open"]), 0.0008),
        facecolor=color, edgecolor=color, zorder=3,
    )
    ax.add_patch(rect)

pad = (day_high - day_low) * 0.18
y_top = day_high + pad * 1.6
y_bot = day_low - pad * 1.2
x_left = times[0] - timedelta(minutes=8)
x_right = datetime.fromtimestamp(DAY_BASE, tz=KST).replace(hour=17, minute=30)
ax.set_xlim(x_left, x_right)
ax.set_ylim(y_bot, y_top)


def plot_partial(ax, times, vals, color, label, lw=1.1, ls="-"):
    segs_t, segs_v = [], []
    for t, v in zip(times, vals):
        if v is None:
            if segs_t:
                ax.plot(segs_t, segs_v, color=color, linewidth=lw, linestyle=ls, label=label, zorder=4)
                label = None
                segs_t, segs_v = [], []
            continue
        segs_t.append(t)
        segs_v.append(v)
    if segs_t:
        ax.plot(segs_t, segs_v, color=color, linewidth=lw, linestyle=ls, label=label, zorder=4)


plot_partial(ax, times, sma20, "#e3b341", "SMA20")
plot_partial(ax, times, sma60, "#39c5cf", "SMA60")
plot_partial(ax, times, sma120, "#58a6ff", "SMA120")
plot_partial(ax, times, sma200, "#f778ba", "SMA200")
plot_partial(ax, times, vwap, "#bc8cff", "VWAP", lw=1.3)

labels = []
for side, qty, price, hh, mm in trades:
    ts = bar_ts(hh, mm)
    t = datetime.fromtimestamp(ts, tz=KST)
    labels.append({"side": side, "qty": qty, "price": price, "hh": hh, "mm": mm, "t": t})
labels.sort(key=lambda L: (L["hh"], L["mm"]))

min_gap = (y_top - y_bot) * 0.040
placed_sell, placed_buy = [], []
for L in labels:
    off = (y_top - y_bot) * 0.040
    base = L["price"] + off if L["side"] == "SELL" else L["price"] - off
    bucket = placed_sell if L["side"] == "SELL" else placed_buy
    y = base
    for prev_t, prev_y in bucket:
        if abs((L["t"] - prev_t).total_seconds()) < 22 * 60:
            if L["side"] == "SELL" and y < prev_y + min_gap:
                y = prev_y + min_gap
            if L["side"] == "BUY" and y > prev_y - min_gap:
                y = prev_y - min_gap
    bucket.append((L["t"], y))
    L["label_y"] = y

for L in labels:
    color = SELL_C if L["side"] == "SELL" else BUY_C
    arrow = "▼" if L["side"] == "SELL" else "▲"
    text = f"{arrow}{L['hh']:02d}:{L['mm']:02d} {L['side']} {int(L['qty'])}@{L['price']}"
    va = "bottom" if L["side"] == "SELL" else "top"
    ax.annotate(
        text, xy=(L["t"], L["price"]), xytext=(L["t"], L["label_y"]),
        fontsize=6.6, color=color, fontweight="bold", ha="center", va=va,
        arrowprops=dict(arrowstyle="-", color=color, lw=0.5, alpha=0.6),
        zorder=6,
    )

buy_avg = sum(q * p for s, q, p, *_ in trades if s == "BUY") / sum(q for s, q, p, *_ in trades if s == "BUY")
sell_avg = sum(q * p for s, q, p, *_ in trades if s == "SELL") / sum(q for s, q, p, *_ in trades if s == "SELL")
ax.axhline(buy_avg, color="#88FF88", linewidth=1, linestyle=(0, (4, 3)), zorder=1)
ax.axhline(sell_avg, color="#FF8888", linewidth=1, linestyle=(0, (4, 3)), zorder=1)
ax.text(x_right, buy_avg, f" BUY avg {buy_avg:.2f}", color="#88FF88", fontsize=8.5, va="center", ha="left")
ax.text(x_right, sell_avg, f" SELL avg {sell_avg:.2f}", color="#FF8888", fontsize=8.5, va="center", ha="left")

ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M", tz=KST))
ax.xaxis.set_major_locator(mdates.MinuteLocator(interval=30))
ax.tick_params(colors=TXT, labelsize=8, labelbottom=False)
for spine in ax.spines.values():
    spine.set_color(GRID)
ax.grid(True, color=GRID, linewidth=0.6, alpha=0.7)
ax.set_title("KTB10 Futures · KRX:BMA1!  |  2026-06-30 (화)  |  5분봉", color=TXT2, fontsize=12, fontweight="bold", loc="left", pad=10)
leg = ax.legend(loc="upper left", fontsize=7.5, facecolor=PANEL, edgecolor=GRID, framealpha=0.9, ncol=2)
for t in leg.get_texts():
    t.set_color(TXT2)

ax_vrvp.barh(bin_mid, bin_vol, height=(day_high - day_low) / N_BINS * 0.9, color="#39c5cf", alpha=0.55, zorder=2)
ax_vrvp.set_ylim(y_bot, y_top)
ax_vrvp.set_xlim(0, max(bin_vol) * 1.15)
ax_vrvp.invert_xaxis()
ax_vrvp.tick_params(colors=TXT, labelsize=7, left=False, labelleft=False, labelbottom=False, bottom=False)
for spine in ax_vrvp.spines.values():
    spine.set_color(GRID)
ax_vrvp.set_title("VRVP(당일)", color=TXT, fontsize=8.5, pad=4)

plot_partial(ax_rsi, times, rsi_today, "#d29922", "RSI14", lw=1.1)
ax_rsi.axhline(70, color=DOWN, linewidth=0.7, linestyle=(0, (3, 3)), alpha=0.6)
ax_rsi.axhline(30, color=UP, linewidth=0.7, linestyle=(0, (3, 3)), alpha=0.6)
ax_rsi.axhline(50, color=TXT, linewidth=0.5, alpha=0.4)
ax_rsi.set_ylim(0, 100)
ax_rsi.set_xlim(x_left, x_right)
ax_rsi.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M", tz=KST))
ax_rsi.xaxis.set_major_locator(mdates.MinuteLocator(interval=30))
ax_rsi.tick_params(colors=TXT, labelsize=8)
for spine in ax_rsi.spines.values():
    spine.set_color(GRID)
ax_rsi.grid(True, color=GRID, linewidth=0.5, alpha=0.5)
ax_rsi.set_ylabel("RSI(14)", color=TXT, fontsize=8.5)
last_rsi = [v for v in rsi_today if v is not None]
if last_rsi:
    ax_rsi.text(x_right, last_rsi[-1], f" {last_rsi[-1]:.1f}", color="#d29922", fontsize=8, va="center", ha="left")

plt.tight_layout()
out_path = r"C:\Users\infomax\Desktop\claude협업\journal_assets\ktb10_20260630.png"
plt.savefig(out_path, facecolor=BG, bbox_inches="tight")
print("saved", out_path)
print("last RSI:", last_rsi[-1] if last_rsi else None)
print("last VWAP:", vwap[-1])
print("last SMA20/60/120/200:", sma20[-1], sma60[-1], sma120[-1], sma200[-1])
print("buy_avg", buy_avg, "sell_avg", sell_avg)
