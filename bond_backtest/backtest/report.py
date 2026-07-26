"""성과 리포트 — 비교표(CSV/콘솔) + 자산곡선/드로다운 차트."""
from __future__ import annotations
from pathlib import Path
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

try:
    plt.rcParams["font.family"] = "Malgun Gothic"  # 한글 (Windows)
    plt.rcParams["axes.unicode_minus"] = False
except Exception:
    pass

STAT_COLS = [
    ("total_return", "총수익률", "{:.1%}"),
    ("cagr", "CAGR", "{:.2%}"),
    ("ann_vol", "연변동성", "{:.2%}"),
    ("sharpe", "샤프", "{:.2f}"),
    ("sortino", "소르티노", "{:.2f}"),
    ("mdd", "MDD", "{:.2%}"),
    ("calmar", "칼마", "{:.2f}"),
    ("win_rate", "승률", "{:.1%}"),
    ("profit_factor", "손익비", "{:.2f}"),
    ("n_trades", "거래수", "{:.0f}"),
    ("exposure", "노출도", "{:.1%}"),
]


def build_summary(results: dict) -> pd.DataFrame:
    rows = {}
    for name, res in results.items():
        s = res.stats
        rows[name] = {label: s.get(key, float("nan")) for key, label, _ in STAT_COLS}
    df = pd.DataFrame(rows).T
    if "샤프" in df.columns:
        df = df.sort_values("샤프", ascending=False)
    return df


def print_summary(df: pd.DataFrame):
    fmt = {label: f for _, label, f in STAT_COLS}
    disp = df.copy()
    for label, f in fmt.items():
        if label in disp.columns:
            disp[label] = disp[label].map(lambda v: f.format(v) if pd.notna(v) else "-")
    print(disp.to_string())


def save_summary_csv(df: pd.DataFrame, path):
    df.to_csv(path, encoding="utf-8-sig")


def plot_equity(results: dict, out_path, title="전략별 자산곡선"):
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 9),
                                   gridspec_kw={"height_ratios": [3, 1]}, sharex=True)
    for name, res in results.items():
        ax1.plot(res.equity.index, res.equity.values, label=name, linewidth=1.3)
    ax1.set_title(title)
    ax1.set_ylabel("자산 (기준 100)")
    ax1.legend(loc="upper left", fontsize=9)
    ax1.grid(alpha=0.3)

    for name, res in results.items():
        eq = res.equity
        dd = eq / eq.cummax() - 1.0
        ax2.plot(dd.index, dd.values, label=name, linewidth=1.0)
    ax2.set_ylabel("드로다운")
    ax2.set_xlabel("날짜")
    ax2.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(out_path, dpi=120)
    plt.close(fig)
