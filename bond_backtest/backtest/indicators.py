"""
Pine Script 지표의 파이썬 포팅.
DIY Custom Strategy Builder [ZP]에서 쓰이는 핵심 지표들을 재현한다.
모두 look-ahead 없이 과거 데이터만 사용.
"""
from __future__ import annotations
import numpy as np
import pandas as pd


def ema(s: pd.Series, length: int) -> pd.Series:
    return s.ewm(span=length, adjust=False).mean()


def sma(s: pd.Series, length: int) -> pd.Series:
    return s.rolling(length, min_periods=length).mean()


def rma(s: pd.Series, length: int) -> pd.Series:
    """Wilder's RMA (Pine ta.rma). alpha=1/length."""
    return s.ewm(alpha=1.0 / length, adjust=False).mean()


def true_range(df: pd.DataFrame) -> pd.Series:
    h, l, c = df["high"], df["low"], df["close"]
    pc = c.shift(1)
    return pd.concat([(h - l), (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1)


def atr(df: pd.DataFrame, length: int = 14) -> pd.Series:
    return rma(true_range(df), length)


def adx(df: pd.DataFrame, di_len: int = 14, adx_len: int = 14):
    h, l = df["high"], df["low"]
    up = h.diff()
    down = -l.diff()
    plus_dm = np.where((up > down) & (up > 0), up, 0.0)
    minus_dm = np.where((down > up) & (down > 0), down, 0.0)
    tr = true_range(df)
    trur = rma(tr, di_len)
    plus_di = 100 * rma(pd.Series(plus_dm, index=df.index), di_len) / trur
    minus_di = 100 * rma(pd.Series(minus_dm, index=df.index), di_len) / trur
    dx = 100 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    adx_val = rma(dx, adx_len)
    return adx_val, plus_di, minus_di


def rsi(s: pd.Series, length: int = 14) -> pd.Series:
    delta = s.diff()
    up = rma(delta.clip(lower=0), length)
    down = rma(-delta.clip(upper=0), length)
    rs = up / down.replace(0, np.nan)
    return 100 - 100 / (1 + rs)


def choppiness(df: pd.DataFrame, length: int = 14) -> pd.Series:
    tr = true_range(df)
    atr_sum = tr.rolling(length, min_periods=length).sum()
    hh = df["high"].rolling(length, min_periods=length).max()
    ll = df["low"].rolling(length, min_periods=length).min()
    rng = (hh - ll).replace(0, np.nan)
    return 100 * np.log10(atr_sum / rng) / np.log10(length)


def half_trend(df: pd.DataFrame, amplitude: int = 2, channel_dev: float = 2.0):
    """Half Trend 지표. +1(상승)/-1(하락) 시리즈 반환."""
    h, l, c = df["high"], df["low"], df["close"]
    hp = h.rolling(amplitude, min_periods=1).max()
    lp = l.rolling(amplitude, min_periods=1).min()
    sma_h = sma(h, amplitude).bfill()
    sma_l = sma(l, amplitude).bfill()

    n = len(df)
    trend = np.zeros(n, dtype=int)
    up = np.zeros(n)
    down = np.zeros(n)
    max_low = l.iloc[0]
    min_high = h.iloc[0]
    cl = c.values; hh = h.values; ll = l.values
    hpv = hp.values; lpv = lp.values; smh = sma_h.values; sml = sma_l.values

    t = 0
    for i in range(n):
        if i == 0:
            up[i] = down[i] = cl[i]
            continue
        prev_t = t
        if prev_t == 0:
            max_low = max(max_low, ll[i])
            if hpv[i] < max_low and cl[i] < ll[i-1]:
                t = 1
                min_high = hh[i]
        else:
            min_high = min(min_high, hh[i])
            if lpv[i] > min_high and cl[i] > hh[i-1]:
                t = 0
                max_low = ll[i]
        trend[i] = t
        if t == 0:
            up[i] = max(up[i-1], smh[i]) if up[i-1] else smh[i]
        else:
            down[i] = min(down[i-1], sml[i]) if down[i-1] else sml[i]

    direction = pd.Series(np.where(trend == 0, 1, -1), index=df.index)
    return direction


def range_filter(df: pd.DataFrame, period: int = 20, multiplier: float = 3.0):
    """Range Filter (smooth range 방식). 반환: (filt, direction)."""
    src = df["close"]
    n = len(src)
    absdiff = src.diff().abs()
    avrng = ema(absdiff, period)
    smoothrng = ema(avrng, period * 2 - 1) * multiplier

    filt = np.zeros(n)
    x = src.values
    sr = smoothrng.values
    filt[0] = x[0]
    for i in range(1, n):
        prev = filt[i-1]
        r = sr[i] if not np.isnan(sr[i]) else 0.0
        if x[i] > prev:
            filt[i] = prev if (x[i] - r) < prev else (x[i] - r)
        else:
            filt[i] = prev if (x[i] + r) > prev else (x[i] + r)

    filt_s = pd.Series(filt, index=df.index)
    direction = np.sign(filt_s.diff()).replace(0, np.nan).ffill().fillna(0)
    return filt_s, direction


def rqk(src: pd.Series, h: float = 8.0, r: float = 8.0, x0: int = 25) -> pd.Series:
    """Rational Quadratic Kernel 회귀 (endpoint 방식). look-ahead 없음."""
    x = src.values
    n = len(x)
    out = np.full(n, np.nan)
    size = x0
    idx = np.arange(size)
    weights = (1.0 + (idx ** 2) / (2.0 * r * h * h)) ** (-r)
    wsum = weights.sum()
    for t in range(n):
        if t < size - 1:
            continue
        window = x[t - size + 1: t + 1][::-1]
        out[t] = np.dot(window, weights) / wsum
    return pd.Series(out, index=src.index)
