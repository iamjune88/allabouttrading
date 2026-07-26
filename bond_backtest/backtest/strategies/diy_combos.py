"""
DIY Strategy Builder [ZP] 조합 전략 (A/B/C).
구조: Leading Indicator(진입 방향) + Confirmation Filter(검증).
반환: 목표 포지션 시그널 Series (-1/0/+1).
"""
from __future__ import annotations
import numpy as np
import pandas as pd
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
import indicators as ind


def combo_a_trend(prices: pd.DataFrame, amplitude: int = 2,
                  ema_filter: int = 200, adx_len: int = 14,
                  adx_min: float = 20.0) -> pd.Series:
    """조합 A (보수적 추세): Half Trend + EMA200 필터 + ADX>20."""
    c = prices["close"]
    ht = ind.half_trend(prices, amplitude=amplitude)
    ema200 = ind.ema(c, ema_filter)
    adx_val, _, _ = ind.adx(prices, di_len=adx_len, adx_len=adx_len)

    trending = adx_val > adx_min
    long_ok = (ht > 0) & (c > ema200) & trending
    short_ok = (ht < 0) & (c < ema200) & trending

    sig = pd.Series(0.0, index=c.index)
    sig[long_ok] = 1.0
    sig[short_ok] = -1.0
    return sig.rename(f"A_HalfTrend+EMA{ema_filter}+ADX>{adx_min:g}")


def combo_b_rangefilter(prices: pd.DataFrame, rf_period: int = 20,
                        rf_mult: float = 3.0, ema_fast: int = 50,
                        ema_slow: int = 200, chop_len: int = 14,
                        chop_max: float = 61.8) -> pd.Series:
    """조합 B (균형): Range Filter + 2 EMA(50/200) + Choppiness<61.8."""
    c = prices["close"]
    _, rf_dir = ind.range_filter(prices, period=rf_period, multiplier=rf_mult)
    ema_f = ind.ema(c, ema_fast)
    ema_s = ind.ema(c, ema_slow)
    chop = ind.choppiness(prices, chop_len)

    trending = chop < chop_max
    long_ok = (rf_dir > 0) & (ema_f > ema_s) & trending
    short_ok = (rf_dir < 0) & (ema_f < ema_s) & trending

    sig = pd.Series(0.0, index=c.index)
    sig[long_ok] = 1.0
    sig[short_ok] = -1.0
    return sig.rename(f"B_RangeFilter+EMA{ema_fast}/{ema_slow}+Chop<{chop_max:g}")


def combo_c_rqk_supply(prices: pd.DataFrame, rqk_h: float = 8.0,
                       rqk_r: float = 8.0, rqk_x0: int = 25,
                       rsi_len: int = 14, rsi_ma_len: int = 14,
                       z_window: int = 60, z_entry: float = 0.5,
                       supply: pd.Series = None) -> pd.Series:
    """
    조합 C (공격적): RQK 커널 기울기 + RSI MA 방향 (필수),
    외국인 순매수 z-score는 '역방향 거부(veto)'로만 사용.
    supply: 신뢰 가능한 데일리 외국인 순매수. None이면 미사용.
    """
    c = prices["close"]
    kernel = ind.rqk(c, h=rqk_h, r=rqk_r, x0=rqk_x0)
    k_slope = kernel.diff()

    rsi_v = ind.rsi(c, rsi_len)
    rsi_ma = ind.sma(rsi_v, rsi_ma_len)
    rsi_dir = rsi_ma.diff()

    long_ok = (k_slope > 0) & (rsi_dir >= 0)
    short_ok = (k_slope < 0) & (rsi_dir <= 0)

    tag = "C_RQK+RSIma"
    if supply is not None:
        f = supply.reindex(c.index).ffill()
        mu = f.rolling(z_window, min_periods=z_window // 2).mean()
        sd = f.rolling(z_window, min_periods=z_window // 2).std()
        z = (f - mu) / sd.replace(0, np.nan)
        long_ok = long_ok & ~(z < -z_entry)
        short_ok = short_ok & ~(z > z_entry)
        tag += "+수급veto"

    sig = pd.Series(0.0, index=c.index)
    sig[long_ok] = 1.0
    sig[short_ok] = -1.0
    return sig.rename(tag)
