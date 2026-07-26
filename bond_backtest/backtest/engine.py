"""
백테스트 엔진 (벡터화)
- 입력: 가격 시계열(close) + 전략이 만든 포지션 시그널(-1/0/+1)
- 처리: 시그널을 1봉 시프트(look-ahead 방지) → 수익률 계산 → 거래비용 차감
- 출력: 자산곡선(equity curve), 거래 로그, 성과지표
"""
from __future__ import annotations
from dataclasses import dataclass, field
import numpy as np
import pandas as pd


@dataclass
class BacktestConfig:
    cost_bps: float = 1.0          # 편도 거래비용 (bp, 1bp=0.01%)
    slippage_bps: float = 0.5      # 편도 슬리피지 (bp)
    periods_per_year: int = 252    # 일봉 기준 연율화 계수 (5분봉이면 조정)
    allow_short: bool = True       # 숏 허용 여부
    initial_equity: float = 100.0  # 기준 자산 (수익률 지수)


@dataclass
class BacktestResult:
    equity: pd.Series
    returns: pd.Series
    positions: pd.Series
    trades: pd.DataFrame
    stats: dict = field(default_factory=dict)
    name: str = "strategy"

    def __repr__(self):
        s = self.stats
        return (f"<BacktestResult {self.name}: "
                f"CAGR={s.get('cagr',0):.2%} Sharpe={s.get('sharpe',0):.2f} "
                f"MDD={s.get('mdd',0):.2%} Trades={s.get('n_trades',0)}>")


def run_backtest(prices: pd.DataFrame, signal: pd.Series,
                 config: BacktestConfig | None = None,
                 name: str = "strategy") -> BacktestResult:
    cfg = config or BacktestConfig()
    px = prices["close"].astype(float)
    sig = signal.reindex(px.index).fillna(0.0).clip(-1, 1)
    if not cfg.allow_short:
        sig = sig.clip(lower=0.0)

    # look-ahead 방지: 오늘 만든 시그널은 다음 봉부터 적용
    pos = sig.shift(1).fillna(0.0)

    asset_ret = px.pct_change().fillna(0.0)
    gross = pos * asset_ret

    # 거래비용: 포지션이 바뀐 만큼(회전율) × (비용+슬리피지)
    turnover = pos.diff().abs().fillna(pos.abs())
    cost_rate = (cfg.cost_bps + cfg.slippage_bps) / 1e4
    cost = turnover * cost_rate

    net = gross - cost
    equity = (1.0 + net).cumprod() * cfg.initial_equity

    trades = _extract_trades(px, pos, cost_rate)
    stats = compute_stats(net, equity, pos, trades, cfg)
    return BacktestResult(equity=equity, returns=net, positions=pos,
                          trades=trades, stats=stats, name=name)


def _extract_trades(px: pd.Series, pos: pd.Series, cost_rate: float) -> pd.DataFrame:
    rows = []
    cur_pos = 0.0
    entry_px = None
    entry_dt = None
    for dt, p in pos.items():
        if p != cur_pos:
            if cur_pos != 0.0 and entry_px is not None:
                exit_px = px.loc[dt]
                pnl = cur_pos * (exit_px / entry_px - 1.0) - 2 * cost_rate
                rows.append({
                    "entry_dt": entry_dt, "exit_dt": dt,
                    "side": "LONG" if cur_pos > 0 else "SHORT",
                    "entry_px": entry_px, "exit_px": exit_px,
                    "ret": pnl, "bars": (px.index.get_loc(dt) - px.index.get_loc(entry_dt)),
                })
            if p != 0.0:
                entry_px = px.loc[dt]
                entry_dt = dt
            cur_pos = p
    return pd.DataFrame(rows)


def compute_stats(net: pd.Series, equity: pd.Series, pos: pd.Series,
                  trades: pd.DataFrame, cfg: BacktestConfig) -> dict:
    n = len(net)
    if n == 0 or equity.iloc[-1] <= 0:
        return {"cagr": 0, "sharpe": 0, "mdd": 0, "n_trades": 0}
    total_ret = equity.iloc[-1] / equity.iloc[0] - 1.0
    years = n / cfg.periods_per_year
    cagr = (equity.iloc[-1] / equity.iloc[0]) ** (1 / years) - 1 if years > 0 else 0.0

    ann_vol = net.std() * np.sqrt(cfg.periods_per_year)
    sharpe = (net.mean() * cfg.periods_per_year) / ann_vol if ann_vol > 0 else 0.0
    downside = net[net < 0].std() * np.sqrt(cfg.periods_per_year)
    sortino = (net.mean() * cfg.periods_per_year) / downside if downside > 0 else 0.0

    roll_max = equity.cummax()
    dd = equity / roll_max - 1.0
    mdd = dd.min()
    calmar = cagr / abs(mdd) if mdd < 0 else 0.0

    win_rate = (trades["ret"] > 0).mean() if len(trades) else 0.0
    avg_win = trades.loc[trades["ret"] > 0, "ret"].mean() if (trades["ret"] > 0).any() else 0.0
    avg_loss = trades.loc[trades["ret"] < 0, "ret"].mean() if (trades["ret"] < 0).any() else 0.0
    pf = (trades.loc[trades["ret"] > 0, "ret"].sum() /
          abs(trades.loc[trades["ret"] < 0, "ret"].sum())) if (trades["ret"] < 0).any() else np.inf
    exposure = (pos != 0).mean()

    return {
        "total_return": total_ret, "cagr": cagr, "ann_vol": ann_vol,
        "sharpe": sharpe, "sortino": sortino, "mdd": mdd, "calmar": calmar,
        "n_trades": len(trades), "win_rate": win_rate,
        "avg_win": avg_win, "avg_loss": avg_loss, "profit_factor": pf,
        "exposure": exposure, "start": str(net.index[0].date()),
        "end": str(net.index[-1].date()),
    }
