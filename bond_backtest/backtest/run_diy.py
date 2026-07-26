"""
DIY 조합 전략(A/B/C) 백테스트 러너.
대상: KTB3, KTB10 × 일봉, 5분봉.
결과: output/ 에 성과표(CSV) + 차트(PNG).
"""
from __future__ import annotations
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

import bars_loader as bl
import infomax_parser as ip
from engine import BacktestConfig, run_backtest
from strategies.diy_combos import combo_a_trend, combo_b_rangefilter, combo_c_rqk_supply
import report as rp

OUT = Path(__file__).parent / "output"
OUT.mkdir(exist_ok=True)
# 데일리 수급/가격 원본: 기존 시계열 폴더의 파일을 직접 참조
DAILY_SRC = Path(__file__).parent.parent.parent / "시계열" / "ktb분석_daily.xlsx"

SUPPLY = {}
for prod, key in [("KTB3", "3년"), ("KTB10", "10년")]:
    try:
        SUPPLY[prod] = ip.load_infomax_factor(DAILY_SRC, series_key=key, column="foreign_net")
        print(f"[수급팩터] {prod} 데일리 외국인 순매수 {len(SUPPLY[prod])} obs")
    except Exception as e:
        SUPPLY[prod] = None
        print(f"[수급팩터] {prod} 로드 실패: {e}")

COMBOS = {
    "A_추세정배열": combo_a_trend,
    "B_RangeFilter돌파": combo_b_rangefilter,
    "C_RQK+수급": combo_c_rqk_supply,
}

PRODUCTS = ["KTB3", "KTB10"]
TIMEFRAMES = {
    "일봉": {"tf": "daily", "ppy": 252},
    "5분봉": {"tf": "5min", "ppy": 19656},  # 하루 78봉 × 252
}


def main():
    all_summ = {}
    for product in PRODUCTS:
        for tf_name, tfc in TIMEFRAMES.items():
            try:
                bars = bl.load_bars(product, tfc["tf"])
            except Exception as e:
                print(f"[스킵] {product} {tf_name}: {e}")
                continue
            print("\n" + "=" * 66)
            print(f"{product} / {tf_name}  ({len(bars):,}봉, "
                  f"{bars.index[0]} ~ {bars.index[-1]})")
            print("=" * 66)

            cfg = BacktestConfig(cost_bps=1.0, slippage_bps=0.5,
                                 periods_per_year=tfc["ppy"], allow_short=True)
            results = {}
            for cname, cfn in COMBOS.items():
                try:
                    if cfn is combo_c_rqk_supply and tf_name == "일봉":
                        sig = cfn(bars, supply=SUPPLY.get(product))
                    else:
                        sig = cfn(bars)
                    res = run_backtest(bars, sig, cfg, name=cname)
                    results[cname] = res
                    print(f"  {res}")
                except Exception as e:
                    import traceback
                    print(f"  ⚠ {cname} 실패: {e}")
                    traceback.print_exc()

            if not results:
                continue
            summ = rp.build_summary(results)
            rp.print_summary(summ)
            tag = f"{product}_{tf_name}"
            rp.save_summary_csv(summ, OUT / f"diy_summary_{tag}.csv")
            rp.plot_equity(results, OUT / f"diy_equity_{tag}.png",
                           title=f"DIY 조합 — {product} {tf_name}")
            all_summ[tag] = summ
            print(f"  → output/diy_summary_{tag}.csv, diy_equity_{tag}.png")

    print("\n[전체 완료]", list(all_summ.keys()))


if __name__ == "__main__":
    main()
