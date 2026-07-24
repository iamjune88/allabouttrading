# -*- coding: utf-8 -*-
"""daily.csv 품질 점검: 연속으로 값이 완전히 같은 날(휴장에 전일자료가 딸려온 의심)을 표시."""
import os, sys
import pandas as pd
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass
BASE = os.path.dirname(os.path.abspath(__file__))
# 사용: python verify_daily.py [tag]   (예: python verify_daily.py 2025 -> daily_2025.csv)
_TAG = sys.argv[1] if len(sys.argv) > 1 else ""
_SUF = f"_{_TAG}" if _TAG else ""
DAILY = os.path.join(BASE, "data", f"daily{_SUF}.csv")

def main():
    if not os.path.exists(DAILY):
        print("daily.csv 없음"); return
    df = pd.read_csv(DAILY).sort_values("date").reset_index(drop=True)
    print(f"총 {len(df)}일  ({df['date'].min()} ~ {df['date'].max()})")
    key = ["frgn_net", "secu_net", "trust_net", "bank_net", "frgn_buy"]
    dup = df[key].eq(df[key].shift()).all(axis=1)
    susp = df[dup]
    if len(susp):
        print(f"\n[!] 전일과 값이 완전히 동일한 {len(susp)}일 (휴장/중복 의심 - 확인 요망):")
        for _, r in susp.iterrows():
            print(f"  {r['date']}  외국인 {int(r['frgn_net']):+,}  증권 {int(r['secu_net']):+,}")
    else:
        print("\n의심 중복 없음. OK")

if __name__ == "__main__":
    main()
