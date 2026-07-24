# -*- coding: utf-8 -*-
"""인포맥스 3807 '엑셀보내기' 결과(xls/xlsx)를 정제 시계열로 변환."""
import datetime as dt
import openpyxl
import pandas as pd

# 열 위치 (0-index): A~K 만 사용
COLS = {
    "time": 0, "index": 1,
    "frgn_sell": 2, "frgn_buy": 3, "frgn_net": 4,   # 외국인합
    "indi_sell": 5, "indi_buy": 6, "indi_net": 7,   # 개인
    "inst_sell": 8, "inst_buy": 9, "inst_net": 10,  # 기관계
}

def parse_export(path, date):
    """path: export xlsx/xls 경로, date: 'YYYY-MM-DD' 조회일자.
    반환: (long_df, summary_dict)"""
    if isinstance(date, str):
        date = dt.date.fromisoformat(date)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = []
    summary = None
    for r in ws.iter_rows(values_only=True):
        c0 = r[0]
        # 데이터행: A열이 time
        if isinstance(c0, dt.time):
            rows.append([
                c0, r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10]
            ])
        # 기간누적 요약행: A열에 '기간' 포함(깨져도 length>0 & B가 None & C 숫자)
        elif c0 is not None and r[2] is not None and r[1] is None:
            summary = r  # 전체 20열 (확장 투자자 포함)
    wb.close()

    df = pd.DataFrame(rows, columns=[
        "time", "index",
        "frgn_sell", "frgn_buy", "frgn_net",
        "indi_sell", "indi_buy", "indi_net",
        "inst_sell", "inst_buy", "inst_net",
    ])
    # datetime 결합 (date 객체 사용) 후, 저장용 date 는 ISO 문자열로 통일
    df["datetime"] = df["time"].apply(lambda t: dt.datetime.combine(date, t))
    df.insert(0, "date", date.isoformat())
    df = df.sort_values("datetime").reset_index(drop=True)

    # 당일 최종 누적 = 마지막(시간 최대) 행
    last = df.iloc[-1]
    summ = {
        "date": date.isoformat(),
        "frgn_net": int(last["frgn_net"]),
        "indi_net": int(last["indi_net"]),
        "inst_net": int(last["inst_net"]),
        "frgn_buy": int(last["frgn_buy"]), "frgn_sell": int(last["frgn_sell"]),
        "indi_buy": int(last["indi_buy"]), "indi_sell": int(last["indi_sell"]),
        "inst_buy": int(last["inst_buy"]), "inst_sell": int(last["inst_sell"]),
        "n_bars": len(df),
    }
    return df, summ

if __name__ == "__main__":
    import sys
    path = r"C:\Users\infomax\Downloads\엑셀보내기 샘플.xlsx"
    df, summ = parse_export(path, "2026-07-24")
    print("행수:", len(df), "| 시간범위:", df["time"].min(), "~", df["time"].max())
    print("\n--- head (아침) ---")
    print(df.head(3).to_string())
    print("\n--- tail (마감) ---")
    print(df.tail(3).to_string())
    print("\n--- 당일 최종 누적 요약 ---")
    for k,v in summ.items():
        print(f"  {k}: {v:,}" if isinstance(v,int) else f"  {k}: {v}")
    # 분당 net flow (누적 diff) 예시
    df["frgn_net_delta"] = df["frgn_net"].diff()
    print("\n분당 외국인 순매수(첫 5, 누적차분):")
    print(df[["time","frgn_net","frgn_net_delta"]].head(6).to_string())
