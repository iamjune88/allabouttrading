# -*- coding: utf-8 -*-
"""인포맥스 3807 '엑셀보내기' 결과(xls/xlsx)를 정제 시계열로 변환.

2026-07 양식 변경:
  - 투자자 구성: (외국인합 / 개인 / 기관계)  ->  (외국인합 / 증권 / 투신 / 은행)  [4그룹]
  - 시간축: 1분  ->  30초 단위
  - 열 배치(0-index):
      0 시간, 1 지수,
      2~4  외국인합  매도/매수/순매수,
      5~7  증권      매도/매수/순매수,
      8~10 투신      매도/매수/순매수,
      11~13 은행     매도/매수/순매수,
      (14 이후는 공란)
"""
import datetime as dt
import openpyxl
import pandas as pd

# 투자자 그룹: (키, 매도열, 매수열, 순매수열)  — 0-index
GROUPS = [
    ("frgn",  2,  3,  4),   # 외국인합
    ("secu",  5,  6,  7),   # 증권
    ("trust", 8,  9,  10),  # 투신
    ("bank",  11, 12, 13),  # 은행
]
# 파싱에 필요한 최소 열 개수 (은행 순매수 = 13 이므로 14열)
NCOL = 14

# long-format DataFrame 값 컬럼명과, 각 컬럼이 읽어올 0-index 열 위치(짝을 맞춰 둠)
VALUE_COLS = [f"{g}_{s}" for g, *_ in GROUPS for s in ("sell", "buy", "net")]
VALUE_IDX  = [i for _, sell, buy, net in GROUPS for i in (sell, buy, net)]


def _cell(row, idx):
    """행에서 idx 열 값을 안전하게 꺼낸다(짧은 행/공란은 None)."""
    return row[idx] if idx < len(row) else None


def parse_export(path, date):
    """path: export xlsx/xls 경로, date: 'YYYY-MM-DD' 조회일자.
    반환: (long_df, summary_dict)"""
    if isinstance(date, str):
        date = dt.date.fromisoformat(date)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(values_only=True):
        c0 = _cell(r, 0)
        # 데이터행: A열이 time(시:분:초)
        if isinstance(c0, dt.time):
            rows.append([c0, _cell(r, 1)] + [_cell(r, i) for i in VALUE_IDX])
    wb.close()

    df = pd.DataFrame(rows, columns=["time", "index"] + VALUE_COLS)
    # datetime 결합(date 객체) 후 저장용 date 는 ISO 문자열로 통일
    df["datetime"] = df["time"].apply(lambda t: dt.datetime.combine(date, t))
    df.insert(0, "date", date.isoformat())
    df = df.sort_values("datetime").reset_index(drop=True)

    # 당일 최종 누적 = 마지막(시간 최대) 행. 마감행은 전 그룹이 채워져 있음.
    last = df.iloc[-1]

    def _i(col):
        v = last[col]
        return int(v) if pd.notna(v) else 0

    summ = {"date": date.isoformat()}
    for g, *_ in GROUPS:
        summ[f"{g}_net"] = _i(f"{g}_net")
    for g, *_ in GROUPS:
        summ[f"{g}_buy"] = _i(f"{g}_buy")
        summ[f"{g}_sell"] = _i(f"{g}_sell")
    summ["n_bars"] = len(df)

    # --- 절단(오전 결손) 감지 ---
    # 그리드 export 는 최신 ~1200행만 유지(FIFO). 하루가 1200봉을 넘으면(비정상 패딩/연장
    # 등) 오전이 버퍼에서 밀려나 잘린다. 즉 절단된 날은 반드시 n_bars 가 상한(~1200)에 닿고,
    # 정상일(9시장 959 / 8:45장 989 < 1200)과 겹치지 않으므로 이 단일 신호로 충분히 정확하다.
    # (첫 봉 순매수 비율 같은 지표는 개장 동시호가 때문에 오탐이 많아 쓰지 않는다.)
    summ["first_time"] = str(df.iloc[0]["time"])
    summ["truncated"] = int(len(df) >= 1195)
    return df, summ


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\infomax\Downloads\엑셀보내기 샘플.xlsx"
    date = sys.argv[2] if len(sys.argv) > 2 else "2026-07-24"
    df, summ = parse_export(path, date)
    print("행수:", len(df), "| 시간범위:", df["time"].min(), "~", df["time"].max())
    print("\n--- head (아침) ---")
    print(df.head(3).to_string())
    print("\n--- tail (마감) ---")
    print(df.tail(3).to_string())
    print("\n--- 당일 최종 누적 요약 ---")
    for k, v in summ.items():
        print(f"  {k}: {v:,}" if isinstance(v, int) else f"  {k}: {v}")
    # 30초당 순flow(누적 diff) 예시
    df["frgn_net_delta"] = df["frgn_net"].diff()
    print("\n30초당 외국인 순매수(첫 6, 누적차분):")
    print(df[["time", "frgn_net", "frgn_net_delta"]].head(6).to_string())
