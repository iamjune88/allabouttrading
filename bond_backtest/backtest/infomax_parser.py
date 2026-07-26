"""
인포맥스(Infomax) 엑셀 다운로드 전용 파서.
레이아웃: row0=메타, row1=시리즈명, row2=헤더, row3~=데이터(일자 내림차순).
한 시트에 여러 시리즈 블록이 가로로 이어질 수 있어 각 블록을 분리해 로드.
"""
from __future__ import annotations
import pandas as pd
import numpy as np

COL_MAP = {
    "일자": "date", "거래량": "volume", "현재가": "close", "종가": "close",
    "기준가": "prev_close", "시가": "open", "고가": "high", "저가": "low",
    "미결제약정수량": "open_interest",
    "외국인합순매수수량": "foreign_net", "외국인순매수수량": "foreign_net",
}


def parse_infomax(path, sheet=0, series_name_row=1, header_row=2, data_start=3):
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    names = raw.iloc[series_name_row]
    headers = raw.iloc[header_row]

    block_starts = [i for i, v in enumerate(names) if pd.notna(v) and str(v).strip()]
    if not block_starts:
        raise ValueError("시리즈명 행에서 블록을 찾지 못했습니다.")
    block_starts.append(raw.shape[1])

    # 일자 컬럼은 첫 블록에만 있고 모든 블록이 공유
    date_col_idx = None
    for c in range(raw.shape[1]):
        if str(headers.iloc[c]).strip() == "일자":
            date_col_idx = c
            break
    shared_date = raw.iloc[data_start:, date_col_idx] if date_col_idx is not None else None

    out = {}
    for bi in range(len(block_starts) - 1):
        c0, c1 = block_starts[bi], block_starts[bi + 1]
        series_name = str(names.iloc[c0]).strip()
        block = raw.iloc[data_start:, c0:c1].copy()
        block.columns = [str(headers.iloc[c]).strip() for c in range(c0, c1)]

        std = pd.DataFrame()
        for src, dst in COL_MAP.items():
            if src in block.columns and dst not in std.columns:
                std[dst] = block[src]
        if "date" not in std.columns and shared_date is not None:
            std["date"] = shared_date.values
        if "date" not in std.columns or "close" not in std.columns:
            continue

        std["date"] = pd.to_datetime(std["date"], errors="coerce")
        for c in std.columns:
            if c != "date":
                std[c] = pd.to_numeric(std[c], errors="coerce")
        std = std.dropna(subset=["date", "close"])
        std = std[std["close"] != 0]
        std = std.sort_values("date").set_index("date")
        std = std[~std.index.duplicated(keep="last")]
        out[series_name] = std
    return out


def load_infomax_price(path, series_key=None, sheet=0):
    blocks = parse_infomax(path, sheet=sheet)
    if not blocks:
        raise ValueError(f"{path}: 파싱된 시리즈 없음")
    if series_key is None:
        return next(iter(blocks.values()))
    for name, df in blocks.items():
        if series_key in name:
            return df
    raise KeyError(f"'{series_key}' 시리즈를 찾지 못함. 가능: {list(blocks.keys())}")


def load_infomax_factor(path, series_key=None, column="foreign_net", sheet=0):
    df = load_infomax_price(path, series_key=series_key, sheet=sheet)
    if column not in df.columns:
        raise KeyError(f"'{column}' 컬럼 없음. 가능: {list(df.columns)}")
    return df[column].dropna().rename(f"{series_key or ''}_{column}")
