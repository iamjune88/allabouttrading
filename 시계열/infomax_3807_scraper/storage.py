# -*- coding: utf-8 -*-
"""수집 상태 관리 + 통합 CSV 누적 + 원본 보관.

산출물 (data/ 폴더):
  - intraday.csv : 30초봉 누적치 long-format (date,time,datetime,index,투자자별 매도/매수/순매수)
                   투자자 = 외국인합(frgn)/증권(secu)/투신(trust)/은행(bank)
  - daily.csv    : 일별 최종 누적 요약 (한 행 = 하루)
  - state.json   : {날짜: 'ok'|'holiday'|'error'} 재실행시 이어받기용
원본:
  - raw/3807_YYYYMMDD.xlsx : 엑셀보내기 원본 그대로 보관
"""
import json, os, shutil, datetime as dt
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
RAW_ROOT  = os.path.join(BASE, "raw")
RAW_DIR   = RAW_ROOT
DATA_DIR  = os.path.join(BASE, "data")

for d in (RAW_DIR, DATA_DIR):
    os.makedirs(d, exist_ok=True)


def _paths(tag=""):
    """tag 가 있으면 파일명에 접미어(_tag)를 붙여 연도별 등으로 분리 저장."""
    suf = f"_{tag}" if tag else ""
    return (os.path.join(DATA_DIR, f"intraday{suf}.csv"),
            os.path.join(DATA_DIR, f"daily{suf}.csv"),
            os.path.join(DATA_DIR, f"state{suf}.json"))


# 기본(무태그) 경로. set_output_tag() 로 런타임에 바꿀 수 있음.
INTRADAY, DAILY, STATE = _paths()


def set_output_tag(tag: str):
    """intraday/daily/state 출력 파일을 <name>_<tag> 로, raw 원본은 raw/<tag>/ 하위폴더로 전환.
    상품이 다르면(예: 3년 vs 10년) 같은 날짜라도 raw 파일명(3807_YYYYMMDD.xlsx)이 겹치므로,
    태그별 하위폴더로 분리해 서로 덮어쓰지 않게 한다. (무태그 = raw/ 루트, 기존 3년물 그대로)"""
    global INTRADAY, DAILY, STATE, RAW_DIR
    INTRADAY, DAILY, STATE = _paths(tag)
    RAW_DIR = os.path.join(RAW_ROOT, tag) if tag else RAW_ROOT
    os.makedirs(RAW_DIR, exist_ok=True)
    return INTRADAY, DAILY, STATE


def load_state() -> dict:
    if os.path.exists(STATE):
        with open(STATE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(state: dict):
    tmp = STATE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=0, sort_keys=True)
    os.replace(tmp, STATE)


def mark(state: dict, date, status: str):
    state[str(date)] = status
    save_state(state)


def is_done(state: dict, date) -> bool:
    """ok/holiday 는 완료로 간주(재조회 스킵). error 는 재시도 대상."""
    return state.get(str(date)) in ("ok", "holiday")


def archive_raw(src_path: str, date) -> str:
    """엑셀보내기 원본을 raw/3807_YYYYMMDD.xlsx 로 복사."""
    d = dt.date.fromisoformat(str(date))
    dst = os.path.join(RAW_DIR, f"3807_{d:%Y%m%d}.xlsx")
    shutil.copyfile(src_path, dst)
    return dst


def append_intraday(df: pd.DataFrame):
    """중복 방지(같은 date 재수집 시 기존 제거) + 항상 시간순 정렬 유지."""
    header = not os.path.exists(INTRADAY)
    if not header:
        try:
            old = pd.read_csv(INTRADAY)
            old = old[old["date"] != str(df["date"].iloc[0])]
            df_all = pd.concat([old, df], ignore_index=True)
            df_all = df_all.sort_values(["date", "time"]).reset_index(drop=True)
            df_all.to_csv(INTRADAY, index=False)
            return
        except Exception:
            pass
    df.to_csv(INTRADAY, mode="a", header=header, index=False)


def append_daily(summ: dict):
    row = pd.DataFrame([summ])
    if os.path.exists(DAILY):
        old = pd.read_csv(DAILY)
        old = old[old["date"] != str(summ["date"])]
        pd.concat([old, row], ignore_index=True).sort_values("date").to_csv(DAILY, index=False)
    else:
        row.to_csv(DAILY, index=False)


if __name__ == "__main__":
    print("BASE:", BASE)
    print("state:", load_state())
