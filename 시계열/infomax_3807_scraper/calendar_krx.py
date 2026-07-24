# -*- coding: utf-8 -*-
"""후보 거래일 생성. 주말/(가능하면)공휴일 제외.
확정 판단은 런타임에 export 결과(빈 데이터=휴장)로 이뤄진다 — 이 목록은 후보일 뿐."""
import datetime as dt

try:
    import holidays as _holidays  # pip install holidays (선택)
    _KR = _holidays.SouthKorea
except Exception:
    _KR = None

# KRX 고정 휴장(공휴일 라이브러리가 놓치는 것): 근로자의날, 연말 폐장일
def _krx_extra_holiday(d: dt.date) -> bool:
    if (d.month, d.day) == (5, 1):      # 근로자의 날
        return True
    if (d.month, d.day) == (12, 31):    # 연말 폐장일 (평일이면 휴장)
        return True
    return False


def candidate_trading_days(start, end, ascending=True):
    """start~end(포함) 사이 주말·공휴일 제외 후보일 리스트."""
    start = dt.date.fromisoformat(str(start))
    end   = dt.date.fromisoformat(str(end))
    kr = _KR(years=range(start.year, end.year + 1)) if _KR else {}
    days, d = [], start
    while d <= end:
        if d.weekday() < 5 and d not in kr and not _krx_extra_holiday(d):
            days.append(d)
        d += dt.timedelta(days=1)
    if not ascending:
        days.reverse()
    return days


if __name__ == "__main__":
    ds = candidate_trading_days("2026-07-01", "2026-07-24")
    print(f"{len(ds)} 후보 거래일:", [str(x) for x in ds])
    print("holidays lib:", "설치됨" if _KR else "없음(런타임 감지에 의존)")
