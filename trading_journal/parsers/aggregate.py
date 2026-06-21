"""
체결내역 합산(썸업) 공통 로직

기준: 같은 분(HH:MM) + 같은 종목 + 같은 방향(B/S) + 같은 체결가
      위 4개가 모두 일치하는 행들을 하나로 합산한다.

합산 방식:
- 수량: 합산(sum)
- 가격: 동일하므로 그대로 사용
- 손익/수수료: 존재하면 합산(sum), 없으면 None 유지
- 약정금액: 존재하면 합산(sum)
- 체결시간: 그룹 내 첫 시각 유지 (분 단위로는 어차피 동일)
- 원본 체결건수: 합산된 원본 행 수를 별도 필드로 남겨 추적 가능하게 함
"""
from dataclasses import dataclass, field
from collections import defaultdict


@dataclass
class AggregatedFill:
    broker: str
    account_no: str
    account_name: str
    trade_date: str
    symbol: str
    side: str
    qty: int
    price: float
    fill_time: str          # HH:MM (그룹 내 첫 시각)
    notional: int | None
    pnl: int | None
    fee: int | None
    raw_count: int = 1       # 합산 전 원본 행 개수


def aggregate_fills(raw_fills: list) -> list[AggregatedFill]:
    """
    raw_fills: nh_futures.RawFill 또는 samsung_futures.RawFill 객체 리스트
               (broker, account_no, account_name, trade_date, symbol, side,
                qty, price, fill_time, notional, pnl, fee 속성을 가져야 함)
    그룹 키: (trade_date, account_no, symbol, side, fill_time[:5], price)
    """
    groups: dict[tuple, list] = defaultdict(list)

    for f in raw_fills:
        minute_key = f.fill_time[:5]  # HH:MM (혹시 HH:MM:SS로 들어와도 분단위로 자름)
        key = (f.trade_date, f.account_no, f.symbol, f.side, minute_key, f.price)
        groups[key].append(f)

    aggregated: list[AggregatedFill] = []
    for key, items in groups.items():
        trade_date, account_no, symbol, side, minute_key, price = key
        first = items[0]

        def _sum_or_none(attr):
            vals = [getattr(it, attr, None) for it in items]
            if all(v is None for v in vals):
                return None
            return sum(v for v in vals if v is not None)

        aggregated.append(AggregatedFill(
            broker=first.broker,
            account_no=account_no,
            account_name=first.account_name,
            trade_date=trade_date,
            symbol=symbol,
            side=side,
            qty=sum(it.qty for it in items),
            price=price,
            fill_time=minute_key,
            notional=_sum_or_none("notional"),
            pnl=_sum_or_none("pnl"),
            fee=_sum_or_none("fee"),
            raw_count=len(items),
        ))

    # 정렬: 날짜 -> 시간 -> 종목 -> 방향
    aggregated.sort(key=lambda a: (a.trade_date, a.fill_time, a.symbol, a.side))
    return aggregated


if __name__ == "__main__":
    import sys
    sys.path.insert(0, "/home/claude/trading_journal")
    from parsers.nh_futures import parse_nh_futures

    raws = parse_nh_futures("/mnt/user-data/uploads/20260618_01A103_304526110001_1004.pdf")
    aggs = aggregate_fills(raws)
    print(f"원본 {len(raws)}건 -> 합산 후 {len(aggs)}건\n")
    for a in aggs:
        print(f"{a.fill_time} {a.symbol} {a.side} 수량={a.qty:>4} 가격={a.price} "
              f"손익={a.pnl:>12,} 수수료={a.fee:>8,} (원본{a.raw_count}건 합산)")

    print(f"\n검증 - 합산후 합계수량={sum(a.qty for a in aggs)}, "
          f"합계손익={sum(a.pnl for a in aggs):,}, 합계수수료={sum(a.fee for a in aggs):,}")
