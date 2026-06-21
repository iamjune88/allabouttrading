"""
FIFO 기반 진입-청산(라운드트립) 매칭

전제:
- 선물 거래이므로 매수(롱 진입/숏 청산)와 매도(숏 진입/롱 청산)가
  동일 종목 큐에서 먼저 들어온 포지션부터 상계된다.
- 종목별로 시간순 매수큐/매도큐를 운영하며, 반대방향 체결이 들어올 때마다
  먼저 쌓인 보유분부터 소진(FIFO)한다.
- 날짜를 넘어서도(예: 06-16 매수 보유 -> 06-18 청산) 매칭 가능하도록
  종목 단위로 전체 기간 누적 처리한다. (Fills 시트 전체를 대상으로 실행)
- 손익은 Excel에 이미 적재된 분배손익(pnl, 거래소 정산 기준 당일차금)을
  쓰지 않고, FIFO 매칭 시점의 (청산가 - 진입가) * 수량 * 방향부호로
  별도 계산한다. 거래소 당일차금은 보유기간에 걸친 일별 정산손익이라
  진입~청산 구간 손익과 정의가 다르기 때문이다. 손익검증용으로 거래소
  pnl 합계와 별도 비교 가능하도록 둘 다 남긴다.

주의: 가격 단위가 종목마다 다를 수 있어(예: 채권선물 1틱=0.01, 가격*1000000
원/포인트 등) 손익계산 시 틱가치(tick_value)를 종목별로 지정해야 함.
기본값은 1포인트당 1,000,000원(국채선물 3년/10년 통상 단위)으로 둔다.
종목별로 다르면 SYMBOL_TICK_VALUE에 등록해서 덮어쓴다.
"""
from dataclasses import dataclass, field
from collections import deque, defaultdict
from datetime import datetime

# 종목코드 -> 1포인트당 원화가치. 등록 안 된 종목은 DEFAULT_TICK_VALUE 사용.
SYMBOL_TICK_VALUE: dict[str, float] = {}
DEFAULT_TICK_VALUE = 1_000_000  # 국채선물(3년/10년) 표준: 1포인트 = 100만원


@dataclass
class Lot:
    """체결 1건(또는 합산건)을 FIFO 큐에 넣기 위한 단위. 일부만 소진될 수 있음."""
    trade_date: str
    broker: str
    account_no: str
    symbol: str
    side: str          # 매수 / 매도
    qty_remaining: int
    price: float
    fill_time: str
    fee_per_unit: float  # 수수료를 수량으로 나눈 단가 (라운드트립 시 비례배분용)


@dataclass
class RoundTrip:
    trade_date: str          # 청산일 기준
    broker: str
    account_no: str
    symbol: str
    entry_side: str           # 매수(롱) / 매도(숏)
    entry_time: str
    entry_date: str
    entry_price: float
    exit_time: str
    exit_date: str
    exit_price: float
    qty: int
    holding_minutes: float | None
    gross_pnl: float
    fee_total: float
    net_pnl: float
    result: str               # 익절 / 손절 / 손익없음


def _tick_value(symbol: str) -> float:
    return SYMBOL_TICK_VALUE.get(symbol, DEFAULT_TICK_VALUE)


def _dt(date_str: str, time_str: str) -> datetime | None:
    if not time_str or ":" not in time_str:
        return None
    try:
        h, m = time_str.split(":")[:2]
        return datetime.strptime(f"{date_str} {h}:{m}", "%Y-%m-%d %H:%M")
    except ValueError:
        return None


def match_fifo(fills: list) -> list[RoundTrip]:
    """
    fills: aggregate.AggregatedFill 객체 리스트 (시간순 정렬되어 있어야 함).
           만기청산(side='만기')은 청산 매도/매수와 동일하게 반대포지션을 상계하는
           용도로 취급한다 (만기 시 보유분 강제청산).
    반환: 종목/계좌 단위로 매칭된 라운드트립 리스트 (체결시간 순)
    """
    fills_sorted = sorted(
        fills, key=lambda f: (f.trade_date, f.fill_time if f.fill_time != "만기청산" else "99:99")
    )

    # 계좌+종목 단위로 매수큐/매도큐 운영
    long_queues: dict[tuple, deque[Lot]] = defaultdict(deque)
    short_queues: dict[tuple, deque[Lot]] = defaultdict(deque)
    round_trips: list[RoundTrip] = []

    for f in fills_sorted:
        key = (f.account_no, f.symbol)
        fee_per_unit = (f.fee or 0) / f.qty if f.qty else 0
        side = f.side

        if side in ("매수",):
            opposite_q = short_queues[key]
            remaining = f.qty
            entry_time_for_log = f.fill_time

            while remaining > 0 and opposite_q:
                lot = opposite_q[0]
                matched_qty = min(remaining, lot.qty_remaining)

                gross_pnl = (lot.price - f.price) * matched_qty * _tick_value(f.symbol)  # 숏 청산
                fee_total = lot.fee_per_unit * matched_qty + fee_per_unit * matched_qty
                net_pnl = gross_pnl - fee_total

                entry_dt = _dt(lot.trade_date, lot.fill_time)
                exit_dt = _dt(f.trade_date, f.fill_time)
                holding_min = (exit_dt - entry_dt).total_seconds() / 60 if entry_dt and exit_dt else None

                round_trips.append(RoundTrip(
                    trade_date=f.trade_date, broker=f.broker, account_no=f.account_no,
                    symbol=f.symbol, entry_side="매도", entry_time=lot.fill_time,
                    entry_date=lot.trade_date, entry_price=lot.price,
                    exit_time=f.fill_time, exit_date=f.trade_date, exit_price=f.price,
                    qty=matched_qty, holding_minutes=holding_min,
                    gross_pnl=gross_pnl, fee_total=fee_total, net_pnl=net_pnl,
                    result="익절" if net_pnl > 0 else ("손절" if net_pnl < 0 else "손익없음"),
                ))

                lot.qty_remaining -= matched_qty
                remaining -= matched_qty
                if lot.qty_remaining == 0:
                    opposite_q.popleft()

            if remaining > 0:
                long_queues[key].append(Lot(
                    trade_date=f.trade_date, broker=f.broker, account_no=f.account_no,
                    symbol=f.symbol, side="매수", qty_remaining=remaining, price=f.price,
                    fill_time=f.fill_time, fee_per_unit=fee_per_unit,
                ))

        elif side in ("매도", "만기"):
            opposite_q = long_queues[key]
            remaining = f.qty

            while remaining > 0 and opposite_q:
                lot = opposite_q[0]
                matched_qty = min(remaining, lot.qty_remaining)

                gross_pnl = (f.price - lot.price) * matched_qty * _tick_value(f.symbol)  # 롱 청산
                fee_total = lot.fee_per_unit * matched_qty + fee_per_unit * matched_qty
                net_pnl = gross_pnl - fee_total

                entry_dt = _dt(lot.trade_date, lot.fill_time)
                exit_dt = _dt(f.trade_date, f.fill_time) if f.fill_time != "만기청산" else None
                holding_min = (exit_dt - entry_dt).total_seconds() / 60 if entry_dt and exit_dt else None

                round_trips.append(RoundTrip(
                    trade_date=f.trade_date, broker=f.broker, account_no=f.account_no,
                    symbol=f.symbol, entry_side="매수", entry_time=lot.fill_time,
                    entry_date=lot.trade_date, entry_price=lot.price,
                    exit_time=f.fill_time, exit_date=f.trade_date, exit_price=f.price,
                    qty=matched_qty, holding_minutes=holding_min,
                    gross_pnl=gross_pnl, fee_total=fee_total, net_pnl=net_pnl,
                    result="익절" if net_pnl > 0 else ("손절" if net_pnl < 0 else "손익없음"),
                ))

                lot.qty_remaining -= matched_qty
                remaining -= matched_qty
                if lot.qty_remaining == 0:
                    opposite_q.popleft()

            if remaining > 0 and side == "매도":
                short_queues[key].append(Lot(
                    trade_date=f.trade_date, broker=f.broker, account_no=f.account_no,
                    symbol=f.symbol, side="매도", qty_remaining=remaining, price=f.price,
                    fill_time=f.fill_time, fee_per_unit=fee_per_unit,
                ))
            # 만기인데 상계할 보유분이 없는 잔여수량은 데이터 이상이므로 무시(로그만)
            elif remaining > 0 and side == "만기":
                print(f"[경고] 만기청산 잔여수량 미상계: {key} qty={remaining}")

    round_trips.sort(key=lambda r: (r.exit_date, r.exit_time))
    return round_trips


# 미청산(보유중) 포지션 조회용 헬퍼
def open_positions(fills: list) -> dict:
    """현재 시점 기준 종목별 순보유 잔량을 반환 (FIFO 매칭 후 남은 큐)"""
    fills_sorted = sorted(
        fills, key=lambda f: (f.trade_date, f.fill_time if f.fill_time != "만기청산" else "99:99")
    )
    long_queues: dict[tuple, deque[Lot]] = defaultdict(deque)
    short_queues: dict[tuple, deque[Lot]] = defaultdict(deque)

    # match_fifo와 동일 로직을 재실행하지 않고, 간단히 누적 합산으로 순포지션만 계산
    net: dict[tuple, int] = defaultdict(int)
    avg_cost: dict[tuple, list] = defaultdict(list)  # (qty, price) 리스트, 보유분만
    for f in fills_sorted:
        key = (f.account_no, f.symbol)
        sign = 1 if f.side == "매수" else -1
        net[key] += sign * f.qty

    return {k: v for k, v in net.items() if v != 0}


if __name__ == "__main__":
    import sys
    sys.path.insert(0, "/home/claude/trading_journal")
    from parsers.nh_futures import parse_nh_futures
    from parsers.samsung_futures import parse_samsung_futures
    from parsers.aggregate import aggregate_fills

    nh_raws = parse_nh_futures("/mnt/user-data/uploads/20260618_01A103_304526110001_1004.pdf")
    ss_raws = parse_samsung_futures("/mnt/user-data/uploads/20260616100000000000000000014674.pdf")

    all_fills = aggregate_fills(nh_raws) + aggregate_fills(ss_raws)

    print("=== 미청산 포지션 ===")
    for k, v in open_positions(all_fills).items():
        print(f"{k}: 순포지션={v}")

    print("\n=== FIFO 매칭 결과 ===")
    rts = match_fifo(all_fills)
    for r in rts:
        print(f"{r.symbol} {r.entry_side}진입 {r.entry_date}{r.entry_time}@{r.entry_price} -> "
              f"{r.exit_date}{r.exit_time}@{r.exit_price} 수량={r.qty} "
              f"순손익={r.net_pnl:,.0f} ({r.result})")

    print(f"\n총 {len(rts)}건 라운드트립, 순손익합계={sum(r.net_pnl for r in rts):,.0f}")
