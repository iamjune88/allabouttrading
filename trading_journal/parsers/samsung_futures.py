"""
삼성선물 '선물옵션거래 및 예탁자산현황(가정산보고서)' PDF 파서

대상 메일: [삼성선물]선물옵션거래 및 예탁자산현황(가정산보고서)

핵심 로직:
- [ 거래내역 ] 섹션은 (종목, 방향, 당일체결가) 단위로 이미 손익(당일차금)/
  수수료/약정금액이 합산되어 있음. 이건 시간 정보가 없는 "가격대별 합계"임.
- [ 체결내역 ] 섹션은 (종목, 방향, 가격, 초단위시간, 체결번호) 단위의 개별
  체결 기록이며 손익/수수료가 없음.
- 두 섹션을 (종목, 방향, 가격) 키로 매칭한다.
  같은 키에 속한 체결내역 행들의 수량 합 == 거래내역 개별행의 수량과 일치해야
  정상이며, 그 경우 거래내역의 손익/수수료/약정금액을 수량 비례로 배분한다.
  (가격이 같으므로 사실상 1건으로 합쳐지는 경우가 대부분이며, 요청된
   "시간(분)+종목+방향+가격 일치 시 합산" 기준과 자연히 맞아떨어진다)
- 만기청산(만기) 행은 체결내역에 대응 건이 없으므로 별도 RawFill로 추가한다.
"""
import re
from dataclasses import dataclass
from collections import defaultdict


@dataclass
class RawFill:
    account_no: str
    account_name: str
    trade_date: str        # YYYY-MM-DD
    symbol: str
    side: str                # 매수 / 매도 / 만기
    qty: int
    price: float
    fill_time: str           # HH:MM
    fill_time_sec: str | None  # HH:MM:SS 원본 (만기 행은 None)
    fill_no: str | None
    notional: int | None = None
    pnl: int | None = None
    fee: int | None = None
    broker: str = "삼성선물"


def _pdftotext_layout(pdf_path: str) -> str:
    import pdfplumber
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(layout=True, x_tolerance=3, y_tolerance=3)
            if text:
                pages.append(text)
    return "\n".join(pages)


def _to_int(s: str) -> int:
    return int(s.replace(",", "").replace("+", "").strip())


def _to_float(s: str) -> float:
    return float(s.replace(",", "").strip())


HEADER_RE = re.compile(r"거래일자\s*:\s*(\d{4})/(\d{2})/(\d{2})")
ACCOUNT_RE = re.compile(r"위탁계좌번호\s*:\s*(\S+)\s+(\S+)")

FILL_ROW_RE = re.compile(
    r"^\s*(A\d+)\s+(매수|매도)\s+(\d+)\s+([\d.]+)\s+\S+\s+\S+\s+(\d{2}:\d{2}:\d{2})\s+(\d+)\s*$"
)

# 거래내역 개별행(만기 포함): [주간] 종목 구분 수량 당일체결가 정산가 약정금액 수수료 당일차금
DETAIL_ROW_RE = re.compile(
    r"^\s*(?:주간\s+)?(A\d+)\s+(매수|매도|만기)\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+"
    r"([\d,]+)\s+([\d,]+)\s+(-?[\d,]+)\s*$"
)


def _parse_header(text: str) -> tuple[str, str, str]:
    m = HEADER_RE.search(text)
    a = ACCOUNT_RE.search(text)
    if not m or not a:
        raise ValueError("거래일자/계좌정보를 찾을 수 없습니다")
    trade_date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return trade_date, a.group(1), a.group(2)


def _parse_trade_details(text: str) -> list[dict]:
    """[ 거래내역 ] 섹션의 개별행(소계 제외)을 모두 추출"""
    lines = text.split("\n")
    in_section = False
    details = []
    for line in lines:
        if "[ 거래내역 ]" in line:
            in_section = True
            continue
        if in_section and line.strip().startswith("[") and "거래내역" not in line:
            break
        if not in_section:
            continue
        m = DETAIL_ROW_RE.match(line)
        if m:
            symbol, side, qty, price, settle_price, notional, fee, daily_diff = m.groups()
            details.append({
                "symbol": symbol, "side": side, "qty": _to_int(qty),
                "price": _to_float(price), "notional": _to_int(notional),
                "fee": _to_int(fee), "daily_diff": _to_int(daily_diff),
            })
    return details


def parse_samsung_futures(pdf_path: str) -> list[RawFill]:
    text = _pdftotext_layout(pdf_path)
    trade_date, account_no, account_name = _parse_header(text)

    # 1) 체결내역(개별 체결, 손익 없음) 파싱
    lines = text.split("\n")
    in_section = False
    raw_fills: list[RawFill] = []
    for line in lines:
        if "[ 체결내역 ]" in line:
            in_section = True
            continue
        if in_section and line.strip().startswith("[") and "체결내역" not in line:
            in_section = False
            continue
        if not in_section:
            continue
        m = FILL_ROW_RE.match(line)
        if m:
            symbol, side, qty, price, ftime_sec, fill_no = m.groups()
            raw_fills.append(RawFill(
                account_no=account_no, account_name=account_name, trade_date=trade_date,
                symbol=symbol, side=side, qty=_to_int(qty), price=_to_float(price),
                fill_time=ftime_sec[:5], fill_time_sec=ftime_sec, fill_no=fill_no,
            ))

    # 2) 거래내역 개별행(가격대별 손익/수수료) 파싱
    details = _parse_trade_details(text)

    # 3) (symbol, side, price) 키로 체결내역 그룹핑 후 거래내역과 매칭
    fills_by_key: dict[tuple, list[RawFill]] = defaultdict(list)
    for f in raw_fills:
        fills_by_key[(f.symbol, f.side, f.price)].append(f)

    matched_fill_keys = set()
    for d in details:
        if d["side"] == "만기":
            raw_fills.append(RawFill(
                account_no=account_no, account_name=account_name, trade_date=trade_date,
                symbol=d["symbol"], side="만기", qty=d["qty"], price=d["price"],
                fill_time="만기청산", fill_time_sec=None, fill_no=None,
                notional=d["notional"], pnl=d["daily_diff"], fee=d["fee"],
            ))
            continue

        key = (d["symbol"], d["side"], d["price"])
        group = fills_by_key.get(key, [])
        if not group:
            print(f"[경고] 거래내역의 ({key}) 가격대에 대응하는 체결내역이 없습니다")
            continue
        matched_fill_keys.add(key)
        group_qty = sum(it.qty for it in group)
        if group_qty != d["qty"]:
            print(f"[경고] {key} 체결내역 합계수량({group_qty}) != 거래내역수량({d['qty']})")
        for it in group:
            ratio = it.qty / d["qty"] if d["qty"] else 0
            it.notional = round(d["notional"] * ratio)
            it.fee = round(d["fee"] * ratio)
            it.pnl = round(d["daily_diff"] * ratio)

    for key in fills_by_key:
        if key not in matched_fill_keys:
            print(f"[경고] 체결내역의 ({key})에 대응하는 거래내역 소계가 없습니다 "
                  f"(손익/수수료 미배정)")

    return raw_fills


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else \
        "/mnt/user-data/uploads/20260616100000000000000000014674.pdf"
    fills = parse_samsung_futures(path)
    for f in fills:
        t = f.fill_time_sec or f.fill_time
        print(f"{t} {f.symbol} {f.side} 수량={f.qty:>3} 가격={f.price:>7} "
              f"손익={f.pnl:>10,} 수수료={f.fee:>7,} 번호={f.fill_no}")
    print(f"\n총 {len(fills)}건, 합계수량={sum(f.qty for f in fills)}, "
          f"합계손익={sum(f.pnl or 0 for f in fills):,}, 합계수수료={sum(f.fee or 0 for f in fills):,}")
