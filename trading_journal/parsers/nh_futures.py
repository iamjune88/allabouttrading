"""
NH선물 '선물 거래 확인서' PDF 파서

대상 메일: [NH futures] TradeData(YYYYMMDD,계좌명,계좌번호) - 국문가정산(체결시분)
대상 섹션: [당일 선물거래]

구조 특이사항:
- 종목/정산가/거래유형/B/S는 해당 종목 블록의 첫 줄에만 표기되고,
  이후 줄은 빈 칸으로 이어짐(같은 종목 내 가격대별 분할체결)
- 체결시간은 시:분까지만 존재(초 단위 없음)
- 손익/수수료가 이미 행 단위로 존재 -> 합산 시 그대로 더하면 됨
"""
import re
from dataclasses import dataclass
from pathlib import Path


@dataclass
class RawFill:
    account_no: str
    account_name: str
    trade_date: str       # YYYY-MM-DD
    symbol: str
    side: str              # 매수 / 매도
    qty: int
    price: float
    fill_time: str          # HH:MM
    notional: int
    pnl: int
    fee: int
    broker: str = "NH선물"


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


HEADER_RE = re.compile(r"거래일자\s*:\s*(\d{4}-\d{2}-\d{2})")
ACCOUNT_RE = re.compile(r"계좌번호\s*:\s*([\d\-]+)\s+계좌명\s*:\s*(\S+)")

# 종목 시작 줄: 종목코드 정산가 거래유형(체결) B/S 수량 가격 시간 약정금액 손익 수수료
ROW_WITH_SYMBOL_RE = re.compile(
    r"^\s*(A\d+)\s+([\d.]+)\s+체결\s+(매수|매도)\s+(\d+)\s+([\d.]+)\s+"
    r"(\d{2}:\d{2})\s+([\d,]+)\s+(-?[\d,]+)\s+([\d,]+)\s*$"
)

# 종목 없이 이어지는 줄 (같은 종목, 같은 방향 계속) :
#   수량 가격 시간 약정금액 손익 수수료
ROW_CONTINUATION_RE = re.compile(
    r"^\s{20,}(\d+)\s+([\d.]+)\s+(\d{2}:\d{2})\s+([\d,]+)\s+(-?[\d,]+)\s+([\d,]+)\s*$"
)

# 매도 전환 줄 (종목 없이, B/S만 새로 등장) :
#   매도 수량 가격 시간 약정금액 손익 수수료
ROW_SIDE_SWITCH_RE = re.compile(
    r"^\s*(매수|매도)\s+(\d+)\s+([\d.]+)\s+(\d{2}:\d{2})\s+([\d,]+)\s+(-?[\d,]+)\s+([\d,]+)\s*$"
)


def parse_nh_futures(pdf_path: str) -> list[RawFill]:
    text = _pdftotext_layout(pdf_path)
    lines = text.split("\n")

    trade_date_m = HEADER_RE.search(text)
    account_m = ACCOUNT_RE.search(text)
    if not trade_date_m or not account_m:
        raise ValueError(f"거래일자/계좌정보를 찾을 수 없습니다: {pdf_path}")

    trade_date = trade_date_m.group(1)
    account_no = account_m.group(1)
    account_name = account_m.group(2)

    in_section = False
    current_symbol = None
    fills: list[RawFill] = []

    for line in lines:
        if "[당일 선물거래]" in line:
            in_section = True
            continue
        if in_section and line.strip().startswith("["):
            # 다음 섹션 진입 -> 종료
            break
        if not in_section:
            continue
        if not line.strip() or "총계" in line or "종목" in line and "정산가" in line:
            continue

        m = ROW_WITH_SYMBOL_RE.match(line)
        if m:
            symbol, _settle, side, qty, price, ftime, notional, pnl, fee = m.groups()
            current_symbol = symbol
            current_side = side
            fills.append(RawFill(
                account_no=account_no, account_name=account_name, trade_date=trade_date,
                symbol=symbol, side=side, qty=_to_int(qty), price=_to_float(price),
                fill_time=ftime, notional=_to_int(notional), pnl=_to_int(pnl), fee=_to_int(fee),
            ))
            continue

        m = ROW_SIDE_SWITCH_RE.match(line)
        if m and current_symbol:
            side, qty, price, ftime, notional, pnl, fee = m.groups()
            current_side = side
            fills.append(RawFill(
                account_no=account_no, account_name=account_name, trade_date=trade_date,
                symbol=current_symbol, side=side, qty=_to_int(qty), price=_to_float(price),
                fill_time=ftime, notional=_to_int(notional), pnl=_to_int(pnl), fee=_to_int(fee),
            ))
            continue

        m = ROW_CONTINUATION_RE.match(line)
        if m and current_symbol:
            qty, price, ftime, notional, pnl, fee = m.groups()
            fills.append(RawFill(
                account_no=account_no, account_name=account_name, trade_date=trade_date,
                symbol=current_symbol, side=current_side, qty=_to_int(qty), price=_to_float(price),
                fill_time=ftime, notional=_to_int(notional), pnl=_to_int(pnl), fee=_to_int(fee),
            ))
            continue

    return fills


if __name__ == "__main__":
    import sys
    fills = parse_nh_futures(sys.argv[1] if len(sys.argv) > 1 else
                              "/mnt/user-data/uploads/20260618_01A103_304526110001_1004.pdf")
    for f in fills:
        print(f)
    print(f"\n총 {len(fills)}건, 합계수량={sum(f.qty for f in fills)}, "
          f"합계손익={sum(f.pnl for f in fills):,}, 합계수수료={sum(f.fee for f in fills):,}")
