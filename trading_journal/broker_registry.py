"""
증권사별 PDF 파서 레지스트리

신규 증권사 추가 방법:
1. parsers/<broker>_futures.py 에 parse_<broker>_futures(pdf_path) -> list[RawFill] 작성
   (RawFill은 nh_futures.RawFill과 동일한 필드를 가져야 함:
    account_no, account_name, trade_date, symbol, side, qty, price,
    fill_time, notional, pnl, fee, broker)
2. 아래 BROKER_REGISTRY에 항목 추가:
   - subject_patterns: 메일 제목에서 이 증권사를 식별할 정규식 리스트
   - parser: 파싱 함수
   - attachment_ext: 첨부파일 확장자 필터 (기본 .pdf)

이렇게 등록해두면 gmail_fetcher가 메일 제목만 보고 자동으로 맞는 파서를
선택해서 호출한다. 기존 파서/매칭 로직은 건드릴 필요 없음.
"""
import re
from dataclasses import dataclass
from typing import Callable

from parsers.nh_futures import parse_nh_futures
from parsers.samsung_futures import parse_samsung_futures


@dataclass
class BrokerConfig:
    name: str
    subject_patterns: list[str]   # 정규식 (메일 제목 매칭용)
    parser: Callable[[str], list]
    gmail_query: str = ""          # Gmail 검색 쿼리 (비어있으면 subject:name 사용)
    attachment_ext: tuple[str, ...] = (".pdf",)


BROKER_REGISTRY: list[BrokerConfig] = [
    BrokerConfig(
        name="NH선물",
        subject_patterns=[
            r"\[NH\s*futures\]",
            r"NH선물.*가정산",
        ],
        gmail_query="from:NHfutures@futures.co.kr",
        parser=parse_nh_futures,
    ),
    BrokerConfig(
        name="삼성선물",
        subject_patterns=[
            r"\[삼성선물\]",
            r"삼성선물.*가정산보고서",
        ],
        gmail_query="from:master@ssfutures.com",
        parser=parse_samsung_futures,
    ),
    # 신규 증권사 추가 시 아래와 같은 형태로 항목만 추가:
    # BrokerConfig(
    #     name="키움선물",
    #     subject_patterns=[r"\[키움선물\]", r"키움.*체결내역"],
    #     parser=parse_kiwoom_futures,
    # ),
]


def match_broker(subject: str) -> BrokerConfig | None:
    """메일 제목으로 증권사 설정을 찾는다. 매칭 실패 시 None."""
    for cfg in BROKER_REGISTRY:
        for pattern in cfg.subject_patterns:
            if re.search(pattern, subject):
                return cfg
    return None


def parse_attachment(broker_cfg: BrokerConfig, pdf_path: str) -> list:
    return broker_cfg.parser(pdf_path)


if __name__ == "__main__":
    test_subjects = [
        "[NH futures] TradeData(20260616,(주)카카오뱅크,304526-11-0001) - 국문가정산(체결시분)",
        "[삼성선물]선물옵션거래 및 예탁자산현황(가정산보고서)",
        "[NH futures] 기타 안내 메일 제목",
        "관련없는 메일",
    ]
    for s in test_subjects:
        cfg = match_broker(s)
        print(f"{'매칭됨: ' + cfg.name if cfg else '매칭안됨':12} <- {s}")
