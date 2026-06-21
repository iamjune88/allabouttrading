"""
일일 트레이딩 저널 파이프라인 - 메인 실행 스크립트

흐름:
1. Gmail에서 증권사별 체결내역 메일 검색 (broker_registry에 등록된 모든 증권사)
2. PDF 첨부파일 다운로드
3. 증권사별 파서로 체결내역 파싱
4. 분+종목+방향+가격 동일 건 합산
5. Excel(Fills 시트)에 중복없이 누적 적재
6. 전체 누적 Fills 데이터로 FIFO 진입/청산 매칭 -> Trades 시트 갱신
7. (다음 단계에서 TradingView 라벨링/스샷, AI 데일리 저널링 연결 예정)

실행 방법 (로컬 Claude Code 환경):
    python daily_run.py                     # 당일 메일만 처리
    python daily_run.py --after 2026/06/01  # 특정 날짜 이후 메일 일괄 처리(백필용)
    python daily_run.py --dry-run            # Gmail 호출 없이 로컬 PDF 폴더로 테스트

cron 등록 예시 (매일 16:30 실행):
    30 16 * * 1-5 cd /path/to/trading_journal && /usr/bin/python3 daily_run.py >> logs/daily_run.log 2>&1
"""
import argparse
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

from broker_registry import BROKER_REGISTRY, match_broker
from parsers.aggregate import aggregate_fills
from fifo_matcher import match_fifo
from excel_store import append_fills, append_trades

BASE_DIR = Path(__file__).parent
OUTPUT_XLSX = BASE_DIR / "output" / "trading_journal.xlsx"
DOWNLOAD_DIR = BASE_DIR / "data" / "attachments"
LOG_DIR = BASE_DIR / "logs"


def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}")


def run_gmail_pipeline(after_date: str | None) -> list:
    """Gmail에서 모든 등록 증권사 메일을 검색해 PDF를 받고 파싱까지 수행"""
    from gmail_fetcher import (
        get_gmail_service, fetch_attachments_for_subject_patterns,
    )

    service = get_gmail_service()
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    all_raw_fills = []
    for broker_cfg in BROKER_REGISTRY:
        # subject_patterns의 정규식들 중 가장 단순한 키워드로 Gmail 검색 쿼리 구성
        # (Gmail 검색은 정규식을 지원하지 않으므로 subject: 키워드 매칭으로 넓게 가져온 뒤
        #  로컬에서 match_broker로 한 번 더 정확히 필터링)
        query = broker_cfg.gmail_query or f'subject:"{broker_cfg.name}"'
        log(f"[{broker_cfg.name}] Gmail 검색: {query}")

        mails = fetch_attachments_for_subject_patterns(
            service, query, str(DOWNLOAD_DIR), after_date=after_date,
        )
        log(f"[{broker_cfg.name}] 메일 {len(mails)}건 발견")

        for mail in mails:
            # 메일 제목으로 정확한 증권사 재확인 (오탐 방지)
            matched = match_broker(mail["subject"])
            if not matched or matched.name != broker_cfg.name:
                log(f"  스킵(제목 불일치): {mail['subject']}")
                continue

            for pdf_path in mail["saved_paths"]:
                try:
                    raws = matched.parser(pdf_path)
                    log(f"  파싱 성공: {Path(pdf_path).name} ({len(raws)}건)")
                    all_raw_fills.extend(raws)
                except Exception as e:
                    log(f"  [오류] 파싱 실패: {Path(pdf_path).name} - {e}")
                    log(f"  -> 첨부파일이 체결내역 포맷이 아닐 수 있습니다(NH선물은 다양한 "
                        f"메일을 발송하므로 정상적인 경우일 수 있음). 수동 확인 필요.")

    return all_raw_fills


def run_local_dryrun(pdf_dir: str) -> list:
    """Gmail 없이 로컬 PDF 폴더의 파일들을 직접 파싱 (테스트/백필용)"""
    all_raw_fills = []
    for pdf_path in Path(pdf_dir).glob("*.pdf"):
        # 파일명만으로는 증권사를 알 수 없으므로 두 파서를 순서대로 시도
        parsed = False
        for broker_cfg in BROKER_REGISTRY:
            try:
                raws = broker_cfg.parser(str(pdf_path))
                if raws:
                    log(f"[{broker_cfg.name}] 파싱 성공: {pdf_path.name} ({len(raws)}건)")
                    all_raw_fills.extend(raws)
                    parsed = True
                    break
            except Exception:
                continue
        if not parsed:
            log(f"[경고] 어떤 파서로도 인식되지 않음: {pdf_path.name}")
    return all_raw_fills


def main():
    parser = argparse.ArgumentParser(description="일일 트레이딩 저널 파이프라인")
    parser.add_argument("--after", type=str, default=None,
                         help="이 날짜 이후 메일만 처리 (YYYY/MM/DD)")
    parser.add_argument("--dry-run", action="store_true",
                         help="Gmail 연동 없이 data/attachments 폴더의 PDF로 테스트")
    args = parser.parse_args()

    LOG_DIR.mkdir(exist_ok=True)
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

    if args.dry_run:
        log("=== DRY RUN 모드: 로컬 PDF 폴더 사용 ===")
        all_raw_fills = run_local_dryrun(str(DOWNLOAD_DIR))
    else:
        log("=== Gmail 연동 모드 ===")
        all_raw_fills = run_gmail_pipeline(args.after)

    if not all_raw_fills:
        log("신규 체결내역 없음. 종료.")
        return

    aggregated = aggregate_fills(all_raw_fills)
    log(f"합산 완료: 원본 {len(all_raw_fills)}건 -> 합산 {len(aggregated)}건")

    fill_result = append_fills(str(OUTPUT_XLSX), aggregated)
    log(f"Fills 시트 적재: 신규 {fill_result['inserted']}건, "
        f"중복스킵 {fill_result['skipped_dup']}건, 누적총 {fill_result['total_in_sheet']}건")

    # FIFO 매칭은 신규 건만이 아니라 Fills 시트 전체를 다시 읽어 수행해야
    # 날짜를 넘나드는 매칭(예: 06-16 진입 -> 06-18 청산)이 정확해짐
    import pandas as pd
    df_all = pd.read_excel(str(OUTPUT_XLSX), sheet_name="Fills")
    from parsers.aggregate import AggregatedFill
    all_fills_from_sheet = [
        AggregatedFill(
            broker=row["증권사"], account_no=str(row["계좌번호"]), account_name=row["계좌명"],
            trade_date=str(row["거래일자"])[:10], symbol=row["종목"], side=row["방향"],
            qty=int(row["수량"]), price=float(row["가격"]), fill_time=str(row["체결시각"]),
            notional=row["약정금액"], pnl=row["손익"], fee=row["수수료"],
            raw_count=int(row["원본체결건수"]),
        )
        for _, row in df_all.iterrows()
    ]

    round_trips = match_fifo(all_fills_from_sheet)
    trade_result = append_trades(str(OUTPUT_XLSX), round_trips)
    log(f"Trades 시트 적재: 신규 {trade_result['inserted']}건, "
        f"중복스킵 {trade_result['skipped_dup']}건, 누적총 {trade_result['total_in_sheet']}건")

    log(f"완료. 결과 파일: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
