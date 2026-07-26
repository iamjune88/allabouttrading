# -*- coding: utf-8 -*-
"""
선물거래 증적 메인 실행 스크립트
- Gmail에서 자료 수신 → 파싱·병합 → Excel 저장 + OVN 스냅샷 저장
- Windows 작업 스케줄러에서 매일 실행
"""
import sys
import traceback
from datetime import date, datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent
LOG_FILE = BASE_DIR / "실행로그.txt"


def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run(target_date: str = None):
    """target_date: "YYYY/MM/DD" — 없으면 오늘 날짜"""
    from gmail_fetcher import fetch_pdfs
    from parse_router import parse_and_merge
    from excel_writer import save_to_excel

    if not target_date:
        target_date = date.today().strftime("%Y/%m/%d")

    log(f"=== 실행 시작 | 대상일: {target_date} ===")

    # 1. Gmail에서 자료 다운로드 (제목으로 특정된 PDF/CSV)
    try:
        downloads = fetch_pdfs(target_date)
    except Exception as e:
        log(f"[오류] Gmail 다운로드 실패: {e}")
        log(traceback.format_exc())
        return

    if not downloads:
        log("수신된 자료 없음 — 종료")
        return

    # 2. 파싱 + 병합 — 종목사·날짜별로 체결(CSV 우선)과 미결(PDF/CSV)을 하나로 합친다.
    parsed_list = parse_and_merge(downloads, log=log)

    if not parsed_list:
        log("파싱된 데이터 없음 — 종료")
        return

    # 3. 날짜별로 묶어서 Excel 저장 + OVN 스냅샷 저장
    by_date: dict[str, list] = {}
    for parsed in parsed_list:
        by_date.setdefault(parsed["date"], []).append(parsed)

    from ovn_crosscheck import build_ovn_snapshot, save_snapshot

    for d, items in sorted(by_date.items()):
        try:
            excel_path = save_to_excel(items, d)
            log(f"[완료] Excel 저장 ({d}): {excel_path}")
        except Exception as e:
            log(f"[오류] Excel 저장 실패 ({d}): {e}")
            log(traceback.format_exc())

        # 브로커 확정 미결제약정 스냅샷 저장 (daily_review 크로스체크용).
        # 날짜별 파일이라 재실행해도 해당 날짜만 갱신 — positions.json 체인처럼 오염되지 않는다.
        try:
            snap = build_ovn_snapshot(items, d)
            path = save_snapshot(snap)
            log(f"[완료] OVN 스냅샷 저장 ({d}): 진입 KTB3 {snap['entry']['ktb3']:+d}/"
                f"KTB10 {snap['entry']['ktb10']:+d}  → {path.name}")
        except Exception as e:
            log(f"[경고] OVN 스냅샷 저장 실패 ({d}): {e}")
            log(traceback.format_exc())

    log("=== 실행 완료 ===\n")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else None
    run(target)
