# -*- coding: utf-8 -*-
"""
선물거래 증적 메인 실행 스크립트
- Gmail에서 PDF 수신 → 파싱 → Excel 저장
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
    """
    target_date: "YYYY/MM/DD" — 없으면 오늘 날짜
    """
    from gmail_fetcher import fetch_pdfs
    from pdf_parser import parse_nh, parse_ss
    from excel_writer import save_to_excel

    if not target_date:
        target_date = date.today().strftime("%Y/%m/%d")

    log(f"=== 실행 시작 | 대상일: {target_date} ===")

    # 1. Gmail에서 PDF 다운로드
    try:
        downloads = fetch_pdfs(target_date)
    except Exception as e:
        log(f"[오류] Gmail 다운로드 실패: {e}")
        log(traceback.format_exc())
        return

    if not downloads:
        log("수신된 PDF 없음 — 종료")
        return

    # 2. PDF 파싱
    parsed_list = []
    for item in downloads:
        source = item["source"]
        path = item["path"]
        date_str = item["date"]

        try:
            if source == "NH선물":
                parsed = parse_nh(str(path))
            elif source == "SS선물":
                parsed = parse_ss(str(path))
            else:
                log(f"[경고] 알 수 없는 출처: {source}")
                continue

            if not parsed.get("date"):
                parsed["date"] = date_str

            log(f"  [{source}] 파싱 완료 — 체결 {len(parsed.get('체결', []))}건")
            parsed_list.append(parsed)

        except Exception as e:
            log(f"[오류] {source} PDF 파싱 실패: {path.name} — {e}")
            log(traceback.format_exc())

    if not parsed_list:
        log("파싱된 데이터 없음 — 종료")
        return

    # 3. Excel 저장 — Gmail "after:" 검색은 여러 날짜를 한번에 반환하므로,
    #    각 PDF의 실제 거래일(parsed["date"])별로 묶어서 해당 날짜 시트에만 저장한다.
    by_date: dict[str, list] = {}
    for parsed in parsed_list:
        by_date.setdefault(parsed["date"], []).append(parsed)

    for d, items in sorted(by_date.items()):
        try:
            excel_path = save_to_excel(items, d)
            log(f"[완료] Excel 저장 ({d}): {excel_path}")
        except Exception as e:
            log(f"[오류] Excel 저장 실패 ({d}): {e}")
            log(traceback.format_exc())

    log("=== 실행 완료 ===\n")


if __name__ == "__main__":
    # 인수로 날짜 지정 가능: python main.py 2026/06/16
    target = sys.argv[1] if len(sys.argv) > 1 else None
    run(target)
