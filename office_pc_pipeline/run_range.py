# -*- coding: utf-8 -*-
"""
날짜 구간 일괄 실행 — main.py(증적+스냅샷) → daily_review.py(리뷰) 를 평일마다 반복.

사용:
    python run_range.py 2026/07/20 2026/07/24
    python run_range.py 2026/07/20              # 시작일만 주면 그날 하루

주의(OVN 정확성):
    OVN은 ovn_anchor.json 기준일부터 Excel 누적으로 계산된다. 따라서 구간의
    첫날이 정확하려면 anchor일~구간첫날 사이의 모든 평일 체결이 Excel에 있어야 한다.
    안전하게 하려면 anchor를 구간 직전 신뢰일로 맞추거나, anchor일부터 구간을 잡아라.

    daily_review가 크로스체크 불일치를 만나면 중단(exit 1)한다. 그 날은 원인을
    (Excel 증적 누락/anchor 오류) 확인 후, 필요하면 --ktb3/--ktb10로 수동 확정하고
    이어서 다음 날부터 다시 실행하라.
"""
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

THIS_DIR = Path(__file__).parent
PY = sys.executable


def _weekdays(start: datetime, end: datetime):
    d = start
    while d <= end:
        if d.weekday() < 5:  # 월(0)~금(4)
            yield d
        d += timedelta(days=1)


def _run(script: str, date_slug: str) -> int:
    print(f"\n{'#'*64}\n#  {script}  {date_slug}\n{'#'*64}")
    r = subprocess.run([PY, str(THIS_DIR / script), date_slug], cwd=str(THIS_DIR))
    return r.returncode


def main():
    if len(sys.argv) < 2:
        sys.exit("사용: python run_range.py 시작일YYYY/MM/DD [종료일YYYY/MM/DD]")
    start = datetime.strptime(sys.argv[1], "%Y/%m/%d")
    end = datetime.strptime(sys.argv[2], "%Y/%m/%d") if len(sys.argv) > 2 else start
    if end < start:
        sys.exit("종료일이 시작일보다 빠릅니다.")

    days = list(_weekdays(start, end))
    print(f"대상 평일 {len(days)}일: {', '.join(d.strftime('%m/%d') for d in days)}")

    done, failed = [], []
    for d in days:
        slug = d.strftime("%Y/%m/%d")
        if _run("main.py", slug) != 0:
            print(f"[중단] main.py 실패 ({slug}) — 이후 날짜 중단.")
            failed.append(slug)
            break
        if _run("daily_review.py", slug) != 0:
            print(f"[중단] daily_review.py 실패/불일치 ({slug}) — 원인 확인 후 이 날부터 재실행.")
            failed.append(slug)
            break
        done.append(slug)

    print(f"\n{'='*64}")
    print(f"완료: {', '.join(done) or '없음'}")
    if failed:
        print(f"중단 지점: {failed[0]}")
    print(f"{'='*64}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
