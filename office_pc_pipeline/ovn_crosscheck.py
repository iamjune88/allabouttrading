# -*- coding: utf-8 -*-
"""
오버나잇 포지션 교차검증 (Excel 체결누적 ↔ 브로커 PDF 미결제약정).

파이프라인 신뢰성의 핵심. Excel 누적합만으로는 체결 누락·중복·오파싱을 못 잡지만,
브로커가 확정한 미결제약정(전일잔고=오늘 진입 OVN)은 체결과 독립적인 수치라
둘을 비교하면 Excel 증적의 오류를 잡아낼 수 있다.

데이터 흐름:
  main.py    → PDF 파싱 → build_ovn_snapshot() → ovn_snapshots/{date}.json 저장
  daily_review → load_snapshot() + ovn_tracker.compute_ovn() → crosscheck()

부호 규약: 매수 = +(롱), 매도 = -(숏).
종목 분류: A65XX = KTB3, A67XX = KTB10.
"""
import json
from pathlib import Path

THIS_DIR = Path(__file__).parent
SNAPSHOT_DIR = THIS_DIR / "ovn_snapshots"
MULT = 1_000_000  # 1pt = 100만원/계약


def _leg(code: str):
    """종목코드 → 'ktb3' | 'ktb10' | None"""
    c = str(code)
    if c.startswith("A65"):
        return "ktb3"
    if c.startswith("A67"):
        return "ktb10"
    return None


def _signed(side: str, qty) -> int:
    """구분(매수/매도) + 수량 → 부호 있는 계약수."""
    try:
        q = int(round(float(qty)))
    except (TypeError, ValueError):
        return 0
    if "매수" in str(side):
        return q
    if "매도" in str(side):
        return -q
    return 0


def build_ovn_snapshot(parsed_list: list, date_str: str) -> dict:
    """
    parse_nh()/parse_ss() 결과 목록 → 브로커 확정 OVN 스냅샷.

    반환 dict:
      entry/close: {ktb3, ktb10}   진입(전일잔고 합)/종료(잔고 합) 오버나잇
      renewal_pdf: {ktb3, ktb10}   PDF 갱신차금 직접합 (SS만 제공)
      renewal_calc:{ktb3, ktb10}   정산가차 계산 (전일잔고 × Δ정산가 × 승수)
      settlement:  {ktb3:{prev,today}, ...}  정산가 (검증·차트용)
    """
    entry = {"ktb3": 0, "ktb10": 0}
    close = {"ktb3": 0, "ktb10": 0}
    renewal_pdf = {"ktb3": 0, "ktb10": 0}
    # 정산가차 검산: SS 미결 행별로 (전일잔고 × Δ정산가 × 승수)를 합산한다.
    #   갱신차금은 브로커·정산가별로 다르므로 진입 OVN 합계에 단일 Δ를 곱하면 틀린다.
    #   NH는 전일정산가를 제공하지 않으므로 이 검산은 SS 포지션에 대해서만 성립.
    renewal_calc = {"ktb3": 0, "ktb10": 0}
    settle = {"ktb3": {"prev": None, "today": None},
              "ktb10": {"prev": None, "today": None}}
    brokers = {}

    for parsed in parsed_list:
        src = parsed.get("source", "")
        for m in parsed.get("미결", []):
            leg = _leg(m.get("종목"))
            if not leg:
                continue
            if src == "SS선물":
                # SS: 각 행에 잔고/전일잔고/갱신차금/정산가가 모두 존재
                prev_qty = _signed(m.get("구분"), m.get("전일잔고"))
                entry[leg] += prev_qty
                close[leg] += _signed(m.get("구분"), m.get("잔고"))
                gv = m.get("갱신차금")
                if isinstance(gv, (int, float)):
                    renewal_pdf[leg] += int(gv)
                p = m.get("전일정산가"); t = m.get("당일정산가")
                if isinstance(p, (int, float)) and isinstance(t, (int, float)):
                    # 행별 전일잔고 × 그 행의 정산가차 (부호 있는 전일잔고 사용)
                    renewal_calc[leg] += int(round(prev_qty * (t - p) * MULT))
                    if settle[leg]["prev"] is None:
                        settle[leg]["prev"] = float(p)
                    if settle[leg]["today"] is None:
                        settle[leg]["today"] = float(t)
            elif src == "NH선물":
                # NH: '전일' 구분 행 = 진입 OVN, '당일' 구분 행 = 종료 OVN
                signed = _signed(m.get("BS"), m.get("잔량"))
                if m.get("구분") == "전일":
                    entry[leg] += signed
                    # NH 전일미결제의 정산가는 당일정산가. 전일정산가 미제공 →
                    # NH분은 정산가차 검산에서 제외(갱신차금은 PDF값을 신뢰).
                    if settle[leg]["today"] is None and isinstance(m.get("정산가"), (int, float)):
                        settle[leg]["today"] = float(m["정산가"])
                elif m.get("구분") == "당일":
                    close[leg] += signed
                    if settle[leg]["today"] is None and isinstance(m.get("정산가"), (int, float)):
                        settle[leg]["today"] = float(m["정산가"])
        brokers[src] = True

    return {
        "date": date_str,
        "brokers": sorted(brokers),
        "entry": entry,
        "close": close,
        "renewal_pdf": renewal_pdf,
        "renewal_calc": renewal_calc,
        "settlement": settle,
    }


def save_snapshot(snapshot: dict):
    """날짜별 스냅샷 저장. 재실행 시 해당 날짜만 덮어쓰므로 다른 날짜에 영향 없음."""
    SNAPSHOT_DIR.mkdir(exist_ok=True)
    date_slug = snapshot["date"].replace("-", "").replace("/", "")
    path = SNAPSHOT_DIR / f"{date_slug}.json"
    path.write_text(json.dumps(snapshot, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def load_snapshot(date_str: str):
    """날짜(YYYY-MM-DD 또는 YYYY/MM/DD)에 해당하는 스냅샷 로드. 없으면 None."""
    date_slug = date_str.replace("-", "").replace("/", "")
    path = SNAPSHOT_DIR / f"{date_slug}.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


class CrosscheckError(Exception):
    """Excel 누적 OVN과 브로커 미결제약정이 불일치. daily_review를 중단시킨다."""


def crosscheck(excel_ovn: dict, snapshot: dict, renewal_tol: int = 1):
    """
    excel_ovn: ovn_tracker.compute_ovn() 결과 (ovn_ktb3/ovn_ktb10 = 진입 OVN)
    snapshot:  build_ovn_snapshot() 결과 (브로커 확정)

    진입 OVN이 다르면 CrosscheckError. 갱신차금 두 방식 차이는 경고로만 보고.
    반환: {ok, mismatches[], warnings[], renewal{ktb3,ktb10}, renewal_source}
    """
    mismatches, warnings = [], []

    exc = {"ktb3": int(excel_ovn.get("ovn_ktb3", 0)),
           "ktb10": int(excel_ovn.get("ovn_ktb10", 0))}
    brk = snapshot["entry"]

    for leg in ("ktb3", "ktb10"):
        if exc[leg] != brk[leg]:
            mismatches.append(
                f"{leg.upper()} 진입 OVN 불일치: Excel누적={exc[leg]:+d}  "
                f"브로커미결제={brk[leg]:+d} (차이 {exc[leg]-brk[leg]:+d})")

    # 갱신차금 교차검증: PDF 직접합 vs 정산가차 계산
    renewal = {}
    for leg in ("ktb3", "ktb10"):
        pdf_v = snapshot["renewal_pdf"].get(leg, 0)
        calc_v = snapshot["renewal_calc"].get(leg, 0)
        if pdf_v and calc_v and abs(pdf_v - calc_v) > renewal_tol:
            warnings.append(
                f"{leg.upper()} 갱신차금 방식차: PDF={pdf_v:+,}  정산가차={calc_v:+,} "
                f"(차 {pdf_v-calc_v:+,})")
        # PDF값이 있으면 우선(브로커 확정), 없으면 정산가차 계산값
        renewal[leg] = pdf_v if pdf_v else calc_v

    renewal_source = "PDF갱신차금" if any(snapshot["renewal_pdf"].values()) else "정산가차계산"

    return {
        "ok": not mismatches,
        "mismatches": mismatches,
        "warnings": warnings,
        "renewal": renewal,
        "renewal_source": renewal_source,
    }
