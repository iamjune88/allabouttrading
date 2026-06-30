# -*- coding: utf-8 -*-
"""
NH선물 / SS선물(삼성선물) PDF 파싱 모듈
- NH선물:  [당일 선물거래] 섹션
- SS선물:  [ 체결내역 ] 섹션
- 두 회사 모두 멀티페이지 지원
"""
import re
import pdfplumber


def _num(val):
    """콤마 제거 후 숫자 변환. 실패 시 원본 반환."""
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return val


def _all_lines(pdf_path: str) -> list[str]:
    """PDF 전 페이지 텍스트를 줄 단위 리스트로 반환."""
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend(text.split("\n"))
    return [l.strip() for l in lines if l.strip()]


# ─────────────────────────────────────────────
# NH선물 파서
# ─────────────────────────────────────────────
def parse_nh(pdf_path: str) -> dict:
    """
    NH선물 거래확인서 파싱
    대상 섹션: [당일 선물거래]
    컬럼: 종목 현재가 거래구분 B/S 수량 가격 체결시간 거래금액 손익 수수료
    """
    result = {
        "source": "NH선물",
        "date": "",
        "account": "",
        "체결": [],
        "미결": [],
        "요약": {},
    }

    lines = _all_lines(pdf_path)

    # 날짜 / 계좌번호
    for line in lines:
        m = re.search(r"거래일자\s*[:：]\s*(\d{4}-\d{2}-\d{2})", line)
        if m:
            result["date"] = m.group(1)
        m = re.search(r"계좌번호\s*[:：]\s*([\d\-]+)", line)
        if m:
            result["account"] = m.group(1)

    # 섹션 키워드 — 01A103: [당일 선물거래], 01A101/02A101: [당일선물매매]
    SEC_TRADE_KEYWORDS = ("당일 선물거래", "당일선물거래", "당일선물매매", "당일 선물매매")
    SEC_OPEN_D  = "당일미결제"
    SEC_OPEN_P  = "전일미결제"

    CODE_RE  = re.compile(r"^[A-Z]\d{6,8}")   # NH: A6569000
    BS_RE    = re.compile(r"^(매수|매도)")     # B/S로 시작하는 후속행
    NUM_RE   = re.compile(r"^\d")             # 숫자로 시작하는 후속행 (B/S 생략됨)
    TIME_RE  = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")  # 체결시간 (01A103에만 존재)
    SKIP     = {"합계", "소계", "소 계", "합 계", "총계"}

    section      = None
    current_code = None
    current_현재가 = ""
    current_bs   = ""   # 마지막 B/S (생략된 후속행에서 이어서 사용)

    for line in lines:
        # 섹션 감지
        if "[" in line and "]" in line:
            if any(kw in line for kw in SEC_TRADE_KEYWORDS):
                section = "체결"
                current_code = None
            elif SEC_OPEN_D in line:
                section = "미결_당일"
                current_code = None
            elif SEC_OPEN_P in line:
                section = "미결_전일"
                current_code = None
            else:
                section = None
                current_code = None
            continue

        if section is None:
            continue

        if any(k in line for k in SKIP):
            continue

        parts = line.split()
        if not parts:
            continue

        if section == "체결":
            # ── 케이스 A: 종목코드로 시작 ──
            # 01A101: 종목 현재가 체결 B/S 수량 가격 거래금액 손익 수수료 (9개)
            # 01A103: 종목 현재가 체결 B/S 수량 가격 체결시간 거래금액 손익 수수료 (10개)
            if CODE_RE.match(line) and len(parts) >= 7:
                current_code  = parts[0]
                current_현재가 = parts[1]
                current_bs    = parts[3]
                has_time = len(parts) >= 8 and bool(TIME_RE.match(parts[6]))
                off = 1 if has_time else 0
                try:
                    result["체결"].append({
                        "종목":     current_code,
                        "현재가":   _num(current_현재가),
                        "BS":       current_bs,
                        "수량":     _num(parts[4]),
                        "가격":     _num(parts[5]),
                        "체결시간": parts[6] if has_time else "",
                        "거래금액": _num(parts[6 + off]),
                        "손익":     _num(parts[7 + off]) if len(parts) > 7 + off else "",
                        "수수료":   _num(parts[8 + off]) if len(parts) > 8 + off else "",
                    })
                except Exception:
                    pass

            # ── 케이스 B: 매수/매도로 시작 (B/S 변경된 후속행) ──
            # 01A101: B/S 수량 가격 거래금액 손익 수수료 (6개) / 01A103: + 체결시간 (7개)
            elif BS_RE.match(line) and current_code and len(parts) >= 4:
                current_bs = parts[0]
                has_time = len(parts) >= 5 and bool(TIME_RE.match(parts[3]))
                off = 1 if has_time else 0
                try:
                    result["체결"].append({
                        "종목":     current_code,
                        "현재가":   _num(current_현재가),
                        "BS":       current_bs,
                        "수량":     _num(parts[1]),
                        "가격":     _num(parts[2]),
                        "체결시간": parts[3] if has_time else "",
                        "거래금액": _num(parts[3 + off]),
                        "손익":     _num(parts[4 + off]) if len(parts) > 4 + off else "",
                        "수수료":   _num(parts[5 + off]) if len(parts) > 5 + off else "",
                    })
                except Exception:
                    pass

            # ── 케이스 C: 숫자로 시작 (B/S 동일 반복, 생략된 후속행) ──
            # 01A101: 수량 가격 거래금액 손익 수수료 (5개) / 01A103: + 체결시간 (6개)
            elif NUM_RE.match(line) and current_code and current_bs and len(parts) >= 3:
                has_time = len(parts) >= 4 and bool(TIME_RE.match(parts[2]))
                off = 1 if has_time else 0
                try:
                    result["체결"].append({
                        "종목":     current_code,
                        "현재가":   _num(current_현재가),
                        "BS":       current_bs,
                        "수량":     _num(parts[0]),
                        "가격":     _num(parts[1]),
                        "체결시간": parts[2] if has_time else "",
                        "거래금액": _num(parts[2 + off]),
                        "손익":     _num(parts[3 + off]) if len(parts) > 3 + off else "",
                        "수수료":   _num(parts[4 + off]) if len(parts) > 4 + off else "",
                    })
                except Exception:
                    pass

        elif section in ("미결_당일", "미결_전일") and len(parts) >= 5 and CODE_RE.match(line):
            # 종목 현재가 거래구분 B/S 잔량
            try:
                result["미결"].append({
                    "종목":     parts[0],
                    "현재가":   _num(parts[1]),
                    "거래구분": parts[2],
                    "BS":       parts[3],
                    "잔량":     _num(parts[4]),
                    "구분":     "당일" if section == "미결_당일" else "전일",
                })
            except Exception:
                pass

    # 수수료 합계
    for line in lines:
        m = re.search(r"당일수수료합계\s*[:：]?\s*([\d,]+)", line)
        if m:
            result["요약"]["수수료합계"] = _num(m.group(1))

    return result


# ─────────────────────────────────────────────
# SS선물(삼성선물) 파서
# ─────────────────────────────────────────────
def parse_ss(pdf_path: str) -> dict:
    """
    삼성선물 거래내역보고서 파싱
    대상 섹션: [ 체결내역 ]
    컬럼: 종목 구분 수량 가격 조건구분 결제구분 시간 번호
    추가 섹션: [ 거래내역 ], [ 미결제약정 ]
    """
    result = {
        "source": "SS선물",
        "date": "",
        "account": "",
        "체결": [],    # [ 체결내역 ] — 건별 체결
        "거래": [],    # [ 거래내역 ] — 종목별 집계
        "미결": [],    # [ 미결제약정 ]
        "요약": {},
    }

    lines = _all_lines(pdf_path)

    # 날짜 / 계좌번호
    for line in lines:
        m = re.search(r"거래일자\s*[:：]\s*(\d{4}/\d{2}/\d{2})", line)
        if m:
            result["date"] = m.group(1).replace("/", "-")
        m = re.search(r"위탁계좌번호\s*[:：]\s*([\d\-]+)", line)
        if m:
            result["account"] = m.group(1)

    SEC_EXEC    = "체결내역"     # [ 체결내역 ]
    SEC_TRADE   = "거래내역"     # [ 거래내역 ]
    SEC_OPEN    = "미결제약정"   # [ 미결제약정 ]

    CODE_RE  = re.compile(r"^[A-Z]\d{4}")   # SS: A6566 형태
    CODE2_RE = re.compile(r"^[A-Z]\d{4}")   # 거래내역의 두번째 컬럼 패턴
    SKIP     = {"합계", "소계", "소 계", "합 계"}

    section = None

    for line in lines:
        # 섹션 감지
        if "[" in line and "]" in line:
            if SEC_EXEC in line:
                section = "체결"
            elif SEC_TRADE in line and SEC_EXEC not in line and "미결" not in line:
                section = "거래"
            elif SEC_OPEN in line:
                section = "미결"
            else:
                section = None
            continue

        if section is None:
            continue

        # 합계/소계 행 스킵
        if any(k in line for k in SKIP):
            continue

        parts = line.split()
        if len(parts) < 4:
            continue

        # ── [ 체결내역 ] ──
        # 종목 구분 수량 가격 조건구분 결제구분 시간 번호
        if section == "체결":
            if not CODE_RE.match(line):
                continue
            try:
                result["체결"].append({
                    "종목":     parts[0],
                    "구분":     parts[1],
                    "수량":     _num(parts[2]),
                    "가격":     _num(parts[3]),
                    "조건구분": parts[4] if len(parts) > 4 else "",
                    "결제구분": parts[5] if len(parts) > 5 else "",
                    "시간":     parts[6] if len(parts) > 6 else "",
                    "번호":     parts[7] if len(parts) > 7 else "",
                })
            except Exception:
                pass

        # ── [ 거래내역 ] ──
        # 거래유형 종목 구분 수량 체결가 결제가 거래금액 수수료 실현손익
        elif section == "거래":
            if len(parts) >= 8 and CODE2_RE.match(parts[1] if len(parts) > 1 else ""):
                try:
                    result["거래"].append({
                        "거래유형": parts[0],
                        "종목":     parts[1],
                        "구분":     parts[2],
                        "수량":     _num(parts[3]),
                        "체결가":   _num(parts[4]),
                        "결제가":   _num(parts[5]),
                        "거래금액": _num(parts[6]),
                        "수수료":   _num(parts[7]),
                        "실현손익": _num(parts[8]) if len(parts) > 8 else "",
                    })
                except Exception:
                    pass

        # ── [ 미결제약정 ] ──
        # 종목 구분 잔량 미결잔량 평균매입가 현재가 평가손익 [옵션가치]
        elif section == "미결":
            if not CODE_RE.match(line):
                continue
            if len(parts) >= 6:
                try:
                    result["미결"].append({
                        "종목":     parts[0],
                        "구분":     parts[1],
                        "잔량":     _num(parts[2]),
                        "미결잔량": _num(parts[3]),
                        "평균매입가": _num(parts[4]),
                        "현재가":   _num(parts[5]),
                        "평가손익": _num(parts[6]) if len(parts) > 6 else "",
                    })
                except Exception:
                    pass

    # 위탁자산현황 요약
    for line in lines:
        m = re.search(r"당일거래손익\s*[:：]?\s*([\-\d,]+)", line)
        if m:
            result["요약"]["당일거래손익"] = _num(m.group(1))
        m = re.search(r"당일수수료\s*[:：]?\s*([\d,]+)", line)
        if m:
            result["요약"]["당일수수료"] = _num(m.group(1))

    return result
