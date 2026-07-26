# -*- coding: utf-8 -*-
"""
NH선물 / SS선물(삼성선물) PDF·CSV 파싱 모듈
- NH선물:  [당일 선물거래] 섹션(PDF) / 01A103 CSV(체결+미결)
- SS선물:  [체결내역][미결제약정] 섹션(PDF) / 매매내역 CSV(체결)
"""
import re


def _num(val):
    """콤마 제거 후 숫자 변환. 실패 시 원본 반환."""
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return val


def _all_lines(pdf_path: str) -> list[str]:
    """PDF 전 페이지 텍스트를 줄 단위 리스트로 반환."""
    import pdfplumber
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
    """NH선물 거래확인서 PDF 파싱. 대상: [당일 선물거래]·[전일/당일 미결제]."""
    return _parse_nh_lines(_all_lines(pdf_path))


def _parse_nh_lines(lines: list) -> dict:
    result = {
        "source": "NH선물",
        "date": "",
        "account": "",
        "체결": [],
        "미결": [],
        "요약": {},
    }

    for line in lines:
        m = re.search(r"거래일자\s*[:：]\s*(\d{4}-\d{2}-\d{2})", line)
        if m:
            result["date"] = m.group(1)
        m = re.search(r"계좌번호\s*[:：]\s*([\d\-]+)", line)
        if m:
            result["account"] = m.group(1)

    SEC_TRADE_KEYWORDS = ("당일 선물거래", "당일선물거래", "당일선물매매", "당일 선물매매")
    SEC_OPEN_D = "당일미결제"
    SEC_OPEN_P = "전일미결제"

    CODE_RE = re.compile(r"^[A-Z]\d{6,8}")
    BS_RE = re.compile(r"^(매수|매도)")
    NUM_RE = re.compile(r"^\d")
    TIME_RE = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")
    SKIP = {"합계", "소계", "소 계", "합 계", "총계"}

    section = None
    current_code = None
    current_현재가 = ""
    current_bs = ""

    for line in lines:
        if "[" in line and "]" in line:
            _compact = line.replace(" ", "")  # '[당일 미결제]' → '[당일미결제]'
            if any(kw in line for kw in SEC_TRADE_KEYWORDS):
                section = "체결"
                current_code = None
            elif SEC_OPEN_D in _compact:
                section = "미결_당일"
                current_code = None
            elif SEC_OPEN_P in _compact:
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
            if CODE_RE.match(line) and len(parts) >= 7:
                current_code = parts[0]
                current_현재가 = parts[1]
                current_bs = parts[3]
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
            # 전일: 종목 정산가 미결제 B/S 수량 가격 잔고금액 잔고손익
            # 당일: 종목 정산가 미결제 B/S 수량
            try:
                result["미결"].append({
                    "종목":     parts[0],
                    "정산가":   _num(parts[1]),
                    "거래유형": parts[2],
                    "BS":       parts[3],
                    "잔량":     _num(parts[4]),
                    "구분":     "당일" if section == "미결_당일" else "전일",
                })
            except Exception:
                pass

    for line in lines:
        m = re.search(r"당일수수료합계\s*[:：]?\s*([\d,]+)", line)
        if m:
            result["요약"]["수수료합계"] = _num(m.group(1))

    return result


# ─────────────────────────────────────────────
# SS선물(삼성선물) PDF 파서
# ─────────────────────────────────────────────
def parse_ss(pdf_path: str) -> dict:
    """삼성선물 가정산보고서 PDF 파싱. 대상: [체결내역][거래내역][미결제약정]."""
    return _parse_ss_lines(_all_lines(pdf_path))


def _parse_ss_lines(lines: list) -> dict:
    result = {
        "source": "SS선물",
        "date": "",
        "account": "",
        "체결": [],
        "거래": [],
        "미결": [],
        "요약": {},
    }

    for line in lines:
        m = re.search(r"거래일자\s*[:：]\s*(\d{4}/\d{2}/\d{2})", line)
        if m:
            result["date"] = m.group(1).replace("/", "-")
        m = re.search(r"위탁계좌번호\s*[:：]\s*([\d\-]+)", line)
        if m:
            result["account"] = m.group(1)

    SEC_EXEC = "체결내역"
    SEC_TRADE = "거래내역"
    SEC_OPEN = "미결제약정"

    CODE_RE = re.compile(r"^[A-Z]\d{4}")
    CODE2_RE = re.compile(r"^[A-Z]\d{4}")
    SKIP = {"합계", "소계", "소 계", "합 계"}

    section = None

    for line in lines:
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
        if any(k in line for k in SKIP):
            continue
        parts = line.split()
        if len(parts) < 4:
            continue

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

        # ── [ 미결제약정 ] 실제 컬럼: 종목 구분 잔고 전일잔고 전일정산가 당일정산가 갱신차금 [옵션가치]
        #   잔고=오늘 마감 OVN, 전일잔고=오늘 진입 OVN, 갱신차금=브로커 확정 MTM
        elif section == "미결":
            if not CODE_RE.match(line):
                continue
            if len(parts) >= 7:
                try:
                    result["미결"].append({
                        "종목":       parts[0],
                        "구분":       parts[1],
                        "잔고":       _num(parts[2]),
                        "전일잔고":   _num(parts[3]),
                        "전일정산가": _num(parts[4]),
                        "당일정산가": _num(parts[5]),
                        "갱신차금":   _num(parts[6]),
                        "옵션가치":   _num(parts[7]) if len(parts) > 7 else "",
                    })
                except Exception:
                    pass

    for line in lines:
        m = re.search(r"당일거래손익\s*[:：]?\s*([\-\d,]+)", line)
        if m:
            result["요약"]["당일거래손익"] = _num(m.group(1))
        m = re.search(r"당일수수료\s*[:：]?\s*([\d,]+)", line)
        if m:
            result["요약"]["당일수수료"] = _num(m.group(1))

    return result


# ─────────────────────────────────────────────
# SS선물 CSV 파서 (체결내역 — 틱단위, 미결제약정 없음)
# 인코딩: CP949. 컬럼: 날짜(YYYYMMDD) 거래시간(HH:MM) 계좌 종목 방향(B/S) 수량 가격
# ─────────────────────────────────────────────
_SS_CSV_BS = {"B": "매수", "S": "매도", "매수": "매수", "매도": "매도"}


def parse_ss_csv(csv_path: str) -> dict:
    """삼성선물 매매내역 CSV → parse_ss() '체결' 스키마. 같은 (시각,방향,가격) 수량 합산."""
    return _parse_ss_csv_text(_read_text_cp949(csv_path))


def _read_text_cp949(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("cp949", "euc-kr", "utf-8-sig", "utf-8"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("cp949", errors="replace")


def _parse_ss_csv_text(text: str) -> dict:
    result = {
        "source": "SS선물",
        "date": "",
        "account": "",
        "체결": [],
        "거래": [],
        "미결": [],
        "요약": {},
        "_csv": True,
    }

    bucket = {}
    order = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        cols = [c.strip() for c in line.split(",")]
        if len(cols) < 7 or not re.match(r"^A\d{4}", cols[3]):
            continue
        date_raw, time_s, account, code, bs, qty, price = cols[:7]
        if not result["date"] and re.match(r"^\d{8}$", date_raw):
            result["date"] = f"{date_raw[:4]}-{date_raw[4:6]}-{date_raw[6:8]}"
        if not result["account"]:
            result["account"] = account
        side = _SS_CSV_BS.get(bs.upper(), bs)
        key = (time_s, side, _num(price), code)
        if key not in bucket:
            bucket[key] = 0.0
            order.append(key)
        bucket[key] += _num(qty)

    for (time_s, side, price, code) in order:
        result["체결"].append({
            "종목":     code,
            "구분":     side,
            "수량":     bucket[(time_s, side, price, code)],
            "가격":     price,
            "조건구분": "",
            "결제구분": "",
            "시간":     time_s,
            "번호":     "",
        })

    return result


# ─────────────────────────────────────────────
# NH선물 CSV 파서 (01A103: 체결시분 — 체결+미결 모두 포함)
# 컬럼: 종목 정산가 거래유형(미결제/체결) B/S 매수량 매도량 가격 [체결시간HHMM] 약정금액 손익 [수수료]
#   상단 미결행 = 전일미결(진입 OVN), 하단(빈줄 뒤) 미결행 = 당일미결(종료 OVN).
# ─────────────────────────────────────────────
def parse_nh_csv(csv_path: str) -> dict:
    """NH선물 01A103 CSV → parse_nh() 스키마(체결/미결). CP949."""
    return _parse_nh_csv_text(_read_text_cp949(csv_path))


def _hhmm(v):
    """'0855' → '08:55'. 시간 아니면 ''."""
    s = str(v).strip()
    if re.fullmatch(r"\d{3,4}", s):
        s = s.zfill(4)
        return f"{s[:2]}:{s[2:]}"
    return ""


def _parse_nh_csv_text(text: str) -> dict:
    result = {
        "source": "NH선물",
        "date": "",
        "account": "",
        "체결": [],
        "미결": [],
        "요약": {},
        "_csv": True,
    }

    seen_exec = False
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        cols = [c.strip() for c in line.split(",")]

        if not re.match(r"^A\d{4}", cols[0]):
            for c in cols:
                if re.fullmatch(r"\d{8}", c) and not result["date"]:
                    result["date"] = f"{c[:4]}-{c[4:6]}-{c[6:8]}"
                    break
            if not result["account"] and cols and re.fullmatch(r"\d{10,}", cols[0]):
                result["account"] = cols[0]
            continue

        if len(cols) < 7:
            continue
        code, settle, kind, bs, buy_q, sell_q, price = cols[:7]
        signed_qty = _num(buy_q) if _num(buy_q) else -_num(sell_q) if _num(sell_q) else 0
        side = "매수" if _num(buy_q) else "매도" if _num(sell_q) else bs

        if "체결" in kind:
            seen_exec = True
            time_s = _hhmm(cols[7]) if len(cols) > 7 else ""
            rest = cols[8:] if time_s else cols[7:]
            result["체결"].append({
                "종목":     code,
                "현재가":   _num(settle),
                "BS":       side,
                "수량":     abs(signed_qty),
                "가격":     _num(price),
                "체결시간": time_s,
                "거래금액": _num(rest[0]) if len(rest) > 0 else "",
                "손익":     _num(rest[1]) if len(rest) > 1 else "",
                "수수료":   _num(rest[2]) if len(rest) > 2 else "",
            })
        elif "미결" in kind:
            result["미결"].append({
                "종목":     code,
                "정산가":   _num(settle),
                "거래유형": "미결제",
                "BS":       side,
                "잔량":     abs(signed_qty),
                "구분":     "당일" if seen_exec else "전일",
            })

    return result
