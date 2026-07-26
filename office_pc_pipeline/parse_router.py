# -*- coding: utf-8 -*-
"""
다운로드된 자료(PDF/CSV)를 파싱하고 종목사·날짜별로 하나의 결과로 병합한다.

원칙 (NH·SS 대칭):
  체결 = CSV 우선(틱단위), CSV 없으면 PDF
  미결 = 있는 소스에서(SS는 PDF, NH는 CSV). 크로스체크 앵커.

한 (source, date)에 대해 여러 파일이 오면 항목별로 최선을 골라 하나로 합친다.
파일 종류를 늘려도 이 규칙만 유지하면 되므로 특수 케이스 분기가 없다.
"""
from pdf_parser import parse_nh, parse_ss, parse_ss_csv, parse_nh_csv

# (source, kind) → 파서
_PARSERS = {
    ("NH선물", "pdf"): parse_nh,
    ("NH선물", "csv"): parse_nh_csv,
    ("SS선물", "pdf"): parse_ss,
    ("SS선물", "csv"): parse_ss_csv,
}


def _empty(source, date):
    return {"source": source, "date": date, "account": "",
            "체결": [], "거래": [], "미결": [], "요약": {}}


def parse_and_merge(downloads: list, log=print) -> list:
    """
    downloads: fetch_pdfs() 결과 [{source, path, date, kind, role}, ...]
    반환: [merged_parsed, ...]  (parse_nh/parse_ss와 동일 스키마, source+date별 1개)
    """
    # 1) 개별 파싱 → (source, date)별로 모음
    grouped: dict[tuple, list] = {}
    for item in downloads:
        parser = _PARSERS.get((item["source"], item.get("kind", "pdf")))
        if not parser:
            log(f"[경고] 파서 없음: {item['source']}/{item.get('kind')}")
            continue
        try:
            p = parser(str(item["path"]))
        except Exception as e:
            import traceback
            fname = item["path"].name if hasattr(item["path"], "name") else item["path"]
            log(f"[오류] 파싱 실패 {item.get('role','')} {fname} — {e}")
            log(traceback.format_exc())
            continue
        d = p.get("date") or item["date"]
        p["date"] = d
        p["_kind"] = item.get("kind", "pdf")
        grouped.setdefault((item["source"], d), []).append(p)

    # 2) 그룹별 병합: 체결은 CSV 우선, 미결/거래/계좌는 값이 있는 소스에서
    merged_list = []
    for (source, d), parts in sorted(grouped.items()):
        m = _empty(source, d)
        csv_exec = next((p["체결"] for p in parts if p.get("_kind") == "csv" and p.get("체결")), None)
        pdf_exec = next((p["체결"] for p in parts if p.get("_kind") == "pdf" and p.get("체결")), None)
        m["체결"] = csv_exec if csv_exec is not None else (pdf_exec or [])

        for p in parts:
            if not m["account"] and p.get("account"):
                m["account"] = p["account"]
            if not m["미결"] and p.get("미결"):
                m["미결"] = p["미결"]
            if not m["거래"] and p.get("거래"):
                m["거래"] = p["거래"]
            if not m["요약"] and p.get("요약"):
                m["요약"] = p["요약"]

        srcs = "+".join(sorted({p.get("_kind", "?") for p in parts}))
        exec_src = "CSV" if csv_exec is not None else "PDF" if pdf_exec else "없음"
        log(f"  [{source}] 병합 ({d}) — 체결 {len(m['체결'])}건[{exec_src}], "
            f"미결 {len(m['미결'])}행  (소스: {srcs})")
        if source == "SS선물" and not m["미결"]:
            log(f"  [경고] {source} ({d}) 미결제약정 없음 — 크로스체크 불가 (PDF 미수신?)")
        merged_list.append(m)

    return merged_list
