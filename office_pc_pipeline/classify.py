# -*- coding: utf-8 -*-
"""
거래 구분(전략) 분류 — 하이브리드.

우선순위: overrides(개별 체결/포지션 규칙) → broker_default(출처별 기본값) → '미분류'.
의도(커브냐 방향성이냐)는 체결 데이터에 없으므로, 중개사별 기본값으로 자동 초안을 잡고
예외는 trade_tags.json overrides로 개별 재분류한다(하이브리드).

설정(trade_tags.json)은 단일 진실원천 — 저널·리포트가 매번 이 파일을 읽어 실시간 분류하므로,
분류를 바꿔도 Excel 재생성이 필요 없다.
"""
import fnmatch
import json
from pathlib import Path

CFG_FILE = Path(__file__).parent / "trade_tags.json"


def load_cfg():
    if CFG_FILE.exists():
        try:
            return json.loads(CFG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"broker_default": {}, "ovn_default": "미분류", "overrides": []}


def _hm(s):
    """'9:15'/'09:15' → 분 단위 정수(비교용). 실패 시 None."""
    try:
        h, m = str(s).split(":")[:2]
        return int(h) * 60 + int(m)
    except Exception:
        return None


def _match(rule, fill):
    """rule의 '지정된 필드만' 일치하면 True(미지정 필드는 무시)."""
    d = fill.get("date", "")
    if "date" in rule and rule["date"] != d:
        return False
    if "date_from" in rule and d < rule["date_from"]:
        return False
    if "date_to" in rule and d > rule["date_to"]:
        return False
    if rule.get("ovn") and not fill.get("ovn"):
        return False
    if "source" in rule and rule["source"] != fill.get("source"):
        return False
    if "code" in rule and not fnmatch.fnmatch(str(fill.get("code", "")), rule["code"]):
        return False
    if "side" in rule and rule["side"] != fill.get("side"):
        return False
    t = _hm(fill.get("time", ""))
    if "time_from" in rule and (t is None or t < _hm(rule["time_from"])):
        return False
    if "time_to" in rule and (t is None or t > _hm(rule["time_to"])):
        return False
    return True


def classify(fill, cfg=None):
    """
    fill: {date, source, code, side, time, ovn(bool)} 중 있는 것만.
    반환: 구분 문자열.
    """
    cfg = cfg if cfg is not None else load_cfg()
    for rule in cfg.get("overrides", []):
        if _match(rule, fill):
            return rule.get("구분", "미분류")
    if fill.get("ovn"):
        return cfg.get("ovn_default", "미분류")
    return cfg.get("broker_default", {}).get(fill.get("source"), "미분류")
