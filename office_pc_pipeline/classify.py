# -*- coding: utf-8 -*-
"""
거래 구분(전략) 분류.

우선순위: overrides(개별 규칙) → 실제 커브 판정 → ovn_default / broker_default.

★커브는 중개사나 '오버나잇 여부'로 찍지 않는다 — 커브는 매칭되는 두 다리(3Y↔10Y)가 필요하다.
  그날 '진입 오버나잇'이 KTB3와 KTB10을 반대부호로 동시 보유할 때만, 그 OVN 레그를 커브로 본다.
  그 외(단일 상품 아웃라이트, 인트라데이 스캘프)는 방향성. 예외는 trade_tags.json overrides로 재분류.

설정(trade_tags.json)은 단일 진실원천 — 저널·리포트가 매번 읽어 실시간 분류(Excel 재생성 불필요).
분류는 어디에도 누적 저장하지 않으므로, 이 로직/설정을 고치면 과거 기간리뷰도 자동으로 재계산된다.
"""
import fnmatch
import json
from pathlib import Path

CFG_FILE = Path(__file__).parent / "trade_tags.json"
SNAP_DIR = Path(__file__).parent / "ovn_snapshots"

_curve_cache = {}


def load_cfg():
    if CFG_FILE.exists():
        try:
            return json.loads(CFG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"broker_default": {}, "ovn_default": "방향성", "overrides": []}


def _is_curve_ovn_date(date) -> bool:
    """그날 진입 오버나잇이 KTB3·KTB10을 반대부호로 '동시 보유'하면 True(실제 커브 캐리)."""
    d = str(date or "")
    if d in _curve_cache:
        return _curve_cache[d]
    res = False
    p = SNAP_DIR / (d.replace("-", "") + ".json")
    if p.exists():
        try:
            e = json.loads(p.read_text(encoding="utf-8")).get("entry", {}) or {}
            k3 = int(e.get("ktb3", 0) or 0)
            k10 = int(e.get("ktb10", 0) or 0)
            res = bool(k3 and k10 and (k3 > 0) != (k10 > 0))
        except Exception:
            res = False
    _curve_cache[d] = res
    return res


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
    # 실제 커브: OVN 캐리 레그이면서, 그날 진입 OVN이 3Y·10Y 반대 동시보유일 때만
    if fill.get("ovn") and _is_curve_ovn_date(fill.get("date", "")):
        return "커브"
    if fill.get("ovn"):
        return cfg.get("ovn_default", "방향성")
    return cfg.get("broker_default", {}).get(fill.get("source"), "방향성")
