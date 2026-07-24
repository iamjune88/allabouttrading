# -*- coding: utf-8 -*-
"""Shift+Z 로 열린 엑셀보내기 워크북을 잡아 저장/닫는다.

인포맥스 엑셀보내기는 매번 **새 Excel 인스턴스(프로세스)** 를 띄우는 경우가 있어,
GetActiveObject 로는 엉뚱한 인스턴스를 잡는다. 따라서 Shift+Z 직전 XLMAIN 창 목록을
스냅샷하고, 직후 '새로 생긴 XLMAIN 창'을 comtypes(AccessibleObjectFromWindow)로 붙어
그 인스턴스의 export 워크북만 정확히 저장/닫는다.
"""
import time, ctypes
from ctypes import oledll, byref, POINTER
import win32gui
import comtypes, comtypes.client
from comtypes.automation import IDispatch

_IID_IDISPATCH = comtypes.GUID("{00020400-0000-0000-C000-000000000046}")
_OBJID_NATIVEOM = 0xFFFFFFF0


def snapshot_xlmain():
    hs = set()
    def eh(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == 'XLMAIN':
            hs.add(h)
        return True
    win32gui.EnumWindows(eh, None)
    return hs


def _excel7_child(top):
    kids = []
    def cb(h, _):
        if win32gui.GetClassName(h) == 'EXCEL7':
            kids.append(h)
        return True
    win32gui.EnumChildWindows(top, cb, None)
    return kids[0] if kids else None


def _app_from_window(top_hwnd):
    h7 = _excel7_child(top_hwnd)
    if not h7:
        return None
    ptr = POINTER(IDispatch)()
    oledll.oleacc.AccessibleObjectFromWindow(h7, _OBJID_NATIVEOM, byref(_IID_IDISPATCH), byref(ptr))
    return comtypes.client.GetBestInterface(ptr).Application


def _is_export_wb(wb):
    try:
        if str(wb.Path):
            return False
        ur = wb.Worksheets(1).UsedRange
        return ur.Columns.Count >= 11 and ur.Rows.Count > 50
    except Exception:
        return False


def _save_close(app, wb, save_path):
    app.DisplayAlerts = False
    try:
        wb.SaveAs(save_path, 51)      # 51 = xlsx (positional; comtypes 는 키워드 불가)
        wb.Close(False)
        # export 전용으로 뜬 빈 인스턴스면 종료(프로세스 누적 방지)
        try:
            if app.Workbooks.Count == 0:
                app.Quit()
        except Exception:
            pass
    finally:
        try:
            app.DisplayAlerts = True
        except Exception:
            pass


def capture_export(save_path, before_hwnds, timeout=20.0):
    """Shift+Z 직후 새로 생긴 XLMAIN 창에서 export 워크북을 저장/닫는다.
    성공 True / 타임아웃 False."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        new = snapshot_xlmain() - before_hwnds
        for h in list(new):
            try:
                app = _app_from_window(h)
                if app is None:
                    continue
                for wb in app.Workbooks:
                    if _is_export_wb(wb):
                        _save_close(app, wb, save_path)
                        return True
            except Exception:
                pass
        time.sleep(0.4)
    return False


def cleanup_exports():
    """열려있는 모든 미저장 export 워크북 정리 + 빈 인스턴스 종료(안전용)."""
    n = 0
    for h in list(snapshot_xlmain()):
        try:
            app = _app_from_window(h)
            if app is None:
                continue
            app.DisplayAlerts = False
            for wb in list(app.Workbooks):
                if _is_export_wb(wb):
                    wb.Close(False); n += 1
            try:
                if app.Workbooks.Count == 0:
                    app.Quit()
            except Exception:
                pass
            app.DisplayAlerts = True
        except Exception:
            pass
    return n
