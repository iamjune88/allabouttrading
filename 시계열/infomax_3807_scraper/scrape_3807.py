# -*- coding: utf-8 -*-
"""인포맥스 3807(투자자별 매매동향) 일자별 스크래퍼.

동작(하루 1회 반복):
  1) 3807 창을 찾아 포그라운드
  2) FROM 날짜칸에 대상일자 입력(+Enter) -> 그 날짜 당일 데이터로 재조회
  3) Shift+Z -> 엑셀보내기(새 워크북)
  4) win32com 으로 워크북을 raw/3807_YYYYMMDD.xlsx 로 저장/닫기
  5) 파싱 -> intraday.csv / daily.csv 누적, state.json 기록(재실행 이어받기)

전제: 실행 시 불필요한 Excel 창을 닫아둔 깨끗한 환경(1회성 배치), 3807 화면이 열려 있어야 함.

사용:
  python scrape_3807.py --start 2024-01-01 --end 2026-07-24
  python scrape_3807.py --start 2026-07-01 --end 2026-07-24 --descending
  python scrape_3807.py --dry-run            # 대상일자 목록만 출력
"""
import argparse, os, sys, time, datetime as dt
import ctypes
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass
import win32gui, win32con
from pywinauto import Application, mouse, keyboard

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from parser import parse_export
import storage
import excel_capture
from calendar_krx import candidate_trading_days

# ---- 타이밍(느리면 늘리세요) ----
REQUERY_WAIT = 1.8      # FROM 입력 후 재조회 대기(초)
EXPORT_WAIT  = 0.6      # Shift+Z 후 초기 대기
CAPTURE_TIMEOUT = 20.0  # 엑셀 워크북 등장 대기(초)
BETWEEN_DAYS = 0.4      # 날짜 간 간격
MIN_BARS_TRADING = 30   # 이보다 적으면 휴장/무자료로 간주


def _text_w(h):
    n = win32gui.GetWindowTextLength(h)
    buf = ctypes.create_unicode_buffer(n + 1)
    ctypes.windll.user32.GetWindowTextW(h, buf, n + 1)
    return buf.value


def find_3807_window():
    """infomaxmain 의 AfxFrameOrView120 중 (날짜칸>=2 + GXWND + 종목콤보)을 가진 3807 창 탐색.
    반환: (win_hwnd, grid_hwnd, [date_field_hwnds])"""
    import psutil
    import win32process
    pids = [p.info['pid'] for p in psutil.process_iter(['name', 'pid'])
            if p.info['name'] and p.info['name'].lower() == 'infomaxmain.exe']

    candidates = []

    def scan(top):
        dates, grid, combos = [], None, []
        def cb(h, _):
            nonlocal grid
            cls = win32gui.GetClassName(h)
            t = _text_w(h)
            if cls == 'GXWND':
                grid = grid or h
            if t.isdigit() and len(t) == 8 and win32gui.IsWindowVisible(h):
                dates.append((win32gui.GetWindowRect(h)[0], h))
            if cls.startswith('Afx:') and t:
                combos.append(t)
            return True
        win32gui.EnumChildWindows(top, cb, None)
        return dates, grid, combos

    def eh(h, _):
        _, pid = win32process.GetWindowThreadProcessId(h)
        if pid in pids and win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == 'AfxFrameOrView120':
            r = win32gui.GetWindowRect(h)
            if (r[2]-r[0]) > 300 and (r[3]-r[1]) > 200:
                dates, grid, combos = scan(h)
                if len(dates) >= 2 and grid is not None:
                    # KTB/선물 등 종목 힌트가 있으면 가점
                    hint = any(('KTB' in c) or ('_' in c and c[:3].isalpha()) for c in combos)
                    candidates.append((hint, h, grid, sorted(dates)))
        return True
    win32gui.EnumWindows(eh, None)

    if not candidates:
        raise RuntimeError("3807 창을 찾지 못했습니다. 화면이 열려 있는지 확인하세요.")
    candidates.sort(key=lambda x: not x[0])   # hint=True 우선
    _, win_h, grid_h, dates = candidates[0]
    return win_h, grid_h, [h for _, h in dates]


def date_fields(win_h):
    f = []
    def cb(h, _):
        t = _text_w(h)
        if t.isdigit() and len(t) == 8 and win32gui.IsWindowVisible(h):
            f.append((win32gui.GetWindowRect(h)[0], h, t))
        return True
    win32gui.EnumChildWindows(win_h, cb, None)
    f.sort()
    return f


def set_from_date(win_h, app, date):
    """FROM(가장 왼쪽) 날짜칸에 date 입력. 좌표 클릭으로 최상위(보이는) 컨트롤을 잡음."""
    fields = date_fields(win_h)
    if not fields:
        raise RuntimeError("날짜 입력칸을 찾지 못함")
    from_h = fields[0][1]
    r = win32gui.GetWindowRect(from_h)
    cx, cy = (r[0] + r[2]) // 2, (r[1] + r[3]) // 2
    ymd = f"{date:%Y/%m/%d}"
    mouse.click(button='left', coords=(cx, cy)); time.sleep(0.25)
    keyboard.send_keys('^a'); time.sleep(0.12)
    keyboard.send_keys(ymd.replace('/', '{/}')); time.sleep(0.15)
    keyboard.send_keys('{ENTER}')
    time.sleep(REQUERY_WAIT)
    # 검증: 가장 왼쪽 그룹 값에 대상일이 있는지
    want = f"{date:%Y%m%d}"
    got = [t for _, _, t in date_fields(win_h)]
    left_x = min(x for x, _, _ in date_fields(win_h))
    left_vals = [t for x, _, t in date_fields(win_h) if x == left_x]
    return want in left_vals, left_vals, got


def scrape_one(win_h, grid_h, app, date):
    """한 날짜 처리. 반환 status: 'ok' | 'holiday' | 'error'."""
    app.window(handle=win_h).set_focus(); time.sleep(0.3)
    ok, left_vals, allvals = set_from_date(win_h, app, date)
    if not ok:
        # 한 번 재시도
        set_from_date(win_h, app, date)
    # Shift+Z 내보내기 (직전 XLMAIN 창 스냅샷 -> 새로 생긴 창만 캡처)
    app.window(handle=win_h).set_focus(); time.sleep(0.2)
    before = excel_capture.snapshot_xlmain()
    keyboard.send_keys('+z')
    time.sleep(EXPORT_WAIT)
    raw = os.path.join(storage.RAW_DIR, f"3807_{date:%Y%m%d}.xlsx")
    if os.path.exists(raw):
        os.remove(raw)
    if not excel_capture.capture_export(raw, before, timeout=CAPTURE_TIMEOUT):
        return 'error', "엑셀 워크북 캡처 실패(타임아웃)"
    # 파싱
    try:
        df, summ = parse_export(raw, date)
    except Exception as e:
        return 'error', f"파싱 실패: {e}"
    if summ['n_bars'] < MIN_BARS_TRADING:
        return 'holiday', f"봉 {summ['n_bars']}개(휴장/무자료 추정)"
    storage.append_intraday(df)
    storage.append_daily(summ)
    warn = f"  ⚠절단의심(첫봉 {summ['first_time']})" if summ.get('truncated') else ""
    return 'ok', (f"봉 {summ['n_bars']} | 외국인 {summ['frgn_net']:+,} "
                  f"증권 {summ['secu_net']:+,} 투신 {summ['trust_net']:+,} 은행 {summ['bank_net']:+,}{warn}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--start', required=False)
    ap.add_argument('--end', required=False, default=str(dt.date.today()))
    ap.add_argument('--descending', action='store_true', help='최신일부터')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--retry-errors', action='store_true', help='이전 error 날짜만 재시도')
    ap.add_argument('--tag', default='', help='연도별 등 별도 파일로 저장 (예: --tag 2025 -> intraday_2025.csv)')
    args = ap.parse_args()

    if not args.start:
        print("--start YYYY-MM-DD 필요"); sys.exit(1)

    if args.tag:
        storage.set_output_tag(args.tag)
        print(f"출력 태그: {args.tag}  ({os.path.basename(storage.DAILY)}, {os.path.basename(storage.INTRADAY)})")

    days = candidate_trading_days(args.start, args.end, ascending=not args.descending)
    state = storage.load_state()

    todo = [d for d in days if args.retry_errors and state.get(str(d)) == 'error'
            or not args.retry_errors and not storage.is_done(state, d)]
    print(f"후보 거래일 {len(days)}개 중 처리 대상 {len(todo)}개 "
          f"(이미 완료 {len(days)-len(todo)}개)")
    if args.dry_run:
        for d in todo:
            print("  ", d)
        return

    win_h, grid_h, _ = find_3807_window()
    print(f"3807 창: hwnd={win_h}, grid={grid_h}")
    app = Application(backend="win32").connect(handle=win_h)

    ok = hol = err = 0
    t0 = time.time()
    for i, d in enumerate(todo, 1):
        try:
            status, msg = scrape_one(win_h, grid_h, app, d)
        except Exception as e:
            status, msg = 'error', f"예외: {e}"
        storage.mark(state, d, status)
        tag = {'ok': 'OK ', 'holiday': 'HOL', 'error': 'ERR'}[status]
        print(f"[{i}/{len(todo)}] {d} {tag} {msg}")
        ok += status == 'ok'; hol += status == 'holiday'; err += status == 'error'
        time.sleep(BETWEEN_DAYS)

    dtm = time.time() - t0
    print(f"\n완료: OK {ok} / 휴장 {hol} / 오류 {err}  ({dtm:.0f}s)")
    print(f"  intraday: {storage.INTRADAY}")
    print(f"  daily   : {storage.DAILY}")
    if err:
        print("  오류 재시도: python scrape_3807.py --start ... --retry-errors")


if __name__ == "__main__":
    main()
