# -*- coding: utf-8 -*-
"""
2026-07-01 KTB 트레이딩 저널 생성
캐리 포지션: KTB3 +240 / KTB10 -140  (스티프너 2.4세트 + 네이키드 -60)
오늘 체결: NH선물 KTB10(A6769) SELL 386 / BUY 283 → 추정 마감 KTB10 -243
"""

CHART_PATH = r"C:\Users\infomax\Desktop\claude협업\journal_assets\ktb_20260701_interactive.html"
OUT_PATH   = r"C:\Users\infomax\Desktop\claude협업\journal_20260701.html"

with open(CHART_PATH, encoding="utf-8") as f:
    chart_div = f.read()

# ── 체결 테이블 데이터 ─────────────────────────────────────────────────────
# (seq, side, qty, price, hh, mm, running_pos, pnl_won)
# 손익: 원본 Excel 합산값 기준 (None = 개별집계 생략)
OVERNIGHT_POS = -140  # KTB10 오버나잇

trades_raw = [
    ("SELL", 20, 106.79, 8, 45, -14400000),
    ("SELL",  3, 106.73, 8, 54,  -1980000),
    ("SELL", 17, 106.73, 8, 55, -11220000),  # 16+1 (각각 -10560000,-660000)
    ("SELL", 20, 106.69, 8, 59, -12400000),  # 6+11+3 (-3720000-6820000-1860000)
    ("SELL",  9, 106.61, 9,  2,  -4860000),  # 1+2+6 (-540000-1080000-3240000)
    ("SELL", 10, 106.56, 9,  5,  -4900000),
    ("SELL", 10, 106.51, 9,  8,  -4400000),
    ("BUY",  18, 106.59, 9, 22,  +9360000),  # 6+12
    ("BUY",  30, 106.54, 10, 11, +14100000),
    ("SELL", 10, 106.44, 10, 16,  -3700000),
    ("SELL", 30, 106.44, 10, 17, -11100000),
    ("SELL", 10, 106.41, 10, 19,  -3400000),
    ("BUY",  10, 106.46, 10, 22,  +3900000),
    ("SELL", 17, 106.36, 10, 46,  -4930000),  # 3+14
    ("BUY",  20, 106.39, 10, 53,  +6400000),
    ("SELL", 20, 106.31, 11,  6,  -4800000),
    ("BUY",  20, 106.35, 11,  8,  +5600000),
    ("SELL", 20, 106.27, 11, 31,  -4000000),
    ("SELL", 20, 106.21, 11, 39,  -2800000),
    ("BUY",  25, 106.21, 12,  6,  +3500000),
    ("BUY",  15, 106.22, 12,  6,  +2250000),
    ("SELL", 30, 106.07, 13,  8,         0),  # 10+20, 손익=0
    ("BUY",  30, 106.17, 13, 30,  +3000000),
    ("SELL", 30, 106.14, 14, 12,  -2100000),
    ("BUY",  30, 106.17, 14, 14,  +3000000),
    ("SELL", 20, 106.19, 14, 29,  -2400000),
    ("SELL", 20, 106.22, 14, 35,  -3000000),  # 1+19 (-150000-2850000)
    ("BUY",  20, 106.24, 14, 39,  +3400000),  # 3+17
    ("BUY",  30, 106.33, 15,  0,  +7800000),
    ("SELL",  5, 106.31, 15,  1,  -1200000),
    ("SELL", 20, 106.28, 15,  6,  -4200000),
    ("BUY",  35, 106.17, 15, 23,  +3500000),  # 4+30+1
    ("SELL", 10, 106.17, 15, 28,  -1000000),  # 9+1
    ("SELL",  5, 106.15, 15, 30,   -400000),  # 2+3
    ("SELL", 30, 106.07, 15, 45,         0),
]

# 누적 포지션 계산
rows = []
pos = OVERNIGHT_POS
cum_pnl = 0
for side, qty, price, hh, mm, pnl in trades_raw:
    if side == "SELL":
        pos -= qty
    else:
        pos += qty
    cum_pnl += pnl
    rows.append((side, qty, price, hh, mm, pos, pnl, cum_pnl))

total_sell = sum(r[1] for r in rows if r[0] == "SELL")
total_buy  = sum(r[1] for r in rows if r[0] == "BUY")
total_pnl  = sum(r[6] for r in rows)
final_pos  = rows[-1][5]  # -243

buy_trades  = [(q,p) for s,q,p,*_ in trades_raw if s == "BUY"]
sell_trades = [(q,p) for s,q,p,*_ in trades_raw if s == "SELL"]
buy_avg  = sum(q*p for q,p in buy_trades)  / sum(q for q,p in buy_trades)
sell_avg = sum(q*p for q,p in sell_trades) / sum(q for q,p in sell_trades)

# ── HTML 테이블 행 생성 ────────────────────────────────────────────────────
def fmt_pnl(v):
    if v == 0:
        return '<span style="color:#8b949e">—</span>'
    color = "#3fb950" if v > 0 else "#f85149"
    return f'<span style="color:{color}">{v/10000:+.0f}만</span>'

def fmt_side(s):
    if s == "SELL":
        return '<span class="side-sell">▼ SELL</span>'
    return '<span class="side-buy">▲ BUY</span>'

table_rows = ""
for i, (side, qty, price, hh, mm, pos, pnl, cum_pnl) in enumerate(rows, 1):
    pos_color = "#f85149" if pos < 0 else "#3fb950"
    table_rows += f"""
      <tr>
        <td>{i}</td>
        <td>{hh:02d}:{mm:02d}</td>
        <td>{fmt_side(side)}</td>
        <td>{qty}</td>
        <td>{price:.2f}</td>
        <td style="color:{pos_color};font-weight:600">{pos:+d}</td>
        <td>{fmt_pnl(pnl)}</td>
        <td>{fmt_pnl(cum_pnl)}</td>
      </tr>"""

# ── HTML 본문 ─────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>KTB 트레이딩 저널 — 2026-07-01</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{background:#0d1117;color:#e6edf3;font-family:'Segoe UI',Malgun Gothic,sans-serif;font-size:13px;line-height:1.5;padding:18px}}
  h1{{font-size:18px;font-weight:700;color:#e6edf3;margin-bottom:4px}}
  h2{{font-size:13px;font-weight:600;color:#8b949e;letter-spacing:.5px;text-transform:uppercase;margin:18px 0 8px}}
  .subtitle{{color:#8b949e;font-size:12px;margin-bottom:18px}}
  .cards{{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}}
  .card{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:12px 16px;min-width:160px;flex:1}}
  .card .label{{font-size:11px;color:#8b949e;margin-bottom:4px;letter-spacing:.4px}}
  .card .value{{font-size:20px;font-weight:700}}
  .card .sub{{font-size:11px;color:#8b949e;margin-top:3px}}
  .long{{color:#3fb950}}.short{{color:#f85149}}.neutral{{color:#e6edf3}}
  .stepper{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:14px 16px;margin-bottom:18px}}
  .stepper .row{{display:flex;justify-content:space-between;align-items:center;padding:4px 0;border-bottom:1px solid #21262d}}
  .stepper .row:last-child{{border-bottom:none}}
  .stepper .key{{color:#8b949e;font-size:11px}}
  .stepper .val{{font-weight:600;font-size:13px}}
  .chart-wrap{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:10px;margin-bottom:18px;overflow:hidden}}
  table{{width:100%;border-collapse:collapse;font-size:12px}}
  th{{background:#161b22;color:#8b949e;font-weight:600;padding:6px 8px;text-align:left;border-bottom:2px solid #21262d;position:sticky;top:0}}
  td{{padding:5px 8px;border-bottom:1px solid #161b22}}
  tr:hover td{{background:#161b22}}
  .side-sell{{color:#f85149;font-weight:700}}
  .side-buy{{color:#3fb950;font-weight:700}}
  .table-wrap{{max-height:420px;overflow-y:auto;border:1px solid #21262d;border-radius:6px;margin-bottom:18px}}
  .summary-row{{background:#161b22!important}}
  .summary-row td{{font-weight:700;color:#e6edf3;border-top:2px solid #21262d}}
  .comment-box{{background:#161b22;border:1px solid #21262d;border-radius:8px;padding:14px;margin-bottom:12px}}
  .comment-box label{{font-size:11px;color:#8b949e;display:block;margin-bottom:6px}}
  #commentArea{{width:100%;min-height:110px;background:#0d1117;color:#e6edf3;border:1px solid #30363d;border-radius:4px;padding:10px;font-size:13px;font-family:inherit;resize:vertical;outline:none}}
  #commentArea:focus{{border-color:#58a6ff}}
  .btn-save{{background:#238636;color:#fff;border:none;border-radius:6px;padding:8px 18px;font-size:13px;cursor:pointer;margin-right:8px}}
  .btn-save:hover{{background:#2ea043}}
  #saveStatus{{font-size:11px;color:#8b949e}}
  .note{{font-size:11px;color:#6e7681;margin-top:6px}}
  .pnl-note{{background:#161b22;border-left:3px solid #d29922;border-radius:0 6px 6px 0;padding:10px 14px;margin-bottom:18px;font-size:12px;color:#8b949e}}
  .pnl-note strong{{color:#e3b341}}
</style>
</head>
<body>

<h1>KTB 선물 트레이딩 저널</h1>
<p class="subtitle">2026-07-01 (수) &nbsp;|&nbsp; NH선물 A6769(KTB10 Sep) 체결 35버킷 · 49원본건 &nbsp;|&nbsp; KTB3·KTB10 병행 차트</p>

<!-- 포지션 카드 -->
<h2>포지션 요약</h2>
<div class="cards">
  <div class="card">
    <div class="label">KTB3 오버나잇 캐리</div>
    <div class="value long">+240</div>
    <div class="sub">스티프너 롱레그 2.4세트<br>변화 없음 (오늘 KTB3 체결 0건)</div>
  </div>
  <div class="card">
    <div class="label">KTB10 오버나잇 시작</div>
    <div class="value short">–140</div>
    <div class="sub">스티프너 숏레그 –80<br>네이키드 숏 –60</div>
  </div>
  <div class="card">
    <div class="label">오늘 KTB10 순변화</div>
    <div class="value short">–{total_sell - total_buy}</div>
    <div class="sub">SELL {total_sell}계약 / BUY {total_buy}계약</div>
  </div>
  <div class="card">
    <div class="label">KTB10 마감 추정</div>
    <div class="value short">{final_pos:+d}</div>
    <div class="sub">= –140 – {total_sell - total_buy}<br>미결잔고 PDF 확인 필요</div>
  </div>
  <div class="card">
    <div class="label">SELL avg / BUY avg</div>
    <div class="value neutral" style="font-size:16px">{sell_avg:.3f} / {buy_avg:.3f}</div>
    <div class="sub">스프레드 <span style="color:#3fb950">{sell_avg-buy_avg:+.3f}pt</span></div>
  </div>
</div>

<!-- 스티프너 구조 -->
<div class="stepper">
  <div class="row"><span class="key">스티프너 구성</span><span class="val">KTB3 +100 : KTB10 –33.3 (DV01 매칭, 1세트)</span></div>
  <div class="row"><span class="key">2.4세트 보유</span><span class="val">KTB3 +240 / KTB10 –80 (스티프너) + KTB10 –60 (네이키드) = KTB10 –140</span></div>
  <div class="row"><span class="key">KTB10 종가</span><span class="val">106.07 &nbsp;(Hi 106.82 / Lo 106.04 / O 106.75)</span></div>
  <div class="row"><span class="key">KTB3 종가</span><span class="val">102.93 &nbsp;(Hi 103.19 / Lo 102.93 / O 103.16)</span></div>
  <div class="row"><span class="key">3–10 스프레드 (종가)</span><span class="val">106.07 – 102.93 = <span style="color:#39c5cf">3.14pt</span> (전일 비교용 수기기입)</span></div>
</div>

<!-- P&L 노트 -->
<div class="pnl-note">
  <strong>손익 주의:</strong> 아래 표의 손익은 NH 체결 확인서 기준 (전일 정산가 대비 mark-to-market).
  합계 <strong>{total_pnl/10000:+.0f}만원</strong> (수수료 -66.9만원 별도).
  오버나잇 숏 포지션의 배경 MTM 손익(가격 하락분)은 별도 일일정산으로 계상됨.
</div>

<!-- 인터랙티브 차트 -->
<h2>차트 (KTB10 체결 + KTB3 캐리 — 드래그 확대 / 더블클릭 리셋)</h2>
<div class="chart-wrap">
{chart_div}
</div>

<!-- 체결 테이블 -->
<h2>체결 내역 (집계 35버킷 / 원본 49건)</h2>
<div class="table-wrap">
<table>
  <thead>
    <tr>
      <th>#</th><th>시각</th><th>매매</th><th>수량</th><th>체결가</th>
      <th>KTB10잔고</th><th>건별손익</th><th>누계손익</th>
    </tr>
  </thead>
  <tbody>
    <tr class="summary-row">
      <td colspan="2">전일이월</td>
      <td colspan="3" style="color:#8b949e">KTB3 +240 / KTB10 –140</td>
      <td style="color:#f85149;font-weight:700">–140</td>
      <td>—</td><td>—</td>
    </tr>
{table_rows}
    <tr class="summary-row">
      <td colspan="3"><strong>합계</strong></td>
      <td><span class="side-sell">S:{total_sell}</span> / <span class="side-buy">B:{total_buy}</span></td>
      <td>—</td>
      <td style="color:#f85149;font-weight:700">{final_pos:+d}</td>
      <td>—</td>
      <td>{fmt_pnl(total_pnl)}</td>
    </tr>
  </tbody>
</table>
</div>

<!-- 내 코멘트 -->
<h2>내 코멘트</h2>
<div class="comment-box">
  <label>오늘 매매 리뷰 (저장 전 입력)</label>
  <textarea id="commentArea" placeholder="진입 이유, 청산 타이밍, 아쉬운 점, 내일 전략 등..."></textarea>
</div>
<button class="btn-save" id="saveBtn">💾 코멘트 포함 저장 (.html)</button>
<span id="saveStatus"></span>
<p class="note">※ Ctrl+S 저장하면 레이아웃 깨짐 — 반드시 위 버튼 사용</p>

<script>
document.getElementById('saveBtn').addEventListener('click', function () {{
  var clone = document.documentElement.cloneNode(true);
  ['saveBtn','saveStatus'].forEach(function(id) {{
    var el = clone.querySelector('#' + id);
    if (el) el.remove();
  }});
  var ta = clone.querySelector('#commentArea');
  if (ta) {{
    var live = document.querySelector('#commentArea');
    ta.textContent = live.value;
    ta.setAttribute('readonly','readonly');
  }}
  var html = '<!DOCTYPE html>\\n' + clone.outerHTML;
  var blob = new Blob([html], {{type:'text/html;charset=utf-8'}});
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement('a');
  a.href = url; a.download = 'journal_20260701_저장본.html';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  document.getElementById('saveStatus').textContent = '저장 완료 ✓';
}});
</script>
</body>
</html>"""

with open(OUT_PATH, "w", encoding="utf-8") as f:
    f.write(html)

size_mb = len(html) / 1024 / 1024
print(f"저널 생성 완료: {OUT_PATH}")
print(f"파일크기: {size_mb:.1f} MB")
print(f"KTB10 체결: SELL {total_sell}계약 / BUY {total_buy}계약")
print(f"SELL avg {sell_avg:.3f} / BUY avg {buy_avg:.3f} / 스프레드 {sell_avg-buy_avg:+.3f}pt")
print(f"추정 마감 포지션: KTB10 {final_pos:+d} / KTB3 +240")
print(f"총 손익(체결기준): {total_pnl:+,}원")
