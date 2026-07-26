# Office PC 선물거래증적 파이프라인 (2026-06-30 기준)

`trading_journal/`(홈 PC 레퍼런스 구현, broker_registry.py 기반)과는 별개의 **병행 구현체**다. 둘 중 하나로 통일하지 않고 둘 다 유지하기로 함 — 이 폴더는 office PC에서 실제로 돌아가는 더 단순한 구조(파일별 단일 책임)의 파이프라인을 그대로 옮겨온 것.

## 구성
- `main.py` — Gmail→파싱·병합→Excel + OVN 스냅샷 파이프라인 진입점. `python main.py 2026/07/23`처럼 날짜 인자로 실행.
- `gmail_fetcher.py` — Gmail에서 자료 다운로드. **발신자+제목으로 각 자료를 유일하게 특정**(추측 fallback·dedup 없음). 필요 세트: SS 미결(PDF), SS 체결(CSV), NH 체결+미결(CSV). 지정 메일이 없으면 조용히 넘기지 않고 경고.
- `parse_router.py` — 다운로드분을 파싱하고 (종목사,날짜)별로 하나로 병합. 체결은 CSV 우선(틱), 미결은 있는 소스에서.
- `pdf_parser.py` — NH/SS PDF·CSV 파싱. **SS `[미결제약정]` 컬럼을 실제 양식(종목 구분 잔고 전일잔고 전일정산가 당일정산가 갱신차금)으로 수정** — 이전엔 컬럼이 틀려 잔여계약이 오류였음. NH/SS CSV 파서(`parse_nh_csv`/`parse_ss_csv`) 추가.
- `ovn_tracker.py` — **오버나잇(OVN) = Excel 전체내역 누적합**으로 결정론적 계산. anchor(기준일)+Σ순매수. 월경계까지 순회. positions.json 체인 대체.
- `ovn_crosscheck.py` — Excel 누적 OVN ↔ 브로커 확정 미결제약정(전일잔고) 교차검증. 불일치 시 daily_review 중단. 갱신차금은 PDF값 우선, 정산가차 계산으로 검산.
- `ovn_anchor.json` — OVN 계산 기준일(신뢰 시작일+그날 진입 OVN). `ovn_anchor.json.example` 복사해서 사용(실파일 gitignore).
- `excel_writer.py` — 파싱 결과를 월별 Excel(Sheet1: 전체내역 누적, Sheet2부터: 날짜별 증적)로 저장. 미결잔고 헤더를 신 스키마로 갱신.
- `daily_review.py` — OVN을 Excel에서 자동 계산 + 크로스체크 → TV OHLCV → HTML/차트 → git push. `--ktb3/--ktb10`으로 수동 확정 가능(크로스체크 우회).
- `build_chart.py`, `build_chart_0630.py` — OHLCV를 `tv ohlcv`(tradingview-mcp CLI)로 받아 matplotlib으로 직접 렌더링하는 정적 차트(캔들+체결라벨+SMA/VWAP/RSI/VRVP). TV 스크린샷 방식의 대안.
- `build_chart_interactive.py` — 위와 같은 데이터를 Plotly로 렌더링한 인터랙티브 버전(줌/팬/호버). plotly.js를 파일에 전부 인라인 임베드해서 인터넷 연결 없이도, 파일 하나만 옮겨도 깨지지 않음.

## 알려진 한계 (자세한 내용은 ../TRADING_WORKFLOW_GUIDE.md 참고)
- `tv ohlcv`로 받을 수 있는 5분봉 히스토리가 그날그날 사전이력 봉 수에 따라 다름(어떤 날은 20봉, 어떤 날은 223봉) — SMA60/120/200이 일부만 그려지거나 TV 스냅샷 점선으로 대체될 수 있음.
- 자체계산 RSI(14)가 TV 차트 범례값과 8~11pt 정도 차이남(원인 미해결, VWAP/SMA는 거의 일치).
- ~~잔여계약 수치가 전일 오버나잇 잔량을 반영하지 못하는 경우~~ → **2026-07-26 해결**: SS 미결 컬럼 수정 + OVN을 Excel 누적으로 결정론적 계산 + 브로커 미결제약정 크로스체크.
- OVN 누적은 anchor월~대상월 Excel 파일을 순회한다. **anchor는 신뢰할 수 있는 최근일로 주기적으로 전진**시키는 게 안전(오래된 anchor일수록 읽어야 할 월 파일이 많아짐).

## 의도적으로 git에서 제외된 것 (`.gitignore`)
- `office_pc_pipeline/credentials.json` / `token.json` (OAuth 비밀정보 — 2026-07-26 명시적 규칙 추가, 이전엔 gmail_* 만 커버되어 누락 위험이었음)
- `*.xlsx` (실제 거래 데이터 — 전역 규칙)
- `office_pc_pipeline/다운로드/`, `결과/`, `ovn_snapshots/`, `ovn_anchor.json`, `positions.json`, `실행로그.txt`, `_ohlcv_*.json`
- `*.png` (차트 이미지 — 인터랙티브 차트 `*_interactive.html`는 자체완결형이라 제외 안 됨)

## 다음에 할 일
1. 실 데이터로 07/23 이후 정기운영 검증(이 리팩터는 합성데이터 전 로직 통과 + 07/15·16·21·23 실물로 Fable이 확인).
2. 비고/구분 컬럼 추가 — 방향성매매(브레이크아웃/풀백)/커브매매/헷지매매/차익매매 5종 분류. 브레이크아웃·풀백은 OHLCV+SMA 기반 휴리스틱 자동분류 초안 가능, 나머지는 수기 입력 예정.
