# Office PC 선물거래증적 파이프라인 (2026-06-30 기준)

`trading_journal/`(홈 PC 레퍼런스 구현, broker_registry.py 기반)과는 별개의 **병행 구현체**다. 둘 중 하나로 통일하지 않고 둘 다 유지하기로 함 — 이 폴더는 office PC에서 실제로 돌아가는 더 단순한 구조(파일별 단일 책임)의 파이프라인을 그대로 옮겨온 것.

## 구성
- `main.py` — Gmail→PDF→Excel 파이프라인 진입점. `python main.py 2026/06/30`처럼 날짜 인자로 실행.
- `gmail_fetcher.py` — Gmail에서 NH선물/SS선물 PDF 다운로드. NH는 제목에 `"국문가정산(체결시분)"`이 포함된 메일(01A103, 체결시간 포함)을 우선 쿼리하고, 그 메일이 그날 없으면 제목필터 없이 재시도하는 fallback이 있음(정정포함/02A101 같은 변형이라도 받아서 시간 없이 기록).
- `pdf_parser.py` — NH/SS PDF 텍스트 파싱. NH의 `[당일미결제]`(미결잔고) 섹션 파싱 버그(정의되지 않은 `code` 변수 참조로 매번 조용히 빈 리스트가 되던 문제)를 2026-06-30에 수정함.
- `excel_writer.py` — 파싱 결과를 월별 Excel(Sheet1: 전체내역 누적, Sheet2부터: 날짜별 증적)로 저장.
- `build_chart.py`, `build_chart_0630.py` — OHLCV를 `tv ohlcv`(tradingview-mcp CLI)로 받아 matplotlib으로 직접 렌더링하는 정적 차트(캔들+체결라벨+SMA/VWAP/RSI/VRVP). TV 스크린샷 방식의 대안.
- `build_chart_interactive.py` — 위와 같은 데이터를 Plotly로 렌더링한 인터랙티브 버전(줌/팬/호버). plotly.js를 파일에 전부 인라인 임베드해서 인터넷 연결 없이도, 파일 하나만 옮겨도 깨지지 않음.

## 알려진 한계 (자세한 내용은 ../TRADING_WORKFLOW_GUIDE.md 참고)
- `tv ohlcv`로 받을 수 있는 5분봉 히스토리가 그날그날 사전이력 봉 수에 따라 다름(어떤 날은 20봉, 어떤 날은 223봉) — SMA60/120/200이 일부만 그려지거나 TV 스냅샷 점선으로 대체될 수 있음.
- 자체계산 RSI(14)가 TV 차트 범례값과 8~11pt 정도 차이남(원인 미해결, VWAP/SMA는 거의 일치).
- 잔여계약 수치가 전일 오버나잇 잔량을 반영하지 못하는 경우가 있음(디버깅 보류 중, KTB3 +180 이상치와 같은 계열의 문제로 추정).

## 의도적으로 git에서 제외된 것 (`.gitignore` 기존 규칙)
- `credentials.json` / `token.json` (OAuth 비밀정보)
- `*.xlsx` (실제 거래 데이터)
- `*.png` (차트 스크린샷/이미지 — `journal_assets/`의 PNG들도 포함, 인터랙티브 차트(`*_interactive.html`)는 자체완결형이라 제외되지 않음)
- `다운로드/`, `결과/`, `실행로그.txt` 등 실행 시 생성되는 개인 데이터 (애초에 복사 안 함)

## 다음에 할 일 (메모리: `project_trading_journal_roadmap`)
1. 잔여계약 오버나잇 미반영 버그 — 추후 디버깅
2. 비고/구분 컬럼 추가 — 방향성매매(브레이크아웃/풀백)/커브매매/헷지매매/차익매매 5종 분류. 브레이크아웃·풀백은 OHLCV+SMA 기반 휴리스틱 자동분류 초안 가능, 나머지는 수기 입력 예정.
