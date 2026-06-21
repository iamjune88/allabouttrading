# KTB Bond Trading Toolkit

국채선물(KTB3 / KTB10) 트레이딩 전 과정 자동화.  
INFOMAX 시계열 데이터 자동수집 → 테크니컬/매크로 분석 → 거래증적(Gmail→PDF→FIFO→Excel) → TradingView 저널링까지 원스톱.

---

## 전체 구성

```
workingwithC/
│
├── 시계열/                       ← INFOMAX 연동 Excel (RTF 자동갱신)
│   └── refresh_excel.ps1         ← INFOMAX 연동 파일 일괄 갱신 스크립트
│
├── trading_journal/              ← 거래증적 Python 패키지
│   ├── daily_run.py              ← 메인 실행 (Gmail → 파싱 → Excel)
│   ├── gmail_fetcher.py          ← Gmail OAuth 2.0 / PDF 다운로드
│   ├── fifo_matcher.py           ← FIFO 진입-청산 매칭
│   ├── excel_store.py            ← Fills·Trades 시트 관리 (openpyxl)
│   ├── broker_registry.py        ← 증권사별 파서 레지스트리
│   ├── reauth.py                 ← Gmail OAuth 재인증
│   ├── run_journal.ps1           ← Task Scheduler 래퍼 (--after 자동 계산)
│   ├── requirements.txt
│   ├── config/                   ← credentials.json.json, token.json (gitignore)
│   └── parsers/
│       ├── aggregate.py          ← (분+종목+방향+가격) 합산 공통로직
│       ├── nh_futures.py         ← NH선물 가정산보고서 PDF 파서
│       └── samsung_futures.py    ← 삼성선물 가정산보고서 PDF 파서
│
├── launch_tradingview.ps1        ← TradingView CDP 모드 실행 (AppxPackage 방식)
├── trading_db_update.ps1         ← Excel COM 기반 일별 P&L DB 업데이트
├── daily_pl_update.ps1           ← 오케스트레이터 (Gmail → 파싱 → DB)
├── gmail_fetch_settlement.ps1    ← PowerShell PDF 일괄 다운로드
│
├── TRADING_WORKFLOW_GUIDE.md     ← TV MCP 조작 상세 가이드 (타임스탬프 계산 포함)
├── ktb_analysis_YYYYMMDD.html    ← 테크니컬 분석 HTML 리포트 템플릿
└── journal_YYYYMMDD.html         ← 데일리 거래 저널 HTML 출력물
```

---

## 1. INFOMAX 시계열 데이터 자동화

### 개요

NH증권 INFOMAX 단말에서 RTF(실시간 피드)로 연동된 4개 Excel 파일을 자동 갱신.  
`시계열/` 폴더 내 xlsx를 Excel COM으로 열면 INFOMAX가 자동으로 최신 데이터를 채워넣음.

**대상 파일 (4개):**
| 파일명 | 컬럼 수 | 주요 내용 |
|--------|---------|----------|
| ktb+daily_분석.xlsx | 188열 | KTB OHLCV + 투자자동향(외국인/기관/개인) |
| 미국금리시계열_daily.xlsx | 32열 | US 2/5/10/30Y + SOFR + TIPS 등 |
| 주요커브전략(국고)_daily_.xlsx | — | 국고 2/3/5/10/20/30Y 커브 + 스프레드 |
| 채권시가평가_daily.xlsx | — | 국채 시가평가 금리 일별 |

### 실행

```powershell
& "C:\Users\iamju\Desktop\workingwithC\시계열\refresh_excel.ps1"
```

**자동화 로직:**
1. 12시간 쿨다운 (중복 실행 방지)
2. INFOMAX 프로세스(`infomaxmain`) 실행 여부 확인 → 없으면 LNK 바로가기로 자동 실행 후 60초 대기
3. Excel COM으로 4개 파일 동시 열기 → 15초 대기 (RTF 갱신 시간) → 저장·닫기

**Task Scheduler 등록 (3개 트리거):**
| 트리거 | 시간 |
|--------|------|
| 매일 08:30 | 장 시작 전 |
| 매일 17:30 | 장 마감 후 |
| 로그온 시 | PC 재시작 후 소급 갱신 |

### INFOMAX 실행 경로
```
C:\ProgramData\Microsoft\Windows\Start Menu\Programs\인포맥스\인포맥스.lnk
```

---

## 2. 표준 분석 워크플로우 (5단계)

Claude Code + TradingView MCP 기반 채권 분석 표준 절차.

```
1단계: TradingView 연결
   launch_tradingview.ps1 → CDP 9222포트 → tv_health_check

2단계: INFOMAX Excel 읽기
   refresh_excel.ps1 실행 확인
   → ktb+daily_분석.xlsx: 최근 30일 OHLCV + 투자자동향
   → 미국금리시계열_daily.xlsx: US 커브 / TIPS 브레이크이븐

3단계: 기술적 분석
   chart_set_symbol(KRX:BM31! / KRX:BMA1!)
   → 일봉/주봉 MA(5/20/60/120), RSI, 볼린저밴드
   → 지지/저항 레벨 체크 (data_get_pine_lines)
   → 주요 수급동향 (외국인/기관 누적 매수도)

4단계: 매크로/뉴스 서치 (병렬 WebSearch)
   "Korea KTB bond [월 연도]"
   "Fed rate expectations FOMC [연도]"
   "BOJ JGB super long [연도]"
   "한국은행 기준금리 [연도]"

5단계: 통합 리포트
   ktb_analysis_YYYYMMDD.html 생성
   → lightweight-charts 기반 다크테마
   → 테크니컬/매크로/포지션/익일전략 4섹션
```

---

## 3. 매크로 분석 프레임워크

### KTB 분석 핵심 지표

| 지표 | 데이터 소스 | 의미 |
|------|-------------|------|
| US 10Y yield | 미국금리시계열_daily.xlsx | KTB10 방향성 최대 변수 |
| US 2/10Y 스프레드 | 동 파일 | 글로벌 리세션 선행지표 |
| USD/KRW | INFOMAX 실시간 | 외국인 KTB 매수/매도 촉매 |
| 외국인 KTB 동향 | ktb+daily_분석.xlsx | 수급 방향 |
| JGB 초장기 | WebSearch | BOJ 정책 변화 → KTB 장기물 연동 |
| FOMC dots | WebSearch | 금리 인상/인하 속도 선반영 |

### 5대 모니터링 시나리오 (매크로)
1. **US 금리 급등** (NFP 서프라이즈, FOMC 매파) → KTB 숏 유리
2. **BOJ 긴축 강화** (JGB 초장기 금리 급등) → 글로벌 장기물 동반 매도
3. **USD/KRW 1600 돌파** → 외국인 KTB 매도 압력
4. **한국 경기 부진** (수출·PMI 하락) → 한은 인하 기대 → KTB 강세 리스크
5. **지정학 리스크** (이스라엘, 대만 등) → 안전자산 선호 → KTB 강세 리스크

---

## 4. TradingView Pine Script 백테스팅

Claude Code MCP를 통해 TradingView 내 Pine Script 전략을 직접 개발·테스트.

### 워크플로우

```
1. pine_set_source → 전략 코드 주입
2. pine_smart_compile → 컴파일 + 오류 확인
3. pine_get_errors → 오류 목록 (있을 경우)
4. chart_set_timeframe → 백테스트 봉 단위 설정
5. data_get_strategy_results → 수익률, 승률, MDD, 샤프비율 등
6. capture_screenshot(region="strategy_tester") → 결과 시각화
```

### 주요 백테스팅 지표 (data_get_strategy_results 반환)

| 항목 | 필드명 |
|------|--------|
| 순수익 | `net_profit` |
| 승률 | `percent_profitable` |
| 최대낙폭 | `max_drawdown` |
| 손익비 | `profit_factor` |
| 총 트레이드 수 | `total_trades` |

### KTB 전략 예시 (Pine Script v5)

```pine
//@version=5
strategy("KTB MA Crossover", overlay=true, default_qty_type=strategy.fixed, default_qty_value=1)

fast = ta.ema(close, 5)
slow = ta.ema(close, 20)

if ta.crossover(fast, slow)
    strategy.entry("Long", strategy.long)
if ta.crossunder(fast, slow)
    strategy.close("Long")
    strategy.entry("Short", strategy.short)

plot(fast, color=color.blue, linewidth=1)
plot(slow, color=color.orange, linewidth=2)
```

---

## 5. 거래증적 자동화 파이프라인

### 전체 흐름

```
Gmail (가정산보고서 PDF)
    ↓ gmail_fetcher.py (OAuth 2.0)
PDF 다운로드 → data/attachments/
    ↓ parsers/nh_futures.py 또는 samsung_futures.py (pdfplumber)
RawFill 리스트
    ↓ parsers/aggregate.py (분+종목+방향+가격 합산)
AggregatedFill
    ↓ fifo_matcher.py (선입선출, 이월 포지션 처리)
RoundTrip (진입-청산 쌍)
    ↓ excel_store.py (openpyxl)
trading_journal.xlsx
  ├── Fills 시트  (체결내역 누적, 중복 방지)
  └── Trades 시트 (라운드트립 손익, 중복 방지)
```

### Gmail OAuth 인증

```powershell
# 최초 1회 또는 토큰 만료 시 (7일 - Testing 앱 한도)
cd trading_journal
python reauth.py
# 브라우저 자동 열림 → Google 계정 로그인 → 허용
# config/token.json 자동 저장
```

> `config/credentials.json.json`: Google Cloud Console OAuth 클라이언트 JSON  
> 파일명 그대로 두어야 함 (`.json.json` 이중 확장자)

### 실행 방법

```powershell
cd trading_journal

# 당일 처리
python daily_run.py

# 특정 날짜 이후 백필
python daily_run.py --after 2026/06/15

# 로컬 PDF 테스트 (Gmail 없이)
python daily_run.py --dry-run
```

### Task Scheduler 자동화

| 트리거 | 동작 |
|--------|------|
| 매일 17:00 | 당일 가정산보고서 처리 |
| 로그온 시 | 마지막 실행 이후 미처리분 소급 (last_run.txt 기준) |

스크립트: `trading_journal/run_journal.ps1`

### 지원 증권사

| 증권사 | 발신 주소 | 파서 |
|--------|-----------|------|
| NH선물 | NHfutures@futures.co.kr | parsers/nh_futures.py |
| 삼성선물 | master@ssfutures.com | parsers/samsung_futures.py |

### 증권사 파서 추가

```python
# parsers/newbroker_futures.py 작성 후 broker_registry.py에 추가:
BrokerConfig(
    name="키움선물",
    subject_patterns=[r"\[키움선물\]"],
    gmail_query="from:kiwoom@kiwoom.com",
    parser=parse_kiwoom_futures,
)
```

---

## 6. TradingView 거래 라벨링 + 데일리 저널

### 거래 라벨 규칙

| 구분 | 위치 | 색상 | 예시 |
|------|------|------|------|
| SELL(진입/추가) | 체결가 + 0.07pt 위 | #FF4444 | `▼ 09:15  SELL 50 @ 103.62` |
| BUY(커버/청산) | 체결가 - 0.07pt 아래 | #44BB44 | `▲ 14:22  BUY(커버) 30 @ 103.38` |

같은 봉에 2개 이상 겹칠 경우 0.06pt 이상 추가 간격.

### 타임스탬프 계산 (KST → Unix UTC)

```
UTC timestamp = (KST 시각 - 9시간)의 Unix 초

2026-06-15 09:15 KST
= 2026-06-15 00:15 UTC
= 1781481600 + 15×60 = 1781482500
```

### 데일리 저널 생성 절차

```
1. trading_journal.xlsx에서 당일 체결내역 + FIFO 라운드트립 조회
2. chart_set_symbol(BM31!) + chart_set_timeframe("5")
3. draw_clear → draw_shape × 각 체결건 라벨
4. capture_screenshot → ktb3_YYYYMMDD_trades.png
5. KTB10(BMA1!) 반복
6. journal_YYYYMMDD.html 생성
   - 체결내역 테이블 (KTB3 / KTB10)
   - FIFO 라운드트립 전체 (익절=green / 손절=red)
   - 4섹션 기술적 분석 그리드
   - 익일 전략 3-column verdict
```

---

## 7. TradingView 심볼 매핑

| 종목 | TradingView 심볼 | 가정산보고서 코드 |
|------|-----------------|----------------|
| KTB 3년물 | `KRX:BM31!` | A6566 |
| KTB 10년물 | `KRX:BMA1!` | A6766 |
| KTB 30년물 | `KRX:BML1!` | A6966 |
| US 2yr (ZT) | `ZT1!` | — |
| US 10yr (ZN) | `ZN1!` | — |
| German Bund | `FGBL1!` | — |

---

## 8. 셋업 (다른 환경에서 재현)

### 1) Python 환경

```powershell
cd trading_journal
pip install -r requirements.txt
```

**requirements.txt:**
```
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
openpyxl
pandas
pdfplumber
```

### 2) TradingView CDP 모드 실행

```powershell
& "C:\Users\iamju\Desktop\workingwithC\launch_tradingview.ps1"
```

> TV는 Microsoft Store(AppxPackage) 설치. 일반 경로 탐색 실패하므로 반드시 스크립트 사용.  
> 상세: `TRADING_WORKFLOW_GUIDE.md` 참조

### 3) Gmail OAuth 초기 인증

```powershell
cd trading_journal
python reauth.py
```

### 4) INFOMAX Excel 자동화 등록 (Task Scheduler)

```powershell
# 08:30, 17:30, 로그온 트리거 등록 (최초 1회)
# trading_journal/run_journal.ps1 과 동일한 방식으로 schtasks /create 사용
```

---

## 주의사항

- `config/token.json` / `config/credentials.json.json` 은 gitignore — 다른 환경에서 `reauth.py` 재실행 필요
- Gmail refresh token 유효기간: Testing 상태이면 **7일**, Production 전환 시 무제한
- PDF 파서는 `pdfplumber` 사용 (pdftotext 불필요, Windows 지원)
- FIFO 매칭은 Fills 시트 **전체** 기준 — 이월 포지션도 정확히 처리
- INFOMAX 연동 Excel은 `ReadOnly=False`로 열어야 RTF 갱신됨

---

## 관련 파일

| 파일 | 설명 |
|------|------|
| `TRADING_WORKFLOW_GUIDE.md` | TV MCP 조작 상세가이드 (타임스탬프 계산, draw_shape 포맷, 지표 체크리스트) |
| `trading_journal/README.md` | Python 패키지 단독 셋업 가이드 |
| `ktb_analysis_YYYYMMDD.html` | 테크니컬 분석 HTML 리포트 (lightweight-charts 다크테마) |
| `journal_YYYYMMDD.html` | 데일리 거래 저널 (체결+FIFO+분석+전략) |

---

## 환경 정보

| 항목 | 내용 |
|------|------|
| OS | Windows 11 |
| Python | 3.12+ |
| TradingView | Microsoft Store (AppxPackage, CDP 9222) |
| INFOMAX | NH증권 단말 (infomaxmain.exe) |
| Gmail | OAuth 2.0 (Testing 앱, 7일 만료) |
| 증권사 | 삼성선물, NH선물 |
| 대상 종목 | KRX:BM31! (KTB3), KRX:BMA1! (KTB10) |
