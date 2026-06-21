# KTB Bond Trading Toolkit

국채선물(KTB3 / KTB10) 트레이딩 자동화 도구 모음.  
Gmail 가정산보고서 자동 수집 → FIFO 손익 계산 → Excel 적재 → TradingView 저널링까지 원스톱 파이프라인.

---

## 구성 요소

```
workingwithC/
├── trading_journal/          ← 핵심 Python 패키지
│   ├── daily_run.py          ← 메인 실행 (Gmail → 파싱 → Excel)
│   ├── gmail_fetcher.py      ← Gmail OAuth 연동 / PDF 다운로드
│   ├── fifo_matcher.py       ← FIFO 진입-청산 매칭 / 라운드트립 계산
│   ├── excel_store.py        ← openpyxl 기반 Fills·Trades 시트 관리
│   ├── broker_registry.py    ← 증권사별 파서 레지스트리
│   ├── reauth.py             ← Gmail OAuth 재인증 (최초 1회 / 토큰 만료 시)
│   ├── run_journal.ps1       ← Task Scheduler용 래퍼 (--after 자동 계산)
│   ├── requirements.txt
│   ├── config/               ← credentials.json.json, token.json (gitignore)
│   └── parsers/
│       ├── aggregate.py      ← (분+종목+방향+가격) 합산 공통 로직
│       ├── nh_futures.py     ← NH선물 PDF 파서
│       └── samsung_futures.py← 삼성선물 PDF 파서
│
├── gmail_auth.ps1            ← PowerShell Gmail OAuth (초기 토큰 발급)
├── gmail_fetch_settlement.ps1← PowerShell PDF 일괄 다운로드 (settlements/)
├── trading_db_update.ps1     ← Excel COM 기반 P&L DB 업데이트
├── daily_pl_update.ps1       ← 오케스트레이터 (Gmail → 파싱 → DB 업데이트)
├── launch_tradingview.ps1    ← TradingView CDP 모드 실행
└── journal_YYYYMMDD.html     ← 데일리 저널 HTML 출력물 (템플릿)
```

---

## 셋업

### 1. Python 환경
```powershell
cd trading_journal
pip install -r requirements.txt
```

### 2. Gmail OAuth 인증 (최초 1회)
```powershell
cd trading_journal
python reauth.py
# 브라우저 자동 열림 → 계정 로그인 → 허용
# config/token.json 자동 생성
```

> `config/credentials.json.json` : Google Cloud Console에서 발급한 OAuth 클라이언트 JSON을 이 경로에 저장

### 3. 일일 실행
```powershell
# 당일 체결내역 처리
python daily_run.py

# 특정 날짜 이후 백필
python daily_run.py --after 2026/06/01

# 로컬 PDF 폴더로 테스트 (Gmail 없이)
python daily_run.py --dry-run
```

### 4. 자동화 (Task Scheduler)
`run_journal.ps1` 이 이미 두 가지 트리거로 등록돼 있음:
- **매일 17:00** 실행
- **로그온 시** 실행 (마지막 실행일 기준 미업데이트분 자동 소급)

---

## 데이터 흐름

```
Gmail (가정산보고서 PDF)
    ↓ gmail_fetcher.py
PDF 다운로드 (data/attachments/)
    ↓ parsers/nh_futures.py / samsung_futures.py
RawFill 리스트
    ↓ parsers/aggregate.py
AggregatedFill (분+종목+방향+가격 합산)
    ↓ fifo_matcher.py
RoundTrip (FIFO 진입-청산 매칭)
    ↓ excel_store.py
trading_journal.xlsx
  ├── Fills  시트 (체결내역 누적, 중복 방지)
  └── Trades 시트 (라운드트립 손익, 중복 방지)
```

---

## TradingView 저널링

`run_journal.ps1` 또는 `daily_run.py` 실행 후, TradingView MCP로:
1. 5분봉 차트 전환 (`KRX:BM31!` / `KRX:BMA1!`)
2. 체결 시각·가격 기준 진입(▼빨강)/청산(▲초록) 레이블 자동 드로잉
3. 스크린샷 저장
4. `journal_YYYYMMDD.html` 생성 (체결 테이블 + FIFO 라운드트립 + 분석)

---

## 증권사 파서 추가 방법

```python
# parsers/newbroker_futures.py 작성 후
# broker_registry.py 에 추가:
BrokerConfig(
    name="키움선물",
    subject_patterns=[r"\[키움선물\]"],
    gmail_query="from:kiwoom@kiwoom.com",
    parser=parse_kiwoom_futures,
)
```

---

## 주의사항

- `config/token.json` / `config/credentials.json.json` 은 `.gitignore` 처리됨 — 다른 환경에서는 `reauth.py` 재실행 필요
- Gmail refresh token 유효기간: 앱이 `Testing` 상태이면 **7일** (Production 전환 시 무제한)
- PDF 파서는 `pdfplumber` 사용 (`pdftotext` 불필요)
- FIFO 매칭은 Fills 시트 **전체** 기준 — 날짜 이월 포지션도 정확히 처리

---

## 관련 환경

| 항목 | 내용 |
|------|------|
| Python | 3.12+ |
| 주요 라이브러리 | pdfplumber, openpyxl, pandas, google-api-python-client |
| TradingView MCP | CDP 9222 포트 (launch_tradingview.ps1 참조) |
| 대상 종목 | KRX:BM31! (KTB3년물), KRX:BMA1! (KTB10년물) |
| 증권사 | 삼성선물 (master@ssfutures.com), NH선물 (NHfutures@futures.co.kr) |
