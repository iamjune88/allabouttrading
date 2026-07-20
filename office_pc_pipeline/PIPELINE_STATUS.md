# KTB 선물 데일리 리뷰 파이프라인 — 현황 및 이슈 정리

작성일: 2026-07-20  
작성: Claude Sonnet 4.6 (claude-sonnet-4-6)

---

## 1. 파이프라인 개요

### 목적
KTB3년물(BM31!) / KTB10년물(BMA1!) 선물 거래의 일일 리뷰 자동화.

### 세 가지 핵심 작업
1. **Excel 거래내역 증적** — NH선물 + SS선물(삼성선물) 체결내역 → `선물거래_2026-07.xlsx`
2. **TradingView 시장데이터** — CDP(Chrome DevTools Protocol) 통해 5분봉 OHLCV 취득
3. **HTML 리뷰 생성** — 오버나잇 포지션, 갱신차금, 건별손익, 차트 포함

### 실행 순서
```
python main.py 2026/07/16       # Gmail에서 PDF 다운로드 → 파싱 → Excel 증적 → positions.json
python daily_review.py 2026/07/16  # Excel + TV OHLCV → HTML 생성 → git push
```

---

## 2. 파일 구조

| 파일 | 역할 |
|---|---|
| `main.py` | Gmail API로 NH/SS선물 PDF 수신, 파싱, Excel 기록, positions.json 저장 |
| `pdf_parser.py` | NH선물/SS선물 PDF 텍스트 파싱 (체결내역, 미결제약정) |
| `daily_review.py` | OVN 로드, TV OHLCV 취득, 갱신차금 계산, HTML/차트 생성, git push |
| `positions.json` | 전일 OVN 및 정산가 임시 저장 (문제의 근원 — 아래 참조) |
| `excel_writer.py` | openpyxl 기반 Excel 기록 |
| `gmail_fetcher.py` | Gmail API 인증 및 메일 수신 |

---

## 3. 종목 매핑

| 거래소 코드 | 종목 | TV 심볼 | 승수 |
|---|---|---|---|
| A65XX | KTB3년물 선물 | BM31! | 1,000,000원/pt |
| A67XX | KTB10년물 선물 | BMA1! | 1,000,000원/pt |

### 세션 시간
- SESSION_START: 08:45 KST (DAY_BASE - 900)
- SESSION_END: 15:45 KST (DAY_BASE + 24300)
- 야간선물 15:45봉 제외 (전일 야간 마지막 봉 혼입 방지)

---

## 4. SS선물 미결제약정 PDF 컬럼 매핑

```
종목    구분  잔고  전일잔고  전일정산가  당일정산가  갱신차금  옵션가치
A6569  매수  160    240     103.10     102.81    -43,000,000   0
A6769  매도   20     30     105.90     105.18    +21,600,000   0
```

- `잔고` = 오늘 마감 OVN
- `전일잔고` = 어제 마감 = 오늘 진입 OVN
- `갱신차금` = 실제 MTM 손익 (PDF 값이 가장 정확)

---

## 5. OVN 추적 로직 (현재 구조 — 문제 있음)

### 현재 방식
1. `main.py` → SS선물 PDF 파싱 → `positions.json`에 OVN/정산가 저장
2. `daily_review.py` → `positions.json` 읽어서 전일 OVN 로드
3. 당일 체결 net 더해서 당일 마감 OVN 계산
4. `positions.json` merge 저장 (다음 날 입력값)

### 문제점 (한 달간 반복 발생)
- **체인 의존성**: positions.json이 하루라도 잘못 쓰이면 이후 날짜 전부 오염
- **과거 날짜 재실행 시 체인 역전**: main.py 재실행이 당일 데이터 덮어씀
- **SS선물 거래 없는 날 PDF 없음**: 갱신차금 자동계산 불가
- **수동 개입 금지 규칙과 충돌**: PDF에 있는 데이터인데 수동 입력 요구 → 유저 분노
- **merge 저장 버그**: daily_review가 positions.json 덮어써서 main.py 데이터 소실 (한 세션에서 발생)

### 올바른 방식 (미구현)
- OVN = **Excel 누적합** (기준일 + Σ체결) — positions.json 체인 불필요
- PDF 미결제약정과 **크로스체크** (Excel 증적 누락 가능성 대비)
- 불일치 시 경고 출력, PDF 우선 사용
- 갱신차금 = 총OVN × (당일정산가 - 전일정산가) × 1,000,000 (브로커 무관)

---

## 6. 갱신차금 계산

### PDF 정산가 우선
```python
# positions.json에 ktb3_ovn_pnl 있으면 PDF 값 사용
ovn_pnl_ktb3 = pos_data["ktb3_ovn_pnl"]
```

### OHLCV 근사 (PDF 없을 때)
```python
ovn_pnl_ktb3 = OVN_KTB3 * (today_close - prev_close) * 1_000_000
```

---

## 7. TradingView CDP 이슈

### 알려진 문제
1. **chart_ready 이벤트 신뢰 불가**: 심볼 이름 업데이트 시 발생, 봉 데이터 로드 완료 아님
2. **과거 날짜 봉 데이터 로드 실패**: TV가 현재 날짜 봉만 안정적으로 제공. 과거 날짜는 뷰포트 range 명령 후에도 다음날 봉 반환하는 경우 있음
3. **BMA1!/BM31! 혼동**: TV가 이전 심볼 봉 그대로 반환. 가격 절대값 검증으로 대응 (BMA1은 항상 BM31보다 ~2-3pt 높음)

### 현재 방어 로직
- `price_near` 파라미터: 전일 정산가 ±1.5pt 범위 이탈 시 재시도 (최대 4회)
- `session_bars` 체크: 오늘 세션 봉 없으면 재시도
- `process()` fallback: 세션 봉 없으면 최근 봉으로 대체, bars 완전 빈 경우 dummy 봉

### 과거 날짜 재실행 시 process() 크래시 (2026-07-20 수정)
TV가 다음날 봉 반환 → `SESSION_END+300` 필터 후 bars 빈 리스트 → max() 크래시.
수정: filtered 빈 경우 원본 bars 사용.

---

## 8. 실행 이력 요약 (2026-07월)

| 날짜 | 상태 | OVN(시작) | OVN(종료) | 건별손익 | 갱신차금 | 총손익 |
|---|---|---|---|---|---|---|
| 07/13 | 완료 | +90/-30 | +180/-60 | +3,810만 | -900만 | +2,910만 |
| 07/14 | 완료 | +180/-60 | +240/-30 | -190만 | -660만 | -850만 |
| 07/15 | 완료 | +240/-30 | +40/+20 | +910만 | +1,290만 | +2,200만 |
| 07/16 | **오류** | 160/-20(잘못됨) | — | -570만 | +780만(추정) | +210만(추정) |

### 07/16 오류 내용
- 07/15 daily_review가 ending OVN +40/+20 올바르게 계산
- Claude가 임의로 positions.json을 160/-20으로 덮어씀 (오해)
- 07/16 journal이 OVN 160/-20으로 잘못 기록됨
- **정확한 07/16 OVN**: NH선물 +100/0, SS선물 -60/+20 → 합계 **+40/+20**
- 07/16 재실행 필요

---

## 9. 미해결 과제

1. **07/16 journal 재실행** — OVN +40/+20으로 다시
2. **OVN 추적 리팩토링** — positions.json 체인 → Excel 누적합 + PDF 크로스체크
3. **SS선물 거래 없는 날 갱신차금** — 전일 SS선물 OVN 캐리 × OHLCV 근사
4. **과거 날짜 재실행 프로토콜** — 날짜 지정 재실행 시 체인 역전 방지

---

## 10. 환경

- OS: Windows 11 Pro
- Python: Miniconda3
- TradingView: Microsoft Store (AppxPackage) — 일반 경로 실행 불가
- CDP 포트: 9222
- Gmail API: `credentials.json` / `token.json`
- Git repo: `https://github.com/iamjune88/allabouttrading.git`
- Excel: `결과/선물거래_2026-07.xlsx`
- positions.json: `선물거래증적/positions.json`

---

*이 문서는 Fable 모델 전환 전 현황 파악용으로 작성됨.*
